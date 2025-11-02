import os
import time
import psutil
import logging
from datetime import datetime
from typing import Dict, Optional, Tuple, Any, Set, List, cast

from googleapiclient.discovery import build  # type: ignore[import-untyped]
from googleapiclient.errors import HttpError  # type: ignore[import-untyped]
from google.oauth2.service_account import Credentials  # type: ignore[import-untyped]
import sys
import subprocess
import re
from pathlib import Path

# ---- Configuration via environment variables ----
POLL_INTERVAL_SECONDS: int = int(os.getenv("POLL_INTERVAL_SECONDS", "5"))
GOOGLE_SHEET_ID: Optional[str] = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SERVICE_ACCOUNT_FILE: Optional[str] = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")

# Optional: semicolon-separated pairs "exe[=Display Name]". Example:
# GAMES="eldenring.exe=ELDEN RING;witcher3.exe=The Witcher 3;apex.exe"
GAMES_ENV: str = os.getenv("GAMES", "")

# Default games (only used if GAMES env not supplied)
DEFAULT_GAMES: Dict[str, str] = {
    "eldenring.exe": "ELDEN RING",
    "witcher3.exe": "The Witcher 3",
    "apex.exe": "Apex Legends",
    "fortniteclient-win64-shipping.exe": "Fortnite",
    "league of legends.exe": "League of Legends",
}

SCOPES: List[str] = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_TITLE: str = "Gaming"

# Task name for Windows Task Scheduler
TASK_NAME: str = "GamingLogWatcher"

# Steam-based auto detection config (no whitelist needed)
EXCLUDED_EXE_NAMES: Set[str] = {
    "steam.exe",
    "steamservice.exe",
    "steamwebhelper.exe",
    "steamerrorreporter.exe",
}


def parse_games_env(games_env: str) -> Dict[str, str]:
    """
    Parse GAMES env var into a dict of { exe_name_lower: display_name }.
    Accepts "exe" or "exe=Display Name" entries separated by semicolons.
    """
    mapping: Dict[str, str] = {}
    for entry in [e.strip() for e in games_env.split(";") if e.strip()]:
        if "=" in entry:
            exe, name = entry.split("=", 1)
            exe = exe.strip().lower()
            name = name.strip()
            if exe:
                mapping[exe] = name or exe
        else:
            exe = entry.strip().lower()
            if exe:
                # Use the exe (without extension) capitalized as a display fallback
                display: str = os.path.splitext(os.path.basename(exe))[0]
                mapping[exe] = display
    return mapping


def build_game_map() -> Dict[str, str]:
    user_map: Dict[str, str] = parse_games_env(GAMES_ENV)
    if user_map:
        return user_map
    # Fallback to defaults if user did not configure
    return {k.lower(): v for k, v in DEFAULT_GAMES.items()}

# New: resolve Steam libraries and detect games by path
def _parse_libraryfolders_vdf(vdf_path: Path) -> Set[str]:
    libs: Set[str] = set()
    try:
        text: str = vdf_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return libs
    # Support both old and new formats
    # Old: "0"  "C:\\Path"
    libs.update([m.group(1) for m in re.finditer(r'"\d+"\s*"([^"]+)"', text, flags=re.IGNORECASE)])
    # New: "path" "C:\\Path"
    libs.update([m.group(1) for m in re.finditer(r'"path"\s*"([^"]+)"', text, flags=re.IGNORECASE)])
    return {os.path.normpath(p) for p in libs if p}

def get_steam_library_common_dirs() -> List[str]:
    # Only use the default Steam path; do not search registry or env
    default_base: str = r"C:\Program Files (x86)\Steam"
    libs: Set[str] = set()

    if os.path.isdir(default_base):
        libs.add(os.path.normpath(default_base))
        vdf: Path = Path(default_base) / "steamapps" / "libraryfolders.vdf"
        libs.update(_parse_libraryfolders_vdf(vdf))

    # Convert to steamapps/common directories (only those that exist)
    common_dirs: List[str] = []
    for lib in libs:
        common_dir: Path = Path(lib) / "steamapps" / "common"
        try:
            if common_dir.is_dir():
                common_dirs.append(os.path.normcase(str(common_dir.resolve())))
        except Exception:
            common_dirs.append(os.path.normcase(os.path.normpath(str(common_dir))))
    # Deduplicate while preserving order
    dedup: List[str] = []
    seen: Set[str] = set()
    for d in common_dirs:
        if d not in seen:
            seen.add(d)
            dedup.append(d)
    return dedup

def _nice_title(name: str) -> str:
    s: str = re.sub(r"[_\-]+", " ", name).strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s.title() if s else name

def derive_game_name_from_path(exe_path: str, library_common_dirs: List[str]) -> str:
    exe: Path = Path(exe_path)
    exe_lower: str = os.path.normcase(str(exe))
    for lib in library_common_dirs:
        lib_lower: str = os.path.normcase(lib)
        if exe_lower.startswith(lib_lower):
            try:
                rel: str = os.path.relpath(exe_lower, lib_lower)
            except Exception:
                rel = exe.name
            parts: List[str] = [p for p in Path(rel).parts if p not in (os.sep, "/", "")]
            if parts:
                return _nice_title(parts[0])
    # Fallback: parent folder
    try:
        return _nice_title(exe.parent.name or exe.stem)
    except Exception:
        return exe.stem

def detect_running_games_steam(library_common_dirs: List[str]) -> Dict[int, Tuple[str, str]]:
    matches: Dict[int, Tuple[str, str]] = {}
    for p in psutil.process_iter(attrs=[], ad_value=None):  # type: ignore[call-arg]
        try:
            exe_path: str = p.exe() or ""
            if not exe_path:
                continue
            base: str = os.path.basename(exe_path).lower()
            if base in EXCLUDED_EXE_NAMES:
                continue
            exe_norm: str = os.path.normcase(os.path.normpath(exe_path))
            if any(exe_norm.startswith(lib) for lib in library_common_dirs):
                display: str = derive_game_name_from_path(exe_path, library_common_dirs)
                matches[p.pid] = (base, display)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception:
            continue
    return matches

# Helpers to install/uninstall a Windows Scheduled Task
def _pythonw_executable() -> str:
    py: Path = Path(sys.executable)
    cand: Path = py.with_name("pythonw.exe")
    return str(cand if cand.exists() else py)

def _script_path() -> str:
    if getattr(sys, "frozen", False):
        return sys.executable
    return str(Path(__file__).resolve())

def _build_run_command() -> str:
    workdir: Path = Path(_script_path()).parent
    if getattr(sys, "frozen", False):
        inner: str = f'"{_script_path()}" --run'
    else:
        inner = f'"{_pythonw_executable()}" "{_script_path()}" --run'
    return f'cmd.exe /c "cd /d {workdir} && {inner}"'

def install_task(on_startup: bool = False, as_system: bool = False) -> None:
    cmd: str = _build_run_command()
    args: List[str] = ["schtasks", "/Create", "/F", "/TN", TASK_NAME, "/RL", "HIGHEST"]
    if on_startup:
        args += ["/SC", "ONSTART"]
        if as_system:
            args += ["/RU", "SYSTEM"]
    else:
        args += ["/SC", "ONLOGON"]
    args += ["/TR", cmd]
    logging.info("Creating scheduled task '%s' (%s)...", TASK_NAME, "startup" if on_startup else "logon")
    subprocess.run(args, check=True)
    logging.info("Scheduled task created.")

def uninstall_task() -> None:
    logging.info("Deleting scheduled task '%s' if it exists...", TASK_NAME)
    subprocess.run(["schtasks", "/Delete", "/TN", TASK_NAME, "/F"], check=False)
    logging.info("Scheduled task removed (if it existed).")

def task_exists() -> bool:
    try:
        completed: subprocess.CompletedProcess[bytes] = subprocess.run(
            ["schtasks", "/Query", "/TN", TASK_NAME], 
            capture_output=True
        )
        return completed.returncode == 0
    except Exception:
        return False

def start_task() -> None:
    logging.info("Starting scheduled task '%s'...", TASK_NAME)
    subprocess.run(["schtasks", "/Run", "/TN", TASK_NAME], check=False)


class Session:
    def __init__(self, pid: int, exe_name: str, display_name: str, start_time: datetime) -> None:
        self.pid: int = pid
        self.exe_name: str = exe_name
        self.display_name: str = display_name
        self.start_time: datetime = start_time

    def finalize(self, end_time: datetime) -> Dict[str, str]:
        duration_seconds: int = max(0, int((end_time - self.start_time).total_seconds()))
        duration_minutes: float = round(duration_seconds / 60.0, 2)
        return {
            "game": self.display_name,
            "exe": self.exe_name,
            "start_iso": self.start_time.isoformat(timespec="seconds"),
            "end_iso": end_time.isoformat(timespec="seconds"),
            "duration_minutes": str(duration_minutes),
        }


def get_sheets_service() -> Optional[Any]:
    if not GOOGLE_SERVICE_ACCOUNT_FILE or not os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE):
        logging.error("GOOGLE_SERVICE_ACCOUNT_FILE not set or file not found.")
        return None
    try:
        creds: Credentials = Credentials.from_service_account_file(  # type: ignore[attr-defined]
            GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        return cast(Any, build("sheets", "v4", credentials=creds, cache_discovery=False))
    except Exception as e:
        logging.exception("Failed to initialize Google Sheets service: %s", e)
        return None


def ensure_sheet_exists(service: Any, spreadsheet_id: str, title: str) -> None:
    # ...existing code...
    try:
        meta: Dict[str, Any] = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets: List[Dict[str, Any]] = meta.get("sheets", [])
        for s in sheets:
            props: Dict[str, Any] = s.get("properties", {})
            if props.get("title") == title:
                return
        # Create the sheet if not present
        batch_req: Dict[str, Any] = {
            "requests": [
                {"addSheet": {"properties": {"title": title}}}
            ]
        }
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_req).execute()
        # Optionally add a header row if empty
        try:
            headers: List[List[str]] = [["Ended At", "Game", "Started At", "Ended At (dup)", "Duration (min)"]]
            service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{title}!A1:E1",
                valueInputOption="USER_ENTERED",
                body={"values": headers},
            ).execute()
        except Exception:
            pass
    except HttpError as e:  # type: ignore[misc]
        logging.exception("Failed to ensure sheet exists: %s", e)
        raise


def append_session(service: Any, spreadsheet_id: str, title: str, payload: Dict[str, str]) -> None:
    """
    Append a row to the sheet with session details.
    Columns: Ended At, Game, Started At, Ended At (dup), Duration (min).
    """
    row: List[str] = [
        payload["end_iso"],
        payload["game"],
        payload["start_iso"],
        payload["end_iso"],
        payload["duration_minutes"],
    ]
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{title}!A:E",
        valueInputOption="USER_ENTERED",
        body={"values": [row]},
    ).execute()


def get_process_identity(p: psutil.Process) -> Tuple[str, str]:
    # ...existing code...
    name: str = ""
    try:
        name = (p.name() or "").lower()
    except Exception:
        pass
    base: str = name
    try:
        exe_path: str = p.exe() or ""
        if exe_path:
            base = os.path.basename(exe_path).lower()
    except Exception:
        pass
    return name, base


def detect_running_games(game_map: Dict[str, str]) -> Dict[int, Tuple[str, str]]:
    # ...existing code...
    matches: Dict[int, Tuple[str, str]] = {}
    for p in psutil.process_iter(attrs=[], ad_value=None):  # type: ignore[call-arg]
        try:
            name, base = get_process_identity(p)
            for key in (name, base):
                if key in game_map:
                    matches[p.pid] = (key, game_map[key])
                    break
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception:
            continue
    return matches

def run_watcher() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    if not GOOGLE_SHEET_ID:
        logging.error("GOOGLE_SHEET_ID not set.")
        return

    service: Optional[Any] = get_sheets_service()
    if not service:
        return

    # Ensure target sheet exists
    ensure_sheet_exists(service, GOOGLE_SHEET_ID, SHEET_TITLE)

    # Prefer Steam auto-detection using the default install path
    library_common_dirs: List[str] = get_steam_library_common_dirs()
    use_steam: bool = bool(library_common_dirs)
    game_map: Dict[str, str] = {}
    
    if use_steam:
        logging.info("Auto-detecting games under Steam libraries:")
        for d in library_common_dirs:
            logging.info("  %s", d)
    else:
        game_map = build_game_map()
        if not game_map:
            logging.error("No Steam libraries found at the default path and no games configured. Install Steam in the default path or set GAMES env.")
            return
        logging.info("Watching for games (allowlist fallback): %s", ", ".join(sorted(set(game_map.values()))))

    active_sessions: Dict[int, Session] = {}

    while True:
        try:
            # choose detection strategy
            running: Dict[int, Tuple[str, str]]
            if use_steam:
                running = detect_running_games_steam(library_common_dirs)
            else:
                running = detect_running_games(game_map)

            # Start sessions for newly seen PIDs
            for pid, (exe_name, display_name) in running.items():
                if pid not in active_sessions:
                    active_sessions[pid] = Session(
                        pid=pid,
                        exe_name=exe_name,
                        display_name=display_name,
                        start_time=datetime.now(),
                    )
                    logging.info("Started session: %s (pid=%d)", display_name, pid)

            # Finalize sessions for PIDs no longer present
            ended_pids: List[int] = [pid for pid in active_sessions.keys() if pid not in running]
            for pid in ended_pids:
                session: Session = active_sessions.pop(pid)
                end_time: datetime = datetime.now()
                payload: Dict[str, str] = session.finalize(end_time)
                try:
                    append_session(service, GOOGLE_SHEET_ID, SHEET_TITLE, payload)
                    logging.info(
                        "Logged session: %s | %s -> %s | %s min",
                        payload["game"],
                        payload["start_iso"],
                        payload["end_iso"],
                        payload["duration_minutes"],
                    )
                except HttpError as e:  # type: ignore[misc]
                    logging.exception("Failed to append to Google Sheet: %s", e)

            time.sleep(POLL_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            logging.info("Exiting watcher.")
            break
        except Exception as e:
            logging.exception("Unexpected error in loop: %s", e)
            time.sleep(POLL_INTERVAL_SECONDS)

def main() -> None:
    import argparse
    parser: argparse.ArgumentParser = argparse.ArgumentParser(prog="GamingLog", description="Game session logger")
    parser.add_argument("--install-task", action="store_true", help="Install scheduled task (run at user logon)")
    parser.add_argument("--install-startup", action="store_true", help="Install scheduled task to run at system startup (admin required)")
    parser.add_argument("--uninstall-task", action="store_true", help="Remove scheduled task")
    parser.add_argument("--run", action="store_true", help=argparse.SUPPRESS)  # internal: used by the task
    args: argparse.Namespace = parser.parse_args()

    if args.uninstall_task:
        uninstall_task()
        return
    if args.install_startup:
        install_task(on_startup=True, as_system=True)
        return
    if args.install_task:
        install_task(on_startup=False, as_system=False)
        return
    if args.run:
        run_watcher()
        return

    # Default: ensure it runs via scheduler
    if not task_exists():
        logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
        logging.info("Scheduled task not found. Creating it to run at user logon.")
        install_task(on_startup=False, as_system=False)
    start_task()
    # Exit immediately; the watcher will run under Task Scheduler

if __name__ == "__main__":
    main()
