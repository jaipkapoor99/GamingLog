import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import psutil

try:
    import winreg
except ImportError:
    winreg = None  # type: ignore[assignment]

from dotenv import load_dotenv

# type: ignore[import-untyped]
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build  # type: ignore[import-untyped]
from googleapiclient.errors import HttpError  # type: ignore[import-untyped]

# Load environment variables from .env file
# Look for .env in the same directory as the executable/script
if getattr(sys, "frozen", False):
    # Running as compiled executable
    dotenv_path = Path(sys.executable).parent / ".env"
else:
    # Running as script
    dotenv_path = Path(__file__).parent / ".env"

load_dotenv(dotenv_path=dotenv_path)

# ---- Constants ----
SECONDS_TO_MINUTES: float = 60.0

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
    Format: "exe" or "exe=Display Name" entries separated by semicolons.
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
                display: str = os.path.splitext(os.path.basename(exe))[0]
                mapping[exe] = display
    return mapping


def build_game_map() -> Dict[str, str]:
    """Build game map from environment variable or use defaults."""
    user_map: Dict[str, str] = parse_games_env(GAMES_ENV)
    if user_map:
        return user_map
    return {k.lower(): v for k, v in DEFAULT_GAMES.items()}


def _get_steam_path_from_registry() -> Optional[str]:
    """Get Steam installation path from Windows registry."""
    if winreg is None:
        return None

    try:
        # Try 64-bit registry first
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Valve\Steam"
        )
        path, _ = winreg.QueryValueEx(key, "InstallPath")
        winreg.CloseKey(key)
        return str(path) if path else None
    except OSError:
        try:
            # Try 32-bit registry
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Valve\Steam")
            path, _ = winreg.QueryValueEx(key, "InstallPath")
            winreg.CloseKey(key)
            return str(path) if path else None
        except OSError:
            return None


def _parse_libraryfolders_vdf(vdf_path: Path) -> Set[str]:
    """Parse Steam's libraryfolders.vdf file to find library paths."""
    libs: Set[str] = set()

    if not vdf_path.exists():
        logging.debug("VDF file not found: %s", vdf_path)
        return libs

    try:
        text: str = vdf_path.read_text(encoding="utf-8", errors="ignore")
        # Support both old and new VDF formats
        # Old: "0"  "C:\\Path"
        libs.update(
            [
                m.group(1)
                for m in re.finditer(r'"\d+"\s*"([^"]+)"', text, flags=re.IGNORECASE)
            ]
        )
        # New: "path" "C:\\Path"
        libs.update(
            [
                m.group(1)
                for m in re.finditer(r'"path"\s*"([^"]+)"', text, flags=re.IGNORECASE)
            ]
        )
        logging.debug("Found %d Steam library paths in VDF", len(libs))
    except (OSError, UnicodeDecodeError) as e:
        logging.warning("Failed to parse VDF file %s: %s", vdf_path, e)

    return {os.path.normpath(p) for p in libs if p}


def get_steam_library_common_dirs() -> List[str]:
    """Find Steam library 'common' directories (Windows only)."""
    # Try registry first to support custom Steam installations
    steam_path = _get_steam_path_from_registry()
    if not steam_path:
        # Fallback to default installation path
        steam_path = r"C:\Program Files (x86)\Steam"

    libs: Set[str] = set()

    if os.path.isdir(steam_path):
        libs.add(os.path.normpath(steam_path))
        vdf: Path = Path(steam_path) / "steamapps" / "libraryfolders.vdf"
        libs.update(_parse_libraryfolders_vdf(vdf))
        logging.debug("Steam installation found at: %s", steam_path)
    else:
        logging.debug("Steam not found at: %s", steam_path)

    # Convert to steamapps/common directories
    common_dirs: List[str] = []
    for lib in libs:
        common_dir: Path = Path(lib) / "steamapps" / "common"
        try:
            if common_dir.is_dir():
                common_dirs.append(os.path.normcase(str(common_dir.resolve())))
        except OSError:
            common_dirs.append(os.path.normcase(os.path.normpath(str(common_dir))))

    # Deduplicate while preserving order
    seen: Set[str] = set()
    dedup: List[str] = []
    for d in common_dirs:
        if d not in seen:
            seen.add(d)
            dedup.append(d)
    return dedup


def _nice_title(name: str) -> str:
    """Convert a directory or file name to a nice display title."""
    s: str = re.sub(r"[_\-]+", " ", name).strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s.title() if s else name


def derive_game_name_from_path(exe_path: str, library_common_dirs: List[str]) -> str:
    """Derive a game display name from its installation path."""
    exe: Path = Path(exe_path)
    exe_lower: str = os.path.normcase(str(exe))

    for lib in library_common_dirs:
        lib_lower: str = os.path.normcase(lib)
        if exe_lower.startswith(lib_lower):
            try:
                rel: str = os.path.relpath(exe_lower, lib_lower)
                parts: List[str] = [
                    p for p in Path(rel).parts if p not in (os.sep, "/", "")
                ]
                if parts:
                    return _nice_title(parts[0])
            except (ValueError, OSError):
                pass

    # Fallback: use parent folder name
    try:
        return _nice_title(exe.parent.name or exe.stem)
    except (AttributeError, OSError):
        return exe.stem


def detect_running_games_steam(
    library_common_dirs: List[str],
) -> Dict[int, Tuple[str, str]]:
    """Detect running games from Steam library directories.
    Only tracks processes with memory > 2GB to avoid launcher processes."""
    MIN_MEMORY_BYTES: int = 2 * 1024 * 1024 * 1024  # 2GB
    matches: Dict[int, Tuple[str, str]] = {}
    for p in psutil.process_iter(
        ["pid", "name", "exe", "memory_info"]
    ):  # type: ignore[misc]
        try:
            exe_path: str = p.info.get("exe") or ""
            if not exe_path:
                continue

            base: str = os.path.basename(exe_path).lower()
            if base in EXCLUDED_EXE_NAMES:
                continue

            exe_norm: str = os.path.normcase(os.path.normpath(exe_path))
            if any(exe_norm.startswith(lib) for lib in library_common_dirs):
                # Only track processes with memory > 2GB (actual game, not
                # launcher)
                memory_info = p.info.get("memory_info")
                if memory_info and memory_info.rss >= MIN_MEMORY_BYTES:
                    display: str = derive_game_name_from_path(
                        exe_path, library_common_dirs
                    )
                    matches[p.pid] = (base, display)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return matches


def _pythonw_executable() -> str:
    """Get path to pythonw.exe (or python.exe if not available)."""
    py: Path = Path(sys.executable)
    cand: Path = py.with_name("pythonw.exe")
    return str(cand if cand.exists() else py)


def _script_path() -> str:
    """Get path to the current script or executable."""
    if getattr(sys, "frozen", False):
        return sys.executable
    return str(Path(__file__).resolve())


def install_task() -> None:
    """Install autostart using Windows Startup folder (no admin required)."""
    startup_folder = (
        Path(os.environ.get("APPDATA", ""))
        / "Microsoft"
        / "Windows"
        / "Start Menu"
        / "Programs"
        / "Startup"
    )
    script_path = _script_path()
    shortcut_path = startup_folder / f"{TASK_NAME}.bat"

    # Create a batch file that launches the exe in the background
    batch_content = f'@echo off\nstart "" /min "{script_path}" --run'

    try:
        startup_folder.mkdir(parents=True, exist_ok=True)
        shortcut_path.write_text(batch_content, encoding="utf-8")
        logging.info("Created startup script at: %s", shortcut_path)
        print("✓ Startup script installed successfully!")
        print(f"  Location: {shortcut_path}")
        print("  The watcher will start automatically when you log in.")
    except OSError as e:
        logging.error("Failed to create startup script: %s", str(e))
        print("✗ ERROR: Failed to create startup script.")
        print(f"  Details: {str(e)}")
        sys.exit(1)


def uninstall_task() -> None:
    """Remove the startup script."""
    startup_folder = (
        Path(os.environ.get("APPDATA", ""))
        / "Microsoft"
        / "Windows"
        / "Start Menu"
        / "Programs"
        / "Startup"
    )
    shortcut_path = startup_folder / f"{TASK_NAME}.bat"

    logging.info("Deleting startup script '%s' if it exists...", shortcut_path)
    try:
        if shortcut_path.exists():
            shortcut_path.unlink()
            logging.info("Startup script removed.")
            print(f"✓ Startup script removed: {shortcut_path}")
        else:
            print(f"No startup script found at: {shortcut_path}")
    except OSError as e:
        logging.error("Failed to remove startup script: %s", str(e))
        print(f"✗ ERROR: Failed to remove startup script: {str(e)}")


def task_exists() -> bool:
    """Check if the startup script exists."""
    startup_folder = (
        Path(os.environ.get("APPDATA", ""))
        / "Microsoft"
        / "Windows"
        / "Start Menu"
        / "Programs"
        / "Startup"
    )
    shortcut_path = startup_folder / f"{TASK_NAME}.bat"
    return shortcut_path.exists()


def start_task() -> None:
    """Start the watcher in the background."""
    script_path = _script_path()
    logging.info("Starting watcher: %s", script_path)

    if getattr(sys, "frozen", False):
        # For .exe, start in background
        subprocess.Popen(  # pylint: disable=consider-using-with
            [script_path, "--run"],
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
    else:
        # For Python script, use pythonw
        subprocess.Popen(  # pylint: disable=consider-using-with
            [_pythonw_executable(), script_path, "--run"]
        )

    print("✓ Watcher started in background")
    logging.info("Watcher started")


class Session:  # pylint: disable=too-few-public-methods
    def __init__(
        self, pid: int, exe_name: str, display_name: str, start_time: datetime
    ) -> None:
        self.pid: int = pid
        self.exe_name: str = exe_name
        self.display_name: str = display_name
        self.start_time: datetime = start_time

    def finalize(self, end_time: datetime) -> Dict[str, str]:
        """Calculate session duration and return formatted data."""
        duration_seconds: int = max(
            0, int((end_time - self.start_time).total_seconds())
        )
        duration_minutes: float = round(duration_seconds / SECONDS_TO_MINUTES, 2)
        return {
            "game": self.display_name,
            "exe": self.exe_name,
            "start_iso": self.start_time.isoformat(timespec="seconds"),
            "end_iso": end_time.isoformat(timespec="seconds"),
            "duration_minutes": str(duration_minutes),
        }


def get_sheets_service() -> Optional[Any]:
    """Initialize Google Sheets API service."""
    if not GOOGLE_SERVICE_ACCOUNT_FILE or not os.path.exists(
        GOOGLE_SERVICE_ACCOUNT_FILE
    ):
        logging.error("GOOGLE_SERVICE_ACCOUNT_FILE not set or file not found.")
        return None
    try:
        creds: Credentials = Credentials.from_service_account_file(  # type: ignore[attr-defined]
            GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        return build(
            "sheets", "v4", credentials=creds, cache_discovery=False
        )  # type: ignore[no-any-return]
    except (OSError, ValueError) as e:
        logging.exception("Failed to initialize Google Sheets service: %s", e)
        return None


def ensure_sheet_exists(service: Any, spreadsheet_id: str, title: str) -> None:
    """Ensure the target sheet exists in the spreadsheet with proper headers."""
    try:
        meta: Dict[str, Any] = (
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        )
        sheets: List[Dict[str, Any]] = meta.get("sheets", [])

        # Check if sheet already exists
        for s in sheets:
            props: Dict[str, Any] = s.get("properties", {})
            if props.get("title") == title:
                return

        # Create the sheet if not present
        batch_req: Dict[str, Any] = {
            "requests": [{"addSheet": {"properties": {"title": title}}}]
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body=batch_req
        ).execute()

        # Add header row
        headers: List[List[str]] = [
            ["Game", "Duration (min)", "Started At", "Ended At"]
        ]
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{title}!A1:D1",
            valueInputOption="USER_ENTERED",
            body={"values": headers},
        ).execute()
    except HttpError as e:  # type: ignore[misc]
        logging.exception("Failed to ensure sheet exists: %s", e)
        raise


def append_session(
    service: Any, spreadsheet_id: str, title: str, payload: Dict[str, str]
) -> None:
    """Append a row to the sheet with session details."""
    row: List[str] = [
        payload["game"],
        payload["duration_minutes"],
        payload["start_iso"],
        payload["end_iso"],
    ]
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{title}!A:D",
        valueInputOption="USER_ENTERED",
        body={"values": [row]},
    ).execute()


def get_process_identity(p: psutil.Process) -> Tuple[str, str]:
    """Get process name and executable basename (deprecated - kept for compatibility)."""
    name: str = ""
    base: str = ""

    try:
        name = (p.name() or "").lower()
        base = name
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return "", ""

    try:
        exe_path: str = p.exe() or ""
        if exe_path:
            base = os.path.basename(exe_path).lower()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

    return name, base


def detect_running_games(game_map: Dict[str, str]) -> Dict[int, Tuple[str, str]]:
    """Detect running games based on the configured game map."""
    matches: Dict[int, Tuple[str, str]] = {}
    for p in psutil.process_iter(["pid", "name", "exe"]):  # type: ignore[misc]
        try:
            name = (p.info.get("name") or "").lower()
            exe_path = p.info.get("exe") or ""

            if not name:
                continue

            base: str = name
            if exe_path:
                base = os.path.basename(exe_path).lower()

            if not base:
                continue

            for key in (name, base):
                if key in game_map:
                    matches[p.pid] = (key, game_map[key])
                    break
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return matches


def run_watcher() -> None:  # pylint: disable=too-many-locals,too-many-branches,too-many-statements
    """Main game session monitoring loop."""
    # Configure logging with file output for background process
    # When frozen (running as .exe), use current directory; otherwise use
    # script directory
    if getattr(sys, "frozen", False):
        log_file = Path(sys.executable).parent / "gaminglog.log"
    else:
        log_file = Path(__file__).parent / "gaminglog.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
    )

    # Validate required configuration early
    if not GOOGLE_SHEET_ID:
        logging.error("GOOGLE_SHEET_ID environment variable is not set!")
        sys.exit(1)

    if not GOOGLE_SERVICE_ACCOUNT_FILE:
        logging.error("GOOGLE_SERVICE_ACCOUNT_FILE environment variable is not set!")
        sys.exit(1)

    if not os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE):
        logging.error("Service account file not found: %s", GOOGLE_SERVICE_ACCOUNT_FILE)
        sys.exit(1)

    service: Optional[Any] = get_sheets_service()
    if not service:
        logging.error("Failed to initialize Google Sheets service. Exiting.")
        sys.exit(1)

    # Ensure target sheet exists
    ensure_sheet_exists(service, GOOGLE_SHEET_ID, SHEET_TITLE)

    # Prefer Steam auto-detection
    library_common_dirs: List[str] = get_steam_library_common_dirs()
    use_steam: bool = bool(library_common_dirs)
    game_map: Dict[str, str] = {}

    if use_steam:
        logging.info("Auto-detecting games from Steam libraries:")
        for d in library_common_dirs:
            logging.info("  %s", d)
    else:
        game_map = build_game_map()
        if not game_map:
            logging.error("No Steam libraries found and no games configured.")
            logging.error(
                "Either install Steam to the default path or set GAMES environment variable."
            )
            sys.exit(1)
        logging.info(
            "Watching for games: %s", ", ".join(sorted(set(game_map.values())))
        )

    active_sessions: Dict[int, Session] = {}

    while True:
        try:
            # Detect running games
            running: Dict[int, Tuple[str, str]]
            if use_steam:
                running = detect_running_games_steam(library_common_dirs)
            else:
                running = detect_running_games(game_map)

            # Start sessions for newly detected games
            for pid, (exe_name, display_name) in running.items():
                if pid not in active_sessions:
                    active_sessions[pid] = Session(
                        pid=pid,
                        exe_name=exe_name,
                        display_name=display_name,
                        start_time=datetime.now(),
                    )
                    logging.info("Started session: %s (pid=%d)", display_name, pid)

            # Finalize sessions for games that have stopped (no longer in
            # running dict OR process doesn't exist)
            ended_pids: List[int] = []
            for pid in list(active_sessions.keys()):
                if pid not in running or not psutil.pid_exists(pid):
                    ended_pids.append(pid)

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
        except Exception as e:  # pylint: disable=broad-exception-caught
            logging.exception("Unexpected error in loop: %s", e)
            time.sleep(POLL_INTERVAL_SECONDS)


def main() -> None:
    """Main entry point for the application."""
    import argparse  # pylint: disable=import-outside-toplevel

    parser: argparse.ArgumentParser = argparse.ArgumentParser(
        prog="GamingLog", description="Game session logger for Windows"
    )
    parser.add_argument(
        "--install-task",
        action="store_true",
        help="Install scheduled task (run at user logon)",
    )
    parser.add_argument(
        "--install-startup",
        action="store_true",
        help="Install scheduled task to run at system startup (admin required)",
    )
    parser.add_argument(
        "--uninstall-task", action="store_true", help="Remove scheduled task"
    )
    parser.add_argument(
        "--run", action="store_true", help=argparse.SUPPRESS
    )  # Internal: used by the task
    args: argparse.Namespace = parser.parse_args()

    if args.uninstall_task:
        uninstall_task()
        return

    if args.install_startup:
        # --install-startup is deprecated, use --install-task instead
        install_task()
        return

    if args.install_task:
        install_task()
        return

    if args.run:
        run_watcher()
        return

    # Default: ensure it runs via scheduler
    if not task_exists():
        logging.basicConfig(
            level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
        )
        logging.info("Startup script not found. Creating it to run at user logon.")
        install_task()

    start_task()


if __name__ == "__main__":
    main()
