"""
MasterAutomation Desktop Launcher
Wrapper entrypoint for frozen PyInstaller Windows executable.
"""
import os
import sys
import multiprocessing

# Critical: Prevents child process fork bombs on Windows with PyInstaller
multiprocessing.freeze_support()

import socket
import threading
import time
import webbrowser
from pathlib import Path


def find_free_port(preferred: int = 8501) -> int:
    """Find a free TCP port, starting with the preferred port."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind(("127.0.0.1", preferred))
            return preferred
    except OSError:
        pass

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def get_base_dir() -> Path:
    """Determine base directory whether running from source or frozen binary."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def patch_streamlit_static_dir(base_dir: Path) -> Path | None:
    """Ensure Streamlit serves bundled static assets in frozen environments.

    In frozen PyInstaller bundles (especially on systems without Python installed),
    Streamlit's default file_util.get_static_dir() attempts to resolve relative to
    file_util.__file__, which points to a nonexistent development site-packages path
    or inside base_library.zip. If not patched to point to the actual extracted
    static files directory, Streamlit mounts no static routes and Starlette returns
    a 404 'Not Found' error for all web pages.
    """
    try:
        import streamlit.file_util as file_util
        orig_get_static_dir = file_util.get_static_dir

        meipass = getattr(sys, "_MEIPASS", None)
        exe_dir = Path(sys.executable).parent

        candidates = [
            base_dir / "streamlit" / "static",
            base_dir / "static",
            exe_dir / "_internal" / "streamlit" / "static",
            exe_dir / "streamlit" / "static",
        ]
        if meipass:
            candidates.insert(0, Path(meipass) / "streamlit" / "static")
            candidates.insert(1, Path(meipass) / "static")

        resolved = None
        for cand in candidates:
            if cand.is_dir() and (cand / "index.html").is_file():
                resolved = cand.resolve()
                break

        if resolved is None and getattr(sys, "frozen", False):
            # Fallback search if static was somehow nested elsewhere in bundle
            search_roots = [base_dir, exe_dir]
            if meipass and Path(meipass) not in search_roots:
                search_roots.append(Path(meipass))
            for root in search_roots:
                if root.is_dir():
                    for match in root.rglob("index.html"):
                        if "streamlit" in str(match) or "static" in str(match):
                            resolved = match.parent.resolve()
                            break
                    if resolved:
                        break

        if resolved is not None:
            static_dir_str = str(resolved)

            def _frozen_get_static_dir() -> str:
                return static_dir_str

            file_util.get_static_dir = _frozen_get_static_dir
            if "streamlit.file_util" in sys.modules:
                sys.modules["streamlit.file_util"].get_static_dir = _frozen_get_static_dir

            # Also proactively monkeypatch starlette_static_routes if loaded
            if "streamlit.web.server.starlette.starlette_static_routes" in sys.modules:
                sys.modules["streamlit.web.server.starlette.starlette_static_routes"].file_util.get_static_dir = _frozen_get_static_dir

            print(f"  [Static Assets] Loaded from: {resolved}")
            return resolved
        else:
            print("  [Static Assets] Using default Streamlit static resolution")
            return None
    except Exception as e:
        print(f"  [Warning] Failed to patch Streamlit static directory: {e}")
        return None


def configure_output_streams() -> None:
    """Keep diagnostics available when the packaged app has no console window."""
    if sys.stdout is not None and sys.stderr is not None:
        return
    try:
        log_path = Path(sys.executable).resolve().parent / "masterautomation.log"
        stream = open(log_path, "a", encoding="utf-8", buffering=1)
    except OSError:
        stream = open(os.devnull, "w")
    if sys.stdout is None:
        sys.stdout = stream
    if sys.stderr is None:
        sys.stderr = stream


def main():
    configure_output_streams()
    base_dir = get_base_dir()
    os.chdir(base_dir)

    if str(base_dir) not in sys.path:
        sys.path.insert(0, str(base_dir))

    patch_streamlit_static_dir(base_dir)

    # Point to the app entrypoint
    app_path = base_dir / "app.py"
    if not app_path.exists():
        app_path = base_dir / "master_update_automation" / "app.py"

    port = find_free_port(8501)
    url = f"http://localhost:{port}"

    print("=" * 64)
    print("  Master Data Automation — Consumer Master Update System")
    print("=" * 64)
    print(f"  Starting local server on: {url}")
    print("  Your default web browser will open automatically...")
    print("  To exit, close this window or press Ctrl+C.")
    print("=" * 64)

    # Background browser opener
    def _open_browser():
        time.sleep(1.8)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=_open_browser, daemon=True).start()

    # Import and run Streamlit via bootstrap in-process
    import streamlit.web.bootstrap as bootstrap
    from streamlit import config as _config

    # Set config options programmatically (takes precedence over config.toml)
    _config.set_option("server.port", port)
    _config.set_option("server.headless", False)
    _config.set_option("server.maxUploadSize", 3072)
    _config.set_option("browser.gatherUsageStats", False)
    _config.set_option("global.developmentMode", False)

    bootstrap.run(
        str(app_path),
        is_hello=False,
        args=[],
        flag_options={},
    )


if __name__ == "__main__":
    main()
