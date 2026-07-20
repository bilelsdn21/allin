"""
Setup checker for the Desktop\\code projects (new-PC survival tool).

WHAT IT DOES
  1. Checks Python version.
  2. Checks every library the 3 projects need.
  3. Installs ONLY the missing ones (nothing is reinstalled).
  4. Checks tkinter, pythonw, and Google Chrome.
  5. Prints a simple PASS/FAIL report at the end.

Run it by double-clicking SETUP.bat, or:
    python setup_check.py            (report + install what's missing)
    python setup_check.py --check    (report only, install nothing)

Uses only Python's standard library, so it always runs.
"""

import importlib.util
import os
import subprocess
import sys

# import name -> pip package name (they are not always the same)
REQUIREMENTS = {
    # --- check passport ---
    "PIL": "pillow",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlrd": "xlrd",            # old .xls name lists
    "fitz": "pymupdf",         # PDF passports
    "tkinterdnd2": "tkinterdnd2",  # drag-and-drop (optional but wanted)
    # --- BUS PROJECT ---
    "streamlit": "streamlit",
    "requests": "requests",
    # --- see placess ---
    "flask": "flask",
    "werkzeug": "werkzeug",
    "selenium": "selenium",
}

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
]


def ok(msg):
    print(f"  [OK]      {msg}")


def missing(msg):
    print(f"  [MISSING] {msg}")


def main():
    check_only = "--check" in sys.argv
    problems = []

    print("=" * 60)
    print("SETUP CHECK - Desktop\\code projects")
    print("=" * 60)

    # 1. Python itself ------------------------------------------------------
    v = sys.version_info
    print(f"\nPython: {v.major}.{v.minor}.{v.micro}")
    if (v.major, v.minor) >= (3, 10):
        ok("Python version is good (3.10+)")
    else:
        missing("Python is too old - install Python 3.11+ "
                "(Microsoft Store is fine)")
        problems.append("Python 3.11+")

    # 2. tkinter (ships WITH Python, cannot be pip-installed) ---------------
    if importlib.util.find_spec("tkinter"):
        ok("tkinter (windows/buttons) is available")
    else:
        missing("tkinter is not available - reinstall Python and tick "
                "the 'tcl/tk' option. pip CANNOT fix this one.")
        problems.append("tkinter")

    # 3. pythonw (starts the app without a black console) -------------------
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if os.path.exists(pythonw):
        ok(f"pythonw found: {pythonw}")
    else:
        missing("pythonw.exe not found next to python.exe - the desktop "
                "shortcut must use python.exe instead (black window shows).")
        problems.append("pythonw (cosmetic only)")

    # 4. libraries - check first, install ONLY what's missing ---------------
    print("\nLibraries:")
    to_install = []
    for import_name, pip_name in REQUIREMENTS.items():
        if importlib.util.find_spec(import_name):
            ok(pip_name)
        else:
            missing(pip_name)
            to_install.append(pip_name)

    if to_install:
        if check_only:
            print(f"\n--check mode: would install -> {' '.join(to_install)}")
            problems.append(f"libraries: {', '.join(to_install)}")
        else:
            print(f"\nInstalling {len(to_install)} missing librarie(s): "
                  f"{' '.join(to_install)}")
            print("(only these - nothing already installed is touched)\n")
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--user",
                 *to_install])
            if r.returncode == 0:
                print("\nInstall finished.")
                # re-verify
                importlib.invalidate_caches()
                still = [p for i, p in REQUIREMENTS.items()
                         if p in to_install
                         and not importlib.util.find_spec(i)]
                if still:
                    problems.append(f"libraries that failed: "
                                    f"{', '.join(still)}")
                else:
                    ok("all libraries are now installed")
            else:
                problems.append(f"pip install failed for: "
                                f"{', '.join(to_install)} "
                                f"(is there internet?)")
    else:
        print("\n  Nothing to install - all libraries already there.")

    # 5. Google Chrome (needed by 'see placess' downloader only) ------------
    print("\nGoogle Chrome (needed only for 'see placess' downloader):")
    if any(os.path.exists(p) for p in CHROME_PATHS):
        ok("Chrome is installed")
    else:
        missing("Chrome not found - install it ONLY if you use the "
                "'see placess' downloader. The other projects don't "
                "need it.")
        problems.append("Chrome (only for see placess)")

    # 6. final report -------------------------------------------------------
    print("\n" + "=" * 60)
    if problems:
        print("RESULT: NOT READY YET. Still needed:")
        for p in problems:
            print(f"   - {p}")
        print("Fix the above (or ask Claude), then run this again.")
    else:
        print("RESULT: ALL GOOD. Every project can run on this PC.")
    print("=" * 60)
    return 1 if problems else 0


if __name__ == "__main__":
    code = main()
    input("\nPress Enter to close...")
    sys.exit(code)
