# Crimson Desert Teleporter Tool

A standalone utility for Crimson Desert that adds quick map-marker teleportation, waypoint management, and temporary invulnerability.

## Features

- **F5** — Teleport to the currently placed map destination marker
- **F6** — Save current map marker as a named waypoint
- **F8** — Abort / return to pre-teleport position
- 10-second invulnerability window after each teleport
- Local + community-shared waypoint library
- Integrated MapGenie overlay for marker placement

All hotkeys are rebindable from the in-app settings.

## Usage

Each versioned folder (`v*/`) contains the Python source for that version. Prebuilt Windows executables and source zips for each version are attached to the matching [GitHub Release](https://github.com/DameonK/Crimson-Desert-Teleporation-Tool/releases) (the latest is v2.1.6).

To use:

- Grab the exe from Releases and launch as Administrator, or
- Clone this repo and run `v2.1.6/run_teleporter.bat` — it bootstraps Python, installs `pymem` + `pywebview` into a local `pylibs/`, then starts the tool, or
- Clone and run `v2.1.6/build_exe.bat` to rebuild the exe yourself (see below).

The game must be running. The tool attaches to `CrimsonDesert.exe` and installs code-cave hooks for position, velocity, health, and map-marker capture.

## Building from source (`build_exe.bat`)

`build_exe.bat` produces a fresh `dist/CrimsonDesertTeleporter.exe` from `cd_teleporter.py`. It does the following, in order, and no steps require any global installs:

1. **Locate a Python** with `tkinter` — checks, in order: a bundled `python/python.exe` next to the script, the `py` launcher, `where python` on PATH, and a handful of common install paths under `%LocalAppData%` / `%ProgramFiles%`. Windows Store stub pythons are skipped.
2. **Ensure dependencies** — if `pymem` or `pywebview` can't be imported, bootstraps `pip` (via `ensurepip` if needed) and installs `pymem`, `pywebview`, `bottle`, and `proxy_tools` into a local `pylibs/` folder — no global site-packages pollution.
3. **Ensure PyInstaller** — installs it into the chosen Python if missing.
4. **Build** — runs PyInstaller with `--onefile --windowed --uac-admin` (the exe requests Administrator rights on launch, which the tool needs to read game memory), embeds `teleporter.ico`, includes `pymem`/`webview` submodules, and points at the local `pylibs/`.

Output lands at `dist/CrimsonDesertTeleporter.exe`. Re-running the script re-uses the cached `pylibs/` and is much faster the second time.

## Game version compatibility

The tool relies on AOB (array-of-bytes) signatures to locate hook points inside the game's executable. Every time the game updates, those signatures may shift and need re-verifying before the tool will attach cleanly.

## License

MIT — see `LICENSE`.
