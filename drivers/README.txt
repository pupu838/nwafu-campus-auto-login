Optional offline driver bundle
=============================

Place `msedgedriver.exe` in this folder before building the release EXE.

Why this helps:
- On a brand new machine, Selenium Manager may need network access to download EdgeDriver.
- Campus-network first login often starts without full internet access.
- Bundling EdgeDriver avoids that first-run dependency.

Build with:
- `powershell -ExecutionPolicy Bypass -File .\build_release.ps1`

Notes:
- Keep Edge and msedgedriver major versions compatible.
- If you do not provide this file, the app will still try Selenium Manager fallback.
