# ClickShot

**Fast, clean screenshots with instant naming — auto-saved to PNG + PowerPoint + Word**  
Windows 10/11 • Buttons only (no hotkeys) • Multi‑monitor aware • Taskbar excluded

---

## ✨ What ClickShot does

- Starts a **new session** every time you launch it (you’ll choose a **save folder**, **project name**, and an optional **PPTX template**).
- Lets you capture using **three buttons**:
  1. **Current Monitor** — captures the monitor under your mouse **excluding** the taskbar.
  2. **Active Window (Click to pick)** — click any window; ClickShot captures its clean bounds (no desktop, no shadows).
  3. **Region (Drag a rectangle)** — Snipping‑Tool style crosshair overlay across all monitors.
- The app **hides itself** before each capture so it never appears in your screenshots.
- After capture, a **name dialog appears with a live preview** (Entry is auto‑focused). Press **Enter** to save instantly.
- Each save writes:
  - a **PNG** into your **project folder**, and
  - a new slide in a **PPTX**, and
  - a new page in a **DOCX**.
- A small **“Saved successfully”** toast appears and the app returns to the front, ready for the next shot.
- Close the app to end the session; re‑launch to create a **new** PPTX/DOCX in a new (or the same) project folder.

> ClickShot is designed for **documentation speed**: shoot → name → saved (PNG + slide + page) → repeat.

---

## 🖥️ System requirements

- **OS**: Windows 10 (1903+) or Windows 11
- **Python (for running from source)**: 3.9–3.12
- **RAM**: 4GB+ recommended for high‑res monitors

---

## 📦 Install & Run (from source)

1. **Download** `main.py` and `requirements.txt` into a folder (e.g., `CLICKSHOT/`).  
2. Open **PowerShell** in that folder and create a virtual environment:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate
   ```
3. **Install dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```
4. **Run**:
   ```powershell
   python .\main.py
   ```

> Running **as Administrator** is recommended (right‑click PowerShell → “Run as Administrator”) for the cleanest window picking and overlay behavior.

---

## 🚀 First‑run Setup (every session)

On launch, ClickShot will ask for:

1. **Save folder** — where PNGs and documents will be stored.
2. **Project name** — used to create a project directory and to name the PPTX/DOCX.
3. **PPTX template (optional)** — choose any `.pptx` to apply your organization branding.

Resulting structure:
```
<Your Save Folder>/
  └─ <Project_Name>/
       ├─ <Project_Name>.pptx
       ├─ <Project_Name>.docx
       └─ *.png  (your screenshots)
```

---

## 🎯 Using ClickShot

### 1) Capture Current Monitor (taskbar excluded)
- Click **“Capture Current Monitor”**.
- The app hides; ClickShot captures the monitor under your mouse **excluding** the taskbar for that monitor.
- A **naming dialog** appears with a preview. Type a name and press **Enter**.
- PNG is saved and appended to PPTX/DOCX. The app returns to the front with a **Saved** toast.

### 2) Capture Active Window (click to pick)
- Click **“Capture Active Window”**.
- A transparent overlay appears — **click the window** you want.
- ClickShot uses **DWM extended frame bounds** to tightly capture that window content (no desktop/taskbar).
- Name → Save → Toast → App returns.

### 3) Capture Region (drag rectangle)
- Click **“Capture Region”**.
- A crosshair overlay covers all monitors (virtual desktop).
- **Click‑drag** to select any rectangle; release to capture. Press **Esc** to cancel.
- Name → Save → Toast → App returns.

> After each save, the **Entry field is already selected**. You can immediately type or simply hit **Enter** to accept the default filename.

---

## 🧠 Behavior & Notes

- **App never appears in screenshots**: it withdraws before capture and resurfaces after saving.
- **Multi‑monitor aware**: Region overlay spans all monitors. “Current Monitor” uses the **monitor under your mouse** and excludes that monitor’s taskbar.  
- **Unique filenames**: If a name already exists, `_1`, `_2`, … are appended automatically.
- **Documents**:
  - **PPTX**: each shot becomes a centered image on a new slide with a caption.
  - **DOCX**: each shot becomes a new page with a centered image and heading.

---

## 🛠️ Troubleshooting

- **Permission/UAC**: If overlays don’t appear above some elevated apps, run ClickShot as Administrator.
- **Antivirus false positives** (for packaged `.exe`): add ClickShot to your AV allow‑list.
- **Fonts/UI scaling**: ClickShot is **per‑monitor DPI aware**; if text looks tiny, adjust Windows Scale settings.
- **App not returning to front**: Windows sometimes blocks z‑order. ClickShot forces a brief `topmost` raise; if needed, click the taskbar icon once.

---

## 🧩 Project files

- `main.py` — application source
- `requirements.txt` — Python dependencies

You can customize icons, colors, and defaults inside `main.py` if you like.

---

## 🏗️ Build a **single‑file .exe** (PyInstaller)

Use PyInstaller to package ClickShot into a portable executable that runs without Python installed.

1. **Create a clean venv and install deps + PyInstaller**:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate
   pip install -r requirements.txt
   pip install pyinstaller
   ```

2. **Build** (GUI app, single file, elevation prompt for admin):
   ```powershell
   pyinstaller --noconsole --onefile --name ClickShot --uac-admin main.py
   ```

   Optional extras:
   - Add an icon: `--icon path\to\clickshot.ico`
   - Reduce AV false‑positives: avoid `--onefile` (use folder mode), or sign the binary (see below).

3. **Result**: `dist\ClickShot.exe`

4. **Test** by double‑clicking `ClickShot.exe`. On first run you’ll pick a save folder, project name, and (optionally) a PPTX template.

### Optional: Using a `.spec` file

If you prefer a reproducible build, create `ClickShot.spec` like this:

```python
# ClickShot.spec (example)
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pptx', 'pptx.oxml', 'pptx.enum', 'docx', 'docx.oxml'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    name='ClickShot',
    icon=None,           # set to 'clickshot.ico' if you have one
    console=False,
    uac_admin=True,      # request elevation on launch
)
```
Build with:
```powershell
pyinstaller ClickShot.spec
```

> **Note:** Hidden imports above are conservative. If your build works without them, you can remove them.

---

## 📦 Build an **Installer .exe** (Inno Setup)

To produce a Windows installer that places the app in Program Files and creates shortcuts:

1. **Install** [Inno Setup](https://jrsoftware.org/isinfo.php).
2. Ensure you have `dist\ClickShot.exe` from PyInstaller.
3. Create `ClickShot.iss` (installer script) like this:

```ini
[Setup]
AppName=ClickShot
AppVersion=1.0
DefaultDirName={pf}\ClickShot
DefaultGroupName=ClickShot
OutputBaseFilename=ClickShot-Setup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\ClickShot.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ClickShot"; Filename: "{app}\ClickShot.exe"
Name: "{commondesktop}\ClickShot"; Filename: "{app}\ClickShot.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked
```

4. Open the `.iss` in Inno Setup and click **Compile**.  
5. You’ll get `Output\ClickShot-Setup.exe` — run it to install.

**Optional:** In Inno Setup you can set `Run as administrator` compatibility for the installed shortcut or embed a manifest in the PyInstaller step using `--uac-admin` (recommended).

---

## 🔐 Code signing (recommended for distribution)

If distributing to teammates, sign the EXE to avoid SmartScreen prompts:
```powershell
signtool sign /fd SHA256 /a /tr http://timestamp.digicert.com /td SHA256 dist\ClickShot.exe
```

---

## 🧰 Tips for smooth sessions

- Keep a **branded PPTX template** ready to get professional slides instantly.
- Use **Window capture** for clean app documentation (no desktop clutter).
- Use **Region** for specific UI elements or bug highlights.
- Press **Enter** in the name dialog to save as fast as possible.

---

## ❓ FAQ

**Q: Does it capture the taskbar?**  
**A:** No. “Current Monitor” uses the **Work area** of that monitor (taskbar excluded).

**Q: Can it capture multiple monitors at once?**  
**A:** Not in one shot; the design favors clarity and speed. You can take monitor shots one after another.

**Q: Where do PPTX/DOCX get saved?**  
**A:** In your **project folder** chosen at session start (same place as PNGs).

**Q: Do I need Office installed?**  
**A:** No. Files are created via `python-pptx` and `python-docx`. You only need Office to open/edit them.

**Q: Why run as Administrator?**  
**A:** Some elevated apps and windows can sit above normal overlays. Elevation ensures our overlays and pickers stay on top.

---

## 🛡️ Privacy

- No data leaves your machine.  
- No telemetry.  
- All captures and documents are stored **locally** in the folder you pick.

---

## 🧑‍💻 Support

If you need tweaks (tray icon, different defaults, auto‑TOC, custom slide layout, etc.), open an issue or ping your ClickShot maintainer.

---

## 📄 License

This project is provided “as‑is” for internal use. Dependencies: MSS (MIT), Pillow (HPND), python‑docx (MIT), python‑pptx (MIT).