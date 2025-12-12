# Building the DMC Automation Tool Executable

## Prerequisites

1. **Python 3.8 or higher** installed on your system
2. **pip** (Python package installer)

## Step-by-Step Build Instructions

### 1. Install Dependencies

Open PowerShell or Command Prompt in the project directory and run:

```powershell
pip install -r requirements.txt
```

### 2. Install PyInstaller

```powershell
pip install pyinstaller
```

### 3. Build the Executable

Run the following command to build the executable using the spec file:

```powershell
pyinstaller DMC_Auto_GUI.spec
```

This will create:
- A `dist` folder containing the executable: `DMC_Automation_Tool.exe`
- A `build` folder (temporary build files)

### 4. Locate Your Executable

The final executable will be located at:
```
dist\DMC_Automation_Tool.exe
```

## Distribution

### What to Include

When distributing the application, you need to include:

1. **DMC_Automation_Tool.exe** - The main executable
2. **Lake folder** - Contains SNS data files (bundled inside the exe)
3. **documents_to_process folder** - For input documents (create if needed)
4. **output folder** - For processed documents (created automatically)

### Folder Structure for Distribution

```
DMC_Automation_Tool/
├── DMC_Automation_Tool.exe
├── documents_to_process/  (for input .docx files)
├── output/                (generated automatically)
└── logs/                  (generated automatically)
```

**Note:** The `Lake` folder is bundled inside the executable, so users don't need it separately unless they want to modify the SNS data files.

## Important Notes

### Ollama Requirement

The application requires **Ollama** to be running locally for AI-powered DMC generation:

1. Download and install Ollama from: https://ollama.ai
2. Pull the required model:
   ```powershell
   ollama pull llama3.1:8b
   ```
3. Ensure Ollama is running (it runs as a service by default)

The application will work with fallback methods if Ollama is not available, but results will be less accurate.

### First Run

On first run, the application will:
- Create `documents_to_process` folder if it doesn't exist
- Create `output` folder if it doesn't exist
- Create `logs` folder if it doesn't exist

### Antivirus Warning

Some antivirus software may flag PyInstaller executables as suspicious. This is a false positive. You may need to:
- Add an exception in your antivirus software
- Digitally sign the executable (for professional distribution)

## Troubleshooting

### "Failed to execute script" error

If you get this error:
1. Try running with console enabled to see error messages:
   - Edit `DMC_Auto_GUI.spec` and change `console=False` to `console=True`
   - Rebuild with `pyinstaller DMC_Auto_GUI.spec`

### Missing dependencies

If the executable fails due to missing modules:
1. Add the missing module to `hiddenimports` in the spec file
2. Rebuild the executable

### Large file size

The executable is large (~40-50 MB) because it includes:
- Python runtime
- All required libraries
- Tkinter GUI framework
- Data files

This is normal for PyInstaller applications and ensures the exe works on any Windows system without requiring Python installation.

## Advanced: Creating an Installer

For professional distribution, consider creating an installer using:
- **Inno Setup** (free): https://jrsoftware.org/isinfo.php
- **NSIS** (free): https://nsis.sourceforge.io/
- **Advanced Installer** (commercial)

This will create a proper installer that:
- Installs to Program Files
- Creates Start Menu shortcuts
- Handles uninstallation
- Can check for Ollama installation
