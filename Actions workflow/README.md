# DataProcessor - Automated Build System

This repository automatically builds Windows (.exe) and Mac (.app) executables using GitHub Actions.

## 🚀 Quick Start

### Step 1: Set Up Repository
1. Create a new GitHub repository
2. Clone it locally: `git clone https://github.com/yourusername/your-repo.git`
3. Copy all files into the repository:
   - `ArtWork.py` (your main script)
   - `requirements.txt`
   - `build_config.py`
   - `.github/workflows/build.yml`
   - This `README.md`

### Step 2: Push to GitHub
```bash
git add .
git commit -m "Initial commit with build system"
git push origin main
```

### Step 3: Download Built Executables

#### Option A: Download from Actions (Automatic builds)
1. Go to your repository on GitHub
2. Click the **"Actions"** tab
3. Click on the latest workflow run
4. Scroll down to **"Artifacts"** section
5. Download:
   - `DataProcessor-Windows` → Contains `.exe` file
   - `DataProcessor-Mac` → Contains `.app` bundle (zipped)

#### Option B: Download from Releases (Tagged versions)
1. Create a version tag:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
2. Go to **"Releases"** section on GitHub
3. Download the executables from the latest release

## 📦 Build Artifacts

| Platform | File | Description |
|----------|------|-------------|
| Windows | `DataProcessor.exe` | Single executable file |
| Mac | `DataProcessor-Mac.zip` | Zipped .app bundle |
| Mac | `DataProcessor.dmg` | DMG installer (optional) |

## 🛠️ Manual Build (Local)

### Windows
```bash
pip install -r requirements.txt
python build_config.py
# Output: dist/DataProcessor.exe
```

### Mac
```bash
pip3 install -r requirements.txt
python3 build_config.py
# Output: dist/DataProcessor.app
```

## 📋 Requirements

- Python 3.11
- All dependencies in `requirements.txt`
- GitHub account with Actions enabled

## 🔧 Customization

### Change App Name
Edit `build_config.py`:
```python
'--name=YourAppName',  # Change this line
```

### Add Icon
- Windows: Add `icon.ico` file and update `build_config.py`:
  ```python
  '--icon=icon.ico',
  ```
- Mac: Add `icon.icns` file and update `build_config.py`:
  ```python
  '--icon=icon.icns',
  ```

## 📝 SharePoint Support

The app includes full SharePoint support through:
- Local file system access to synced SharePoint folders
- Excel file handling with `openpyxl` and `xlsxwriter`
- Hidden/protected sheet reading capability

## 🐛 Troubleshooting

### Build Fails
1. Check the Actions tab for error logs
2. Ensure all dependencies are in `requirements.txt`
3. Verify Python version compatibility

### App Won't Run
- **Windows**: May need Visual C++ Redistributable
- **Mac**: May need to allow app in Security & Privacy settings:
  ```bash
  # Remove quarantine attribute
  xattr -cr /Applications/DataProcessor.app
  ```

### Missing Modules
Add hidden imports to `build_config.py`:
```python
'--hidden-import=missing_module_name',
```

## 📄 License

Your license here

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request
