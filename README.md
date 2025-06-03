# KAU Video Downloader

A cross-platform application for downloading and organizing videos from Google Drive links listed in a spreadsheet.

## Features

- Downloads videos from Google Drive links
- Organizes videos into a structured directory hierarchy
- Creates an Excel workbook with hyperlinks to downloaded videos
- Progress tracking and status updates
- Cross-platform support (Windows, macOS, Android, iOS)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Qamarzaman555/kyvi_video_downloader.git
cd kyvi_video_downloader
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### 1. Run with Python (All Platforms)

Make sure you have Python 3.9+ and all dependencies installed:

```bash
pip install -r requirements.txt
python main.py
```

### 2. Run as an Executable

#### On Windows
- Build the .exe on a Windows machine (see below for build steps).
- After building, run:

```cmd
dist\main.exe
```

#### On macOS
- Build the binary on macOS (see below for build steps).
- After building, run:

```sh
./dist/main
```

### Notes
- The first time you run the script, it will download the spreadsheet and start downloading videos into organized folders.
- The output folder structure will be:
  - `<Root Folder>/<Subject>/<Topic>/<Subtopic>/<Video Title>.mp4`
- The script logs progress to `video_downloader.log`.

## Building Executables

### Windows (.exe)
1. Open a terminal (cmd or PowerShell) on a Windows machine.
2. Install dependencies:
   ```cmd
   pip install -r requirements.txt
   pip install pyinstaller
   ```
3. Build the executable:
   ```cmd
   pyinstaller --onefile --add-data "requirements.txt;." main.py
   ```
   - The .exe will be in the `dist` folder as `main.exe`.

### macOS (Binary)
1. Open a terminal on macOS.
2. Install dependencies:
   ```sh
   pip3 install -r requirements.txt
   pip3 install pyinstaller
   ```
3. Build the executable:
   ```sh
   pyinstaller --onefile --add-data "requirements.txt:." main.py
   ```
   - The binary will be in the `dist` folder as `main`.

### Cross-Platform Note
- You must build the executable on the target OS (Windows for .exe, macOS for binary). PyInstaller does not cross-compile.
- If you need a Windows .exe but only have a Mac, use a Windows PC or a Windows VM.

## Project Structure

```
<Root Directory>/
├── KAUvideos.xlsx
├── <Subject>/
│   ├── <Topic>/
│   │   ├── <Subtopic>/
│   │   │   └── <Video Title>.mp4
```

## Requirements

- Python 3.9+
- See requirements.txt for Python package dependencies

## License

MIT License 