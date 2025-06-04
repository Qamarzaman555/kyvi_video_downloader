# KAU Video Downloader

A cross-platform application for downloading and organizing videos from Google Drive links listed in a spreadsheet. The application provides a user-friendly GUI interface for easy video downloading and organization.

## Features

- Downloads videos from Google Drive links
- Supports YouTube video downloads
- Organizes videos into a structured directory hierarchy
- Creates an Excel workbook with download status tracking
- Progress tracking and status updates
- User-friendly GUI interface
- Cross-platform support (Windows, macOS)

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

### Running the Application

1. Make sure you have Python 3.9+ and all dependencies installed:
```bash
pip install -r requirements.txt
python main.py
```

2. In the GUI:
   - Enter the Google Drive URL of your spreadsheet
   - Select the download directory
   - Click "Start Download"

### Notes
- The application will download the spreadsheet and start downloading videos into organized folders
- The output folder structure will be:
  - `<Root Folder>/<Subject>/<Topic>/<Subtopic>/<Video Title>.mp4`
- The application logs progress to `video_downloader.log`
- A processed Excel file will be created with download status for each video

## Building Executables

### Windows (.exe)
1. Open a terminal (cmd or PowerShell) on a Windows machine
2. Install dependencies:
   ```cmd
   pip install -r requirements.txt
   pip install pyinstaller
   ```
3. Build the executable:
   ```cmd
   pyinstaller --onefile --add-data "requirements.txt;." main.py
   ```
   - The .exe will be in the `dist` folder as `main.exe`

### macOS (Binary)
1. Open a terminal on macOS
2. Install dependencies:
   ```sh
   pip3 install -r requirements.txt
   pip3 install pyinstaller
   ```
3. Build the executable:
   ```sh
   pyinstaller --onefile --add-data "requirements.txt:." main.py
   ```
   - The binary will be in the `dist` folder as `main`

### Cross-Platform Note
- You must build the executable on the target OS (Windows for .exe, macOS for binary)
- PyInstaller does not cross-compile
- If you need a Windows .exe but only have a Mac, use a Windows PC or a Windows VM

## Project Structure

```
<Root Directory>/
├── KAUvideos.xlsx
├── KAUvideos_processed.xlsx
├── video_downloader.log
├── <Subject>/
│   ├── <Topic>/
│   │   ├── <Subtopic>/
│   │   │   └── <Video Title>.mp4
```

## Requirements

- Python 3.9+
- Required Python packages (see requirements.txt):
  - pandas: For spreadsheet handling
  - gdown: For Google Drive downloads
  - openpyxl: For Excel file operations
  - requests: For HTTP requests
  - yt-dlp: For YouTube downloads
  - tqdm: For progress bars
  - tkinter: For GUI (usually comes with Python)

## License

MIT License 