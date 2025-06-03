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
git clone https://github.com/yourusername/kyvi_video_downloader.git
cd kyvi_video_downloader
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Desktop (Windows/macOS)

1. Run the application:
```bash
python main.py
```

2. Select a root folder where videos will be downloaded
3. Click "Start Download" to begin the process
4. Wait for the download to complete
5. Find the organized videos and Excel file in your selected folder

### Mobile (Android/iOS)

*Coming soon*

## Building Executables

### Windows
```bash
pyinstaller --onefile --windowed main.py
```

### macOS
```bash
python setup.py py2app
```

### Android
```bash
buildozer init
buildozer -v android debug
```

### iOS
```bash
toolchain build python3 kivy
toolchain create kyvi_video_downloader .
```

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