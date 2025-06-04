import os
import pandas as pd
import requests
import gdown
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import sys
import logging
from tqdm import tqdm
import re
import yt_dlp
from urllib.parse import urlparse, parse_qs, unquote
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from threading import Thread

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('video_downloader.log'),
        logging.StreamHandler()
    ]
)

def is_youtube_url(url):
    """Check if the URL is a YouTube link."""
    return 'youtube.com' in url or 'youtu.be' in url

def download_from_youtube(url, output_path):
    """Download a video from YouTube using yt-dlp."""
    try:
        ydl_opts = {
            'format': 'best[ext=mp4]',
            'outtmpl': output_path,
            'quiet': True,
            'no_warnings': True
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 1024:
            logging.info(f"Successfully downloaded from YouTube: {output_path}")
            return True
        else:
            logging.error(f"Downloaded YouTube file is too small or doesn't exist: {output_path}")
            if os.path.exists(output_path):
                os.remove(output_path)
            return False
            
    except Exception as e:
        logging.error(f"Failed to download from YouTube {output_path}: {str(e)}")
        if os.path.exists(output_path):
            os.remove(output_path)
        return False

def extract_file_id(url):
    """Extract file ID from Google Drive URL."""
    if not url or not isinstance(url, str):
        return None
        
    try:
        # Handle different Google Drive URL formats
        if 'drive.google.com/file/d/' in url:
            file_id = url.split('/file/d/')[1].split('/')[0]
        elif 'drive.google.com/open?id=' in url:
            file_id = parse_qs(urlparse(url).query)['id'][0]
        elif 'drive.google.com/uc?id=' in url:
            file_id = parse_qs(urlparse(url).query)['id'][0]
        else:
            logging.warning(f"Could not extract file ID from URL: {url}")
            return None
            
        return file_id
    except Exception as e:
        logging.error(f"Error extracting file ID from {url}: {str(e)}")
        return None

def download_from_drive(url, output_path):
    """Download a file from Google Drive."""
    try:
        file_id = extract_file_id(url)
        if not file_id:
            return False
            
        # Construct the download URL
        download_url = f"https://drive.google.com/uc?id={file_id}"
        
        # Make the request with a session to handle cookies
        session = requests.Session()
        response = session.get(download_url, stream=True)
        
        # Check if the response is valid
        if response.status_code != 200:
            logging.error(f"Failed to download {url}. Status code: {response.status_code}")
            return False
            
        # Get the file size
        total_size = int(response.headers.get('content-length', 0))
        if total_size == 0:
            logging.error(f"Failed to get file size for {url}")
            return False
            
        # Download the file
        with open(output_path, 'wb') as f:
            downloaded = 0
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    # Log progress every 10%
                    if total_size > 0:
                        progress = (downloaded / total_size) * 100
                        if int(progress) % 10 == 0:
                            logging.info(f"Download progress for {os.path.basename(output_path)}: {int(progress)}%")
        
        # Verify the download
        if os.path.getsize(output_path) == 0:
            logging.error(f"Downloaded file is empty: {output_path}")
            os.remove(output_path)
            return False
            
        logging.info(f"Successfully downloaded: {output_path}")
        return True
        
    except Exception as e:
        logging.error(f"Error downloading {url}: {str(e)}")
        if os.path.exists(output_path):
            os.remove(output_path)
        return False

def download_video(url, output_path):
    """Download video from either YouTube or Google Drive."""
    if is_youtube_url(url):
        return download_from_youtube(url, output_path)
    else:
        return download_from_drive(url, output_path)

def download_spreadsheet_xlsx(url, dest_path):
    """Download the spreadsheet as an Excel file from the given URL."""
    logging.info("Downloading spreadsheet as Excel (.xlsx)...")
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(response.content)
        logging.info("Spreadsheet downloaded successfully!")
        return True
    except Exception as e:
        logging.error(f"Error downloading spreadsheet: {str(e)}")
        return False

def find_column(df, possible_keywords):
    """Find a column in the DataFrame if any keyword is contained in the column name (case-insensitive, strip whitespace)."""
    for col in df.columns:
        col_norm = str(col).strip().lower()
        for keyword in possible_keywords:
            keyword_norm = keyword.strip().lower()
            if keyword_norm in col_norm:
                return col
    return None

def process_videos_all_sheets(root_folder, spreadsheet_path):
    """Process all sheets in the Excel file and download videos."""
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(spreadsheet_path)
        sheet_names = excel_file.sheet_names
        
        # Skip the first few sheets that don't contain video data
        skip_sheets = ['Introduction - تعارف', 'Watching Duration - دیکھنے کا د', 'Review Allocation']
        sheet_names = [sheet for sheet in sheet_names if sheet not in skip_sheets]
        
        # Process each sheet
        for sheet_name in sheet_names:
            logging.info(f"Processing sheet: {sheet_name}")
            
            # Read the sheet
            df = pd.read_excel(spreadsheet_path, sheet_name=sheet_name)
            
            # Process each video
            for index, row in df.iterrows():
                try:
                    video_title = row['Video Title']
                    if pd.isna(video_title):
                        continue
                        
                    # Get the Google Drive URL
                    drive_url = row['Google Drive URL']
                    if pd.isna(drive_url) or not isinstance(drive_url, str) or not drive_url.startswith('http'):
                        logging.warning(f"Skipping invalid Google Drive URL for video: {video_title}")
                        continue
                    
                    # Get subject, topic, and subtopic
                    subject = row.get('Subject', sheet_name)
                    topic = row.get('Topic', '')
                    subtopic = row.get('Sub Topic', '')  # Updated column name
                    
                    # Create the folder structure
                    folder_path = os.path.join(root_folder, subject)
                    if topic and not pd.isna(topic):
                        folder_path = os.path.join(folder_path, topic)
                    if subtopic and not pd.isna(subtopic):
                        folder_path = os.path.join(folder_path, subtopic)
                    
                    # Create all necessary directories
                    os.makedirs(folder_path, exist_ok=True)
                    
                    # Create a safe filename
                    safe_filename = "".join(c for c in video_title if c.isalnum() or c in (' ', '-', '_')).strip()
                    safe_filename = safe_filename.replace(' ', '_')
                    output_path = os.path.join(folder_path, f"{safe_filename}.mp4")
                    
                    # Skip if already downloaded
                    if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                        logging.info(f"Video already exists: {output_path}")
                        continue
                    
                    # Download the video
                    logging.info(f"Downloading: {video_title}")
                    if download_from_drive(drive_url, output_path):
                        # Update the status in the DataFrame
                        df.at[index, 'Download Status'] = 'Downloaded'
                    else:
                        df.at[index, 'Download Status'] = 'Failed'
                        
                    # Add a small delay between downloads
                    time.sleep(1)
                    
                except Exception as e:
                    logging.error(f"Error processing video {video_title}: {str(e)}")
                    df.at[index, 'Download Status'] = 'Error'
            
            # Save the processed sheet
            output_excel = f"{os.path.splitext(spreadsheet_path)[0]}_processed.xlsx"
            with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a' if os.path.exists(output_excel) else 'w') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
        logging.info("\nAll videos have been downloaded and organized successfully!")
        
    except Exception as e:
        logging.error(f"Error processing spreadsheet: {str(e)}")
        raise

class VideoDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("KAU Video Downloader")
        self.root.geometry("600x400")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # URL Input
        ttk.Label(main_frame, text="Enter Google Drive URL:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.url_entry = ttk.Entry(main_frame, width=50)
        self.url_entry.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Directory Selection
        ttk.Label(main_frame, text="Select Download Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.dir_entry = ttk.Entry(main_frame, width=40)
        self.dir_entry.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_directory).grid(row=3, column=1, padx=5)
        
        # Progress Bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Status Label
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=5)
        
        # Execute Button
        self.execute_button = ttk.Button(main_frame, text="Start Download", command=self.start_download)
        self.execute_button.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        
    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()
    
    def start_download(self):
        url = self.url_entry.get().strip()
        directory = self.dir_entry.get().strip()
        
        if not url:
            messagebox.showerror("Error", "Please enter a URL")
            return
        
        if not directory:
            messagebox.showerror("Error", "Please select a download directory")
            return
        
        # Disable the execute button while downloading
        self.execute_button.config(state='disabled')
        
        # Start download in a separate thread
        Thread(target=self.download_process, args=(url, directory), daemon=True).start()
    
    def download_process(self, url, directory):
        try:
            self.update_status("Starting download process...")
            
            # Validate URL
            response = requests.head(url)
            if response.status_code != 200:
                self.update_status(f"Error: Invalid URL (Status code: {response.status_code})")
                messagebox.showerror("Error", f"Invalid URL (Status code: {response.status_code})")
                self.execute_button.config(state='normal')
                return
            
            # Create downloader instance with status callback
            downloader = VideoDownloader(directory, status_callback=self.update_status)
            
            # Start download
            self.update_status("Downloading videos...")
            downloader.download_videos(url)
            
            self.update_status("Download completed successfully!")
            messagebox.showinfo("Success", "All videos have been downloaded successfully!")
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}")
            messagebox.showerror("Error", str(e))
        finally:
            self.execute_button.config(state='normal')

class VideoDownloader:
    def __init__(self, root_folder, status_callback=None):
        self.root_folder = os.path.abspath(root_folder)
        self.status_callback = status_callback
        logging.info(f"Using root folder: {self.root_folder}")
        
        # Create root folder if it doesn't exist
        os.makedirs(self.root_folder, exist_ok=True)
        
        # Initialize Excel writer
        self.excel_path = os.path.join(self.root_folder, 'downloaded_videos.xlsx')
        self.writer = pd.ExcelWriter(self.excel_path, engine='openpyxl')
        self.downloaded_videos = []

    def update_status(self, message):
        if self.status_callback:
            self.status_callback(message)
        logging.info(message)

    def download_videos(self, url):
        try:
            self.update_status("Validating URL...")
            
            # Check if it's a Google Drive URL (more comprehensive check)
            if not any(pattern in url.lower() for pattern in ['drive.google.com', 'docs.google.com']):
                raise ValueError("Please provide a valid Google Drive URL")
            
            # Convert to direct download URL if it's a sharing URL
            try:
                # Handle different URL formats
                if '/spreadsheets/d/' in url:
                    # Extract the file ID from the URL
                    file_id = url.split('/spreadsheets/d/')[1].split('/')[0]
                    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
                elif '/file/d/' in url:
                    file_id = url.split('/file/d/')[1].split('/')[0]
                    url = f"https://drive.google.com/uc?export=download&id={file_id}"
                elif '/open?id=' in url:
                    file_id = url.split('id=')[1].split('&')[0]
                    url = f"https://drive.google.com/uc?export=download&id={file_id}"
            except Exception as e:
                logging.error(f"Error parsing URL: {str(e)}")
                raise ValueError("Could not parse the Google Drive URL. Please make sure it's a valid sharing link.")
            
            # Download the Excel file
            self.update_status("Downloading spreadsheet...")
            response = requests.get(url)
            response.raise_for_status()
            
            # Save the Excel file
            excel_path = os.path.join(self.root_folder, 'KAUvideos.xlsx')
            with open(excel_path, 'wb') as f:
                f.write(response.content)
            
            # Read all sheets
            self.update_status("Reading spreadsheet...")
            excel_file = pd.ExcelFile(excel_path)
            sheet_names = excel_file.sheet_names
            
            # Skip the first few sheets that don't contain video data
            skip_sheets = ['Introduction - تعارف', 'Watching Duration - دیکھنے کا د', 'Review Allocation']
            sheet_names = [sheet for sheet in sheet_names if sheet not in skip_sheets]
            
            # Process each sheet
            for sheet_name in sheet_names:
                self.update_status(f"Processing sheet: {sheet_name}")
                
                # Read the sheet
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                
                # Process each video
                for index, row in df.iterrows():
                    try:
                        video_title = row['Video Title']
                        if pd.isna(video_title):
                            continue
                            
                        # Get the Google Drive URL
                        drive_url = row['Google Drive URL']
                        if pd.isna(drive_url) or not isinstance(drive_url, str) or not drive_url.startswith('http'):
                            logging.warning(f"Skipping invalid Google Drive URL for video: {video_title}")
                            continue
                        
                        # Get subject, topic, and subtopic
                        subject = row.get('Subject', sheet_name)
                        topic = row.get('Topic', '')
                        subtopic = row.get('Sub Topic', '')  # Updated column name
                        
                        # Create the folder structure
                        folder_path = os.path.join(self.root_folder, subject)
                        if topic and not pd.isna(topic):
                            folder_path = os.path.join(folder_path, topic)
                        if subtopic and not pd.isna(subtopic):
                            folder_path = os.path.join(folder_path, subtopic)
                        
                        # Create all necessary directories
                        os.makedirs(folder_path, exist_ok=True)
                        
                        # Create a safe filename
                        safe_filename = "".join(c for c in video_title if c.isalnum() or c in (' ', '-', '_')).strip()
                        safe_filename = safe_filename.replace(' ', '_')
                        output_path = os.path.join(folder_path, f"{safe_filename}.mp4")
                        
                        # Skip if already downloaded
                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                            self.update_status(f"Video already exists: {video_title}")
                            continue
                        
                        # Download the video
                        self.update_status(f"Downloading: {video_title}")
                        if download_from_drive(drive_url, output_path):
                            # Update the status in the DataFrame
                            df.at[index, 'Download Status'] = 'Downloaded'
                            self.downloaded_videos.append({
                                'Subject': subject,
                                'Topic': topic,
                                'Subtopic': subtopic,
                                'Video Title': video_title,
                                'Local Path': output_path
                            })
                        else:
                            df.at[index, 'Download Status'] = 'Failed'
                            
                        # Add a small delay between downloads
                        time.sleep(1)
                        
                    except Exception as e:
                        logging.error(f"Error processing video {video_title}: {str(e)}")
                        df.at[index, 'Download Status'] = 'Error'
                
                # Save the processed sheet
                output_excel = f"{os.path.splitext(excel_path)[0]}_processed.xlsx"
                with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a' if os.path.exists(output_excel) else 'w') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.update_status("All videos have been downloaded and organized successfully!")
            
        except Exception as e:
            logging.error(f"Error downloading videos: {str(e)}")
            raise

def main():
    root = tk.Tk()
    app = VideoDownloaderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 