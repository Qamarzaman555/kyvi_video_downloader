import os
import sys
import pandas as pd
import requests
import gdown
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue

class VideoDownloader:
    def __init__(self):
        self.SPREADSHEET_URL = "https://bit.ly/KAUvideos"
        self.progress_queue = queue.Queue()
        self.setup_ui()

    def setup_ui(self):
        self.root = tk.Tk()
        self.root.title("KAU Video Downloader")
        self.root.geometry("600x400")
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Folder selection
        ttk.Label(main_frame, text="Select Root Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.select_folder).grid(row=0, column=2)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, length=400, mode='determinate', variable=self.progress_var)
        self.progress_bar.grid(row=1, column=0, columnspan=3, pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=2, column=0, columnspan=3, pady=5)
        
        # Start button
        self.start_button = ttk.Button(main_frame, text="Start Download", command=self.start_download)
        self.start_button.grid(row=3, column=0, columnspan=3, pady=10)
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def update_progress(self):
        while True:
            try:
                msg = self.progress_queue.get_nowait()
                if msg == "DONE":
                    self.status_label.config(text="Download Complete!")
                    self.start_button.config(state="normal")
                    break
                elif isinstance(msg, tuple):
                    progress, status = msg
                    self.progress_var.set(progress)
                    self.status_label.config(text=status)
            except queue.Empty:
                self.root.after(100, self.update_progress)
                break

    def download_spreadsheet(self, dest_path):
        try:
            response = requests.get(self.SPREADSHEET_URL)
            response.raise_for_status()
            with open(dest_path, 'wb') as f:
                f.write(response.content)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download spreadsheet: {str(e)}")
            return False

    def process_videos(self, root_folder, spreadsheet_path):
        try:
            df = pd.read_csv(spreadsheet_path)
            total_rows = len(df)
            excel_path = os.path.join(root_folder, "KAUvideos.xlsx")
            
            # Create Excel workbook
            wb = Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            # Process each subject
            subjects = df['Subject'].unique()
            for subject_idx, subject in enumerate(subjects):
                subject_df = df[df['Subject'] == subject]
                ws = wb.create_sheet(title=subject)
                
                # Add headers
                headers = list(subject_df.columns) + ['Local Link']
                for col, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
                
                # Process each video
                for idx, (_, row) in enumerate(subject_df.iterrows()):
                    progress = ((subject_idx * len(subject_df) + idx) / total_rows) * 100
                    self.progress_queue.put((progress, f"Processing: {row['Video Title']}"))
                    
                    # Create folder structure
                    subject_dir = os.path.join(root_folder, row['Subject'], row['Topic'], row['Subtopic'])
                    os.makedirs(subject_dir, exist_ok=True)
                    
                    # Download video
                    video_filename = f"{row['Video Title']}.mp4"
                    video_path = os.path.join(subject_dir, video_filename)
                    
                    if not os.path.exists(video_path):
                        try:
                            gdown.download(row['Google Drive Link'], video_path, quiet=True)
                        except Exception as e:
                            print(f"Failed to download {video_filename}: {str(e)}")
                            continue
                    
                    # Add row to Excel
                    row_data = list(row)
                    rel_path = os.path.relpath(video_path, root_folder)
                    link = f'=HYPERLINK("{rel_path}", "{video_filename}")'
                    ws.append(row_data + [link])
                
                # Adjust column widths
                for col in range(1, len(headers) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 20
            
            wb.save(excel_path)
            self.progress_queue.put("DONE")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process videos: {str(e)}")
            self.progress_queue.put("DONE")

    def start_download(self):
        root_folder = self.folder_path.get()
        if not root_folder:
            messagebox.showerror("Error", "Please select a root folder")
            return
        
        self.start_button.config(state="disabled")
        self.progress_var.set(0)
        self.status_label.config(text="Starting download...")
        
        # Start download in separate thread
        def download_thread():
            spreadsheet_path = os.path.join(root_folder, "KAUvideos.csv")
            if self.download_spreadsheet(spreadsheet_path):
                self.process_videos(root_folder, spreadsheet_path)
        
        thread = threading.Thread(target=download_thread)
        thread.daemon = True
        thread.start()
        
        # Start progress update
        self.update_progress()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = VideoDownloader()
    app.run() 