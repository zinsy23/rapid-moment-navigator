import os
import re
import glob
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
from pathlib import Path

class RapidMomentNavigator:
    def __init__(self, root):
        self.root = root
        self.root.title("Rapid Moment Navigator")
        self.root.geometry("800x600")
        
        # Map to store relationship between subtitle files and video files
        self.subtitle_to_video_map = {}
        
        # Create main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create search frame
        self.search_frame = ttk.Frame(self.main_frame)
        self.search_frame.pack(fill="x", padx=5, pady=5)
        
        # Create show selection dropdown
        ttk.Label(self.search_frame, text="Show:").pack(side="left", padx=5)
        self.show_var = tk.StringVar()
        self.show_dropdown = ttk.Combobox(self.search_frame, textvariable=self.show_var, state="readonly", width=30)
        self.show_dropdown.pack(side="left", padx=5)
        
        # Create search entry and button
        ttk.Label(self.search_frame, text="Search:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", self.search_subtitles)
        
        self.search_button = ttk.Button(self.search_frame, text="Find All", command=self.search_subtitles)
        self.search_button.pack(side="left", padx=5)
        
        # Create results frame
        self.results_frame = ttk.LabelFrame(self.main_frame, text="Search Results")
        self.results_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create scrolled text widget for results
        self.results_text = scrolledtext.ScrolledText(self.results_frame, wrap=tk.WORD, width=40, height=10)
        self.results_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.results_text.config(state="disabled")
        self.results_text.tag_configure("timecode", foreground="blue", underline=True)
        self.results_text.tag_configure("file_header", foreground="green", font=("TkDefaultFont", 10, "bold"))
        self.results_text.tag_bind("timecode", "<Button-1>", self.on_timecode_click)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief="sunken", anchor="w")
        self.status_bar.pack(fill="x", padx=5, pady=5)
        
        # Initialize the application
        self.load_shows()
        self.map_subtitles_to_videos()
        
        # Store the search results for later reference
        self.search_results = []
    
    def load_shows(self):
        """Load the available shows from the directory structure"""
        shows = [d for d in os.listdir() if os.path.isdir(d) and not d.startswith('.')]
        self.show_dropdown['values'] = shows
        if shows:
            self.show_dropdown.current(0)
    
    def map_subtitles_to_videos(self):
        """Map subtitle files to their corresponding video files"""
        self.status_var.set("Mapping subtitle files to videos...")
        
        for show in self.show_dropdown['values']:
            # Find all subtitle files
            subtitle_files = []
            for ext in ['.srt', '.txt']:
                subtitle_path = os.path.join(show, 'Subtitles')
                if os.path.exists(subtitle_path):
                    subtitle_files.extend(glob.glob(os.path.join(subtitle_path, f'*{ext}')))
            
            # Find all video files
            video_files = []
            for season_dir in [d for d in os.listdir(os.path.join(show)) if d.startswith('Season')]:
                season_path = os.path.join(show, season_dir)
                if os.path.exists(season_path):
                    video_files.extend(glob.glob(os.path.join(season_path, '*.mp4')))
            
            # Map subtitles to videos based on filename
            for subtitle_file in subtitle_files:
                subtitle_basename = os.path.basename(subtitle_file)
                # Remove extension
                subtitle_name = os.path.splitext(subtitle_basename)[0]
                
                # For SRT files that have .mp4.srt extension, we need to extract the true base name
                if subtitle_name.endswith('.mp4'):
                    subtitle_name = subtitle_name[:-4]  # Remove '.mp4'
                
                # Try to find matching video file
                for video_file in video_files:
                    video_basename = os.path.basename(video_file)
                    video_name = os.path.splitext(video_basename)[0]
                    
                    if subtitle_name == video_name or subtitle_basename.startswith(video_name):
                        self.subtitle_to_video_map[subtitle_file] = video_file
                        break
        
        self.status_var.set(f"Ready. Mapped {len(self.subtitle_to_video_map)} subtitle files to videos.")
    
    def search_subtitles(self, event=None):
        """Search for keywords in subtitle files"""
        keyword = self.search_var.get().strip()
        selected_show = self.show_var.get()
        
        if not keyword:
            self.status_var.set("Please enter a search keyword.")
            return
            
        if not selected_show:
            self.status_var.set("Please select a show.")
            return
        
        # Clear previous results
        self.results_text.config(state="normal")
        self.results_text.delete(1.0, tk.END)
        self.search_results = []
        
        # Start search in a separate thread to keep UI responsive
        threading.Thread(target=self._search_thread, args=(keyword, selected_show)).start()
    
    def _search_thread(self, keyword, selected_show):
        """Thread function to handle the search"""
        self.status_var.set(f"Searching for '{keyword}' in {selected_show}...")
        
        subtitle_path = os.path.join(selected_show, 'Subtitles')
        if not os.path.exists(subtitle_path):
            self.status_var.set(f"Subtitle directory not found for {selected_show}")
            return
        
        subtitle_files = []
        for ext in ['.srt', '.txt']:
            subtitle_files.extend(glob.glob(os.path.join(subtitle_path, f'*{ext}')))
        
        # Store the results count
        total_results = 0
        
        # Process each subtitle file
        for subtitle_file in sorted(subtitle_files):
            file_results = []
            
            try:
                with open(subtitle_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # Parse SRT format
                pattern = r'(\d+)\n(\d{2}:\d{2}:\d{2},\d{3})\s+-->\s+(\d{2}:\d{2}:\d{2},\d{3})\n((?:.+\n)+?)(?:\n|$)'
                matches = re.finditer(pattern, content)
                
                for match in matches:
                    subtitle_num = match.group(1)
                    start_time = match.group(2)
                    end_time = match.group(3)
                    text = match.group(4).strip()
                    
                    # Remove HTML tags for searching
                    clean_text = re.sub(r'<[^>]+>', '', text)
                    
                    if keyword.lower() in clean_text.lower():
                        # Convert comma separator to period for MPC
                        mpc_start_time = start_time.replace(',', '.')
                        
                        # Store result
                        result = {
                            'file': subtitle_file,
                            'num': subtitle_num,
                            'start_time': start_time,
                            'end_time': end_time,
                            'text': text,
                            'mpc_start_time': mpc_start_time
                        }
                        file_results.append(result)
                        self.search_results.append(result)
                        total_results += 1
            
            except Exception as e:
                self.status_var.set(f"Error processing {subtitle_file}: {e}")
                continue
            
            # If there are results for this file, display them in the UI
            if file_results:
                self.results_text.config(state="normal")
                
                # Add file header
                file_basename = os.path.basename(subtitle_file)
                self.results_text.insert(tk.END, f"File: {file_basename}\n", "file_header")
                
                # Add each result
                for result in file_results:
                    # Display the timecode and text
                    tag_index = self.results_text.index(tk.END)
                    self.results_text.insert(tk.END, f"{result['start_time']} --> {result['end_time']}\n", "timecode")
                    # Store the result index with the tag
                    result['tag_index'] = tag_index
                    
                    # Display the text without HTML tags
                    clean_text = re.sub(r'<[^>]+>', '', result['text'])
                    self.results_text.insert(tk.END, f"{clean_text}\n\n")
                
                self.results_text.config(state="disabled")
        
        # Update status
        self.status_var.set(f"Found {total_results} matches in {selected_show}")
    
    def on_timecode_click(self, event):
        """Handle click on a timecode to play the video at that time"""
        # Get the index of the clicked text
        index = self.results_text.index(f"@{event.x},{event.y}")
        
        # Find which result was clicked by comparing tag positions
        clicked_result = None
        for result in self.search_results:
            if 'tag_index' in result:
                tag_start = result['tag_index']
                tag_end = self.results_text.index(f"{tag_start} lineend")
                
                # Check if the click was between tag start and end
                if self.results_text.compare(tag_start, "<=", index) and self.results_text.compare(index, "<=", tag_end):
                    clicked_result = result
                    break
        
        if clicked_result:
            subtitle_file = clicked_result['file']
            if subtitle_file in self.subtitle_to_video_map:
                video_file = self.subtitle_to_video_map[subtitle_file]
                self.play_video(video_file, clicked_result['mpc_start_time'])
            else:
                self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def play_video(self, video_file, start_time):
        """Launch Media Player Classic with the video at the specified time"""
        # Convert time to MPC format
        # MPC expects the format hh:mm:ss (without milliseconds)
        mpc_time = start_time.split('.')[0]  # Take only hh:mm:ss part
        
        try:
            # Try to launch MPC-HC
            command = f'start "" "C:\\Program Files\\MPC-HC\\mpc-hc64.exe" "{video_file}" /start {mpc_time}'
            self.status_var.set(f"Opening {os.path.basename(video_file)} at {mpc_time}")
            
            # Execute the command
            subprocess.Popen(command, shell=True)
        except Exception as e:
            self.status_var.set(f"Error launching Media Player Classic: {e}")
            
            # Fall back to default player if MPC fails
            try:
                os.startfile(video_file)
                self.status_var.set(f"Opened {os.path.basename(video_file)} with default player")
            except Exception as e2:
                self.status_var.set(f"Error opening video: {e2}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RapidMomentNavigator(root)
    root.mainloop() 