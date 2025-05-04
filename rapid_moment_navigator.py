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
        
        # Store tag information for click handling
        self.tag_mapping = {}
        
        # Bind click events to the entire text widget
        self.results_text.bind("<ButtonRelease-1>", self.on_text_click)
        
        # Make the cursor change to a hand when hovering over timecodes
        self.results_text.tag_bind("timecode", "<Enter>", lambda e: self.results_text.config(cursor="hand2"))
        self.results_text.tag_bind("timecode", "<Leave>", lambda e: self.results_text.config(cursor=""))
        
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
        self.tag_mapping = {}
        
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
                        
                        # Store result with only the hh:mm:ss part (no milliseconds) for MPC
                        mpc_time_format = mpc_start_time.split('.')[0]
                        
                        # Store result
                        result = {
                            'file': subtitle_file,
                            'num': subtitle_num,
                            'start_time': start_time,
                            'end_time': end_time,
                            'text': text,
                            'mpc_start_time': mpc_time_format,
                            'clean_text': clean_text
                        }
                        file_results.append(result)
                        self.search_results.append(result)
                        total_results += 1
            
            except Exception as e:
                self.status_var.set(f"Error processing {subtitle_file}: {e}")
                continue
            
            # If there are results for this file, display them in the UI
            if file_results:
                # Use main thread to update the UI
                self.root.after(0, self._update_results_ui, subtitle_file, file_results)
        
        # Update status
        self.root.after(0, lambda: self.status_var.set(f"Found {total_results} matches in {selected_show}"))
    
    def _update_results_ui(self, subtitle_file, file_results):
        """Update the UI with search results (called from main thread)"""
        self.results_text.config(state="normal")
        
        # Add file header
        file_basename = os.path.basename(subtitle_file)
        self.results_text.insert(tk.END, f"File: {file_basename}\n", "file_header")
        
        # Add each result
        for result in file_results:
            # Get current position for tagging
            start_pos = self.results_text.index(tk.END)
            
            # Insert the timecode as a clickable tag
            timecode_text = f"{result['start_time']} --> {result['end_time']}"
            self.results_text.insert(tk.END, timecode_text, "timecode")
            
            # Get ending position of the inserted tag
            end_pos = self.results_text.index(tk.END)
            
            # Store the tag range with the result for click handling
            tag_id = f"timecode-{len(self.tag_mapping)}"
            self.results_text.tag_add(tag_id, start_pos, end_pos)
            self.tag_mapping[tag_id] = result
            
            # Insert the rest of the text
            self.results_text.insert(tk.END, "\n")
            self.results_text.insert(tk.END, f"{result['clean_text']}\n\n")
        
        self.results_text.config(state="disabled")
    
    def on_text_click(self, event):
        """Handle click events anywhere in the text widget"""
        # Get click position
        index = self.results_text.index(f"@{event.x},{event.y}")
        
        # Check if the click is on a timecode tag
        for tag_id, result in self.tag_mapping.items():
            tag_ranges = self.results_text.tag_ranges(tag_id)
            if tag_ranges:  # Check if the tag exists
                start = tag_ranges[0]
                end = tag_ranges[1]
                
                # Check if the click is within the tag range
                if self.results_text.compare(start, "<=", index) and self.results_text.compare(index, "<", end):
                    self._handle_timecode_click(result)
                    return
        
        # If we get here, the click wasn't on a timecode
        pass
    
    def _handle_timecode_click(self, result):
        """Process a click on a timecode tag"""
        subtitle_file = result['file']
        if subtitle_file in self.subtitle_to_video_map:
            video_file = self.subtitle_to_video_map[subtitle_file]
            self.play_video(video_file, result['mpc_start_time'])
            self.status_var.set(f"Opening {os.path.basename(video_file)} at {result['mpc_start_time']}")
        else:
            self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def play_video(self, video_file, start_time):
        """Launch Media Player Classic with the video at the specified time"""
        try:
            # Construct the command for MPC-HC
            # MPC-HC accepts the /start parameter in hh:mm:ss format
            mpc_path = "C:\\Program Files\\MPC-HC\\mpc-hc64.exe"
            
            # Check if default MPC path exists
            if not os.path.exists(mpc_path):
                # Try alternative paths
                alternative_paths = [
                    "C:\\Program Files (x86)\\MPC-HC\\mpc-hc.exe",
                    "C:\\Program Files (x86)\\K-Lite Codec Pack\\MPC-HC64\\mpc-hc64.exe",
                    "C:\\Program Files\\K-Lite Codec Pack\\MPC-HC64\\mpc-hc64.exe"
                ]
                
                for path in alternative_paths:
                    if os.path.exists(path):
                        mpc_path = path
                        break
            
            # Construct and execute the command
            command = f'"{mpc_path}" "{video_file}" /start {start_time}'
            
            # Log the command for debugging
            print(f"Executing command: {command}")
            
            # Use subprocess directly without shell=True for better security
            subprocess.Popen(command)
            
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