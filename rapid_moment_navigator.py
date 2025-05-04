import os
import re
import glob
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, Label, Frame
import threading
from pathlib import Path
import sys

class ClickableTimecode(Label):
    """A clickable label widget for timecodes"""
    def __init__(self, parent, timecode, result, callback, **kwargs):
        super().__init__(parent, text=timecode, cursor="hand2", fg="blue", **kwargs)
        self.result = result
        self.callback = callback
        self.bind("<Button-1>", self._on_click)
        # Add underline
        self.config(font=("TkDefaultFont", 10, "underline"))
        
    def _on_click(self, event):
        """Handle click event"""
        self.callback(self.result)

class RapidMomentNavigator:
    def __init__(self, root):
        self.root = root
        self.root.title("Rapid Moment Navigator")
        self.root.geometry("800x600")
        
        # Enable debug mode
        self.debug = True
        
        # Store the script directory for relative path operations
        self.script_dir = os.path.abspath(os.path.dirname(__file__))
        self.debug_print(f"Script directory: {self.script_dir}")
        
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
        
        # Create results frame with canvas for scrolling
        self.results_frame = ttk.LabelFrame(self.main_frame, text="Search Results")
        self.results_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create a canvas with scrollbar for results
        self.results_canvas = tk.Canvas(self.results_frame)
        self.results_scrollbar = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.results_canvas.yview)
        self.results_scrollbar.pack(side="right", fill="y")
        
        self.results_canvas.pack(side="left", fill="both", expand=True)
        self.results_canvas.configure(yscrollcommand=self.results_scrollbar.set)
        
        # Frame inside canvas to hold results
        self.results_container = ttk.Frame(self.results_canvas)
        self.results_container_id = self.results_canvas.create_window((0, 0), window=self.results_container, anchor="nw")
        
        # Configure canvas scrolling
        self.results_container.bind("<Configure>", self._configure_scroll_region)
        self.results_canvas.bind("<Configure>", self._configure_canvas_width)
        
        # Bind mousewheel scrolling
        self.results_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief="sunken", anchor="w")
        self.status_bar.pack(fill="x", padx=5, pady=5)
        
        # Initialize the application
        self.load_shows()
        self.map_subtitles_to_videos()
        
        # Store the search results for later reference
        self.search_results = []
        
        # Debug print
        self.debug_print("Application initialized")
    
    def _configure_scroll_region(self, event):
        """Configure the scroll region of the canvas"""
        self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
    
    def _configure_canvas_width(self, event):
        """Make the canvas width match its container"""
        self.results_canvas.itemconfig(self.results_container_id, width=event.width)
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling"""
        self.results_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def debug_print(self, message):
        """Print debug messages if debug mode is enabled"""
        if self.debug:
            print(f"DEBUG: {message}", flush=True)
    
    def load_shows(self):
        """Load the available shows from the directory structure"""
        shows = [d for d in os.listdir() if os.path.isdir(d) and not d.startswith('.') and d not in ['.git']]
        self.show_dropdown['values'] = shows
        if shows:
            self.show_dropdown.current(0)
        self.debug_print(f"Loaded shows: {shows}")
    
    def map_subtitles_to_videos(self):
        """Map subtitle files to their corresponding video files"""
        self.status_var.set("Mapping subtitle files to videos...")
        
        for show in self.show_dropdown['values']:
            # Find all subtitle files (focus on .srt files)
            subtitle_files = []
            subtitle_path = os.path.join(show, 'Subtitles')
            
            # Check if there's a dedicated Subtitles folder
            if os.path.exists(subtitle_path):
                subtitle_files.extend(glob.glob(os.path.join(subtitle_path, '*.srt')))
            
            # Find all video files anywhere in the show directory
            video_files = []
            
            # Walk through the entire directory structure to find all mp4 files
            for root, dirs, files in os.walk(os.path.join(show)):
                for file in files:
                    if file.endswith('.mp4'):
                        video_files.append(os.path.join(root, file))
            
            self.debug_print(f"Found {len(subtitle_files)} subtitle files and {len(video_files)} video files for {show}")
            
            # Map subtitles to videos based on similarity of filenames
            for subtitle_file in subtitle_files:
                subtitle_basename = os.path.basename(subtitle_file)
                # Remove extension
                subtitle_name = os.path.splitext(subtitle_basename)[0]
                
                # For SRT files that have .mp4.srt extension, extract true base name
                if subtitle_name.endswith('.mp4'):
                    subtitle_name = subtitle_name[:-4]  # Remove '.mp4'
                
                # Try exact matches first, then partial matches
                matched = False
                
                # First pass: look for exact filename matches (without extensions)
                for video_file in video_files:
                    video_basename = os.path.basename(video_file)
                    video_name = os.path.splitext(video_basename)[0]
                    
                    if subtitle_name == video_name:
                        self.subtitle_to_video_map[subtitle_file] = video_file
                        self.debug_print(f"Exact match: {subtitle_basename} -> {video_basename}")
                        matched = True
                        break
                
                # If no exact match, try partial matches
                if not matched:
                    # Clean up filenames for better matching
                    clean_subtitle_name = self._clean_filename(subtitle_name)
                    
                    for video_file in video_files:
                        video_basename = os.path.basename(video_file)
                        video_name = os.path.splitext(video_basename)[0]
                        clean_video_name = self._clean_filename(video_name)
                        
                        # Check if the cleaned names match or one contains the other
                        if (clean_subtitle_name == clean_video_name or
                            clean_subtitle_name in clean_video_name or
                            clean_video_name in clean_subtitle_name):
                            self.subtitle_to_video_map[subtitle_file] = video_file
                            self.debug_print(f"Partial match: {subtitle_basename} -> {video_basename}")
                            matched = True
                            break
        
        self.debug_print(f"Mapped {len(self.subtitle_to_video_map)} subtitle files to videos")
        self.status_var.set(f"Ready. Mapped {len(self.subtitle_to_video_map)} subtitle files to videos.")
    
    def _clean_filename(self, filename):
        """Clean a filename to improve matching chances"""
        # Convert to lowercase
        filename = filename.lower()
        # Remove common separators
        filename = filename.replace('_', ' ').replace('-', ' ').replace('.', ' ')
        # Remove common words/patterns that might differ between subtitle and video filenames
        remove_patterns = ['disc', 'season', 'title', 'episode', 's0', 'e0', 'x0']
        for pattern in remove_patterns:
            filename = filename.replace(pattern, '')
        # Remove numbers that might be disc/episode numbers
        filename = re.sub(r'\b\d{1,2}\b', '', filename)
        # Remove extra spaces
        filename = re.sub(r'\s+', ' ', filename).strip()
        return filename
    
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
        for widget in self.results_container.winfo_children():
            widget.destroy()
        
        self.search_results = []
        
        self.debug_print(f"Searching for '{keyword}' in {selected_show}")
        
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
                self.debug_print(f"Error processing {subtitle_file}: {e}")
                self.status_var.set(f"Error processing {subtitle_file}: {e}")
                continue
            
            # If there are results for this file, display them in the UI
            if file_results:
                # Use main thread to update the UI
                self.root.after(0, self._update_results_ui, subtitle_file, file_results)
        
        # Update status
        self.debug_print(f"Found {total_results} matches in {selected_show}")
        self.root.after(0, lambda: self.status_var.set(f"Found {total_results} matches in {selected_show}"))
    
    def _update_results_ui(self, subtitle_file, file_results):
        """Update the UI with search results (called from main thread)"""
        # Add file header
        file_basename = os.path.basename(subtitle_file)
        file_header = ttk.Label(self.results_container, text=f"File: {file_basename}", 
                               font=("TkDefaultFont", 10, "bold"), foreground="green")
        file_header.pack(anchor="w", padx=5, pady=2)
        
        # Add each result
        for result in file_results:
            # Create a frame for this result
            result_frame = ttk.Frame(self.results_container)
            result_frame.pack(fill="x", padx=5, pady=2, anchor="w")
            
            # Create clickable timecode label
            timecode_text = f"{result['start_time']} --> {result['end_time']}"
            timecode_label = ClickableTimecode(
                result_frame, 
                timecode_text, 
                result, 
                self._handle_timecode_click
            )
            timecode_label.pack(anchor="w")
            
            # Add text label
            text_label = ttk.Label(result_frame, text=result['clean_text'], wraplength=700)
            text_label.pack(anchor="w", padx=10)
            
            # Add some space after each result
            ttk.Separator(self.results_container, orient="horizontal").pack(fill="x", pady=5)
            
            self.debug_print(f"Added clickable timecode for {timecode_text}")
        
        self.debug_print(f"UI updated with {len(file_results)} results from {file_basename}")
    
    def _handle_timecode_click(self, result):
        """Process a click on a timecode tag"""
        subtitle_file = result['file']
        self.debug_print(f"Handling click for subtitle file: {subtitle_file}")
        
        if subtitle_file in self.subtitle_to_video_map:
            video_file = self.subtitle_to_video_map[subtitle_file]
            self.debug_print(f"Found matching video file: {video_file}")
            self.play_video(video_file, result['mpc_start_time'])
            self.status_var.set(f"Opening {os.path.basename(video_file)} at {result['mpc_start_time']}")
        else:
            self.debug_print(f"No matching video file found for {os.path.basename(subtitle_file)}")
            self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def get_absolute_path(self, relative_path):
        """Convert a relative path to an absolute path based on script directory"""
        abs_path = os.path.normpath(os.path.join(self.script_dir, relative_path))
        self.debug_print(f"Converting relative path: {relative_path} to absolute: {abs_path}")
        return abs_path
    
    def play_video(self, video_file, start_time):
        """Launch Media Player Classic with the video at the specified time"""
        try:
            # Convert the relative video path to absolute
            abs_video_path = self.get_absolute_path(video_file)
            
            # Construct the command for MPC-HC
            # Documentation says correct parameter is /startpos hh:mm:ss
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
                        
            self.debug_print(f"Using MPC path: {mpc_path}")
            
            # Try method 1: Using /startpos as a separate parameter
            command = [mpc_path, abs_video_path, "/startpos", start_time]
            self.debug_print(f"Executing command: {command}")
            subprocess.Popen(command)
            
        except Exception as e:
            self.debug_print(f"Error launching Media Player Classic: {str(e)}")
            self.status_var.set(f"Error launching Media Player Classic: {e}")
            
            try:
                # Try method 2: Using shell=True with space-separated arguments
                self.debug_print("Trying alternate launch method with shell=True")
                abs_video_path = self.get_absolute_path(video_file)
                command = f'start "" "{mpc_path}" "{abs_video_path}" /startpos {start_time}'
                self.debug_print(f"Shell command: {command}")
                subprocess.Popen(command, shell=True)
            except Exception as e2:
                try:
                    # Try method 3: Using shell=True with parameter combined with value
                    self.debug_print("Trying another alternative launch method")
                    command = f'start "" "{mpc_path}" "{abs_video_path}" /startpos={start_time}'
                    self.debug_print(f"Shell command: {command}")
                    subprocess.Popen(command, shell=True)
                except Exception as e3:
                    self.debug_print(f"Error with all launch methods, falling back to default player")
                    
                    # Fall back to default player if MPC fails
                    try:
                        abs_video_path = self.get_absolute_path(video_file)
                        os.startfile(abs_video_path)
                        self.status_var.set(f"Opened {os.path.basename(video_file)} with default player")
                    except Exception as e4:
                        self.debug_print(f"Error opening with default player: {str(e4)}")
                        self.status_var.set(f"Error opening video: {e4}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RapidMomentNavigator(root)
    # Force debug output to be flushed immediately
    sys.stdout.flush()
    root.mainloop() 