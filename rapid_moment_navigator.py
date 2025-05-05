import os
import re
import glob
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, Label, Frame, filedialog
import threading
from pathlib import Path
import sys
import argparse
import json

# Constants
PREFS_FILENAME = "rapid_navigator_prefs.json"
DEFAULT_PREFS = {
    "directories": [],
    "exclude_current_dir": False
}

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
    def __init__(self, root, debug=False):
        self.root = root
        self.root.title("Rapid Moment Navigator")
        self.root.geometry("800x600")
        
        # Debug mode setting
        self.debug = debug
        
        # Store the script directory for relative path operations
        self.script_dir = os.path.abspath(os.path.dirname(__file__))
        self.debug_print(f"Script directory: {self.script_dir}")
        
        # Load preferences from file
        self.preferences = self.load_preferences()
        
        # Map to store relationship between subtitle files and video files
        self.subtitle_to_video_map = {}
        
        # Map to store relationship between show names and full paths
        self.show_name_to_path_map = {}
        
        # Create main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create directory management frame
        self.dir_frame = ttk.LabelFrame(self.main_frame, text="Media Directories")
        self.dir_frame.pack(fill="x", padx=5, pady=5)
        
        # Directory listbox with scrollbar
        self.dir_list_frame = ttk.Frame(self.dir_frame)
        self.dir_list_frame.pack(fill="x", padx=5, pady=5, side="left", expand=True)
        
        self.dir_listbox = tk.Listbox(self.dir_list_frame, height=3, width=50)
        self.dir_listbox.pack(side="left", fill="both", expand=True)
        
        dir_scrollbar = ttk.Scrollbar(self.dir_list_frame, orient="vertical", command=self.dir_listbox.yview)
        dir_scrollbar.pack(side="right", fill="y")
        self.dir_listbox.configure(yscrollcommand=dir_scrollbar.set)
        
        # Populate the directory listbox
        self.update_directory_listbox()
        
        # Directory buttons frame
        self.dir_btn_frame = ttk.Frame(self.dir_frame)
        self.dir_btn_frame.pack(fill="y", padx=5, pady=5, side="right")
        
        # Add directory button
        self.add_dir_btn = ttk.Button(self.dir_btn_frame, text="Add Directory", command=self.add_directory)
        self.add_dir_btn.pack(fill="x", padx=5, pady=2)
        
        # Remove directory button
        self.remove_dir_btn = ttk.Button(self.dir_btn_frame, text="Remove Directory", command=self.remove_directory)
        self.remove_dir_btn.pack(fill="x", padx=5, pady=2)
        
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
        
        # Use a direct binding approach for Ctrl+Backspace without KeyRelease complication
        self.search_entry.bind("<Control-BackSpace>", self._ctrl_backspace_handler)
        
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
        shows_paths = []
        self.show_name_to_path_map = {}  # Clear the mapping
        
        # Include current directory only if not excluded
        current_dir = self.get_current_directory()
        search_dirs = []
        
        if not self.preferences.get("exclude_current_dir", False):
            search_dirs.append(current_dir)
        
        # Add custom directories from preferences
        for directory in self.preferences.get("directories", []):
            # Don't duplicate the current directory
            if directory != current_dir and directory not in search_dirs:
                search_dirs.append(directory)
        
        # If no directories to search, force include current directory
        if not search_dirs:
            search_dirs.append(current_dir)
            self.preferences["exclude_current_dir"] = False
            self.save_preferences()
            self.debug_print("No search directories, including current directory")
        
        self.debug_print(f"Searching in {len(search_dirs)} directories")
        
        # Search for show directories in all search directories
        for search_dir in search_dirs:
            if os.path.exists(search_dir) and os.path.isdir(search_dir):
                self.debug_print(f"Searching for shows in: {search_dir}")
                try:
                    dir_contents = [d for d in os.listdir(search_dir) 
                                   if os.path.isdir(os.path.join(search_dir, d)) 
                                   and not d.startswith('.') 
                                   and d not in ['.git']]
                    
                    # Add search_dir prefix to each show directory
                    for show in dir_contents:
                        full_path = os.path.join(search_dir, show)
                        # Check if the directory has a Subtitles folder
                        subtitle_path = os.path.join(full_path, 'Subtitles')
                        if os.path.exists(subtitle_path) and os.path.isdir(subtitle_path):
                            shows_paths.append(full_path)
                            
                            # Use just the show name for display, but handle potential duplicates
                            show_name = os.path.basename(full_path)
                            # If there's a duplicate show name, append the parent directory
                            count = 1
                            original_name = show_name
                            while show_name in self.show_name_to_path_map:
                                parent_dir = os.path.basename(os.path.dirname(full_path))
                                show_name = f"{original_name} ({parent_dir})"
                                count += 1
                                if count > 10:  # Safety to prevent infinite loop
                                    show_name = f"{original_name} ({full_path})"
                                    break
                                    
                            # Add to the mapping
                            self.show_name_to_path_map[show_name] = full_path
                            
                            self.debug_print(f"Found show directory with subtitles: {full_path} -> {show_name}")
                except Exception as e:
                    self.debug_print(f"Error scanning directory {search_dir}: {e}")
        
        # Get sorted list of show names for the dropdown
        show_names = sorted(list(self.show_name_to_path_map.keys()))
        
        # Update dropdown with show names (not full paths)
        self.show_dropdown['values'] = show_names
        if show_names:
            self.show_dropdown.current(0)
            
        self.debug_print(f"Loaded {len(show_names)} shows")
        return shows_paths
    
    def map_subtitles_to_videos(self):
        """Map subtitle files to their corresponding video files"""
        self.status_var.set("Mapping subtitle files to videos...")
        
        # Clear previous mappings
        self.subtitle_to_video_map = {}
        
        # Common video file extensions
        video_extensions = ['.mp4', '.mkv', '.avi', '.mov', '.wmv', '.m4v', '.flv', '.webm']
        
        # Get all show paths from the name-to-path mapping
        show_paths = list(self.show_name_to_path_map.values())
        
        for show_path in show_paths:
            show_name = os.path.basename(show_path)
            # Find all subtitle files (focus on .srt files)
            subtitle_files = []
            subtitle_path = os.path.join(show_path, 'Subtitles')
            
            # Check if there's a dedicated Subtitles folder
            if os.path.exists(subtitle_path):
                subtitle_files.extend(glob.glob(os.path.join(subtitle_path, '*.srt')))
            
            # Find all video files anywhere in the show directory
            video_files = []
            
            # Walk through the entire directory structure to find all video files
            for root, dirs, files in os.walk(show_path):
                for file in files:
                    if any(file.lower().endswith(ext) for ext in video_extensions):
                        video_files.append(os.path.join(root, file))
            
            self.debug_print(f"Found {len(subtitle_files)} subtitle files and {len(video_files)} video files for {show_name}")
            
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
        selected_show_name = self.show_var.get()
        
        if not keyword:
            self.status_var.set("Please enter a search keyword.")
            return
            
        if not selected_show_name:
            self.status_var.set("Please select a show.")
            return
        
        # Get the full path for the selected show
        if selected_show_name not in self.show_name_to_path_map:
            self.status_var.set(f"Show path not found for: {selected_show_name}")
            return
            
        selected_show_path = self.show_name_to_path_map[selected_show_name]
        
        # Clear previous results
        for widget in self.results_container.winfo_children():
            widget.destroy()
        
        self.search_results = []
        
        self.debug_print(f"Searching for '{keyword}' in {selected_show_name} ({selected_show_path})")
        
        # Start search in a separate thread to keep UI responsive
        threading.Thread(target=self._search_thread, args=(keyword, selected_show_path)).start()
    
    def _search_thread(self, keyword, selected_show_path):
        """Thread function to handle the search"""
        self.status_var.set(f"Searching for '{keyword}' in {os.path.basename(selected_show_path)}...")
        
        # Get the full path for the subtitle directory
        subtitle_path = os.path.join(selected_show_path, 'Subtitles')
        if not os.path.exists(subtitle_path):
            self.status_var.set(f"Subtitle directory not found for {os.path.basename(selected_show_path)}")
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
        self.debug_print(f"Found {total_results} matches in {os.path.basename(selected_show_path)}")
        self.root.after(0, lambda: self.status_var.set(f"Found {total_results} matches in {os.path.basename(selected_show_path)}"))
    
    def _update_results_ui(self, subtitle_file, file_results):
        """Update the UI with search results (called from main thread)"""
        # Add file header
        file_basename = os.path.basename(subtitle_file)
        show_name = os.path.basename(os.path.dirname(os.path.dirname(subtitle_file)))
        file_header = ttk.Label(
            self.results_container, 
            text=f"Show: {show_name} | File: {file_basename}", 
            font=("TkDefaultFont", 10, "bold"), 
            foreground="green"
        )
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

    def _ctrl_backspace_handler(self, event):
        """Handle Ctrl+Backspace to delete the word to the left of cursor"""
        entry_widget = event.widget
        
        # Get the current cursor position
        current_pos = entry_widget.index(tk.INSERT)
        
        if current_pos == 0:
            # Nothing to delete if cursor is at the beginning
            return "break"
        
        # Get the current text
        text = entry_widget.get()
        
        # Find the start of the current word
        word_start = current_pos - 1
        
        # Skip spaces if we're in a space region
        while word_start >= 0 and text[word_start] == ' ':
            word_start -= 1
            
        # Find the beginning of the word
        while word_start >= 0 and text[word_start] != ' ':
            word_start -= 1
        
        word_start += 1  # Move past the space or to position 0
        
        # Delete directly using widget methods (more reliable than manipulating StringVar)
        entry_widget.delete(word_start, current_pos)
        
        # Ensure the widget updates immediately
        entry_widget.update_idletasks()
        
        # Prevent the default behavior
        return "break"

    def load_preferences(self):
        """Load preferences from file or create default preferences if file doesn't exist"""
        prefs_path = os.path.join(self.script_dir, PREFS_FILENAME)
        self.debug_print(f"Loading preferences from: {prefs_path}")
        
        try:
            if os.path.exists(prefs_path):
                with open(prefs_path, 'r') as f:
                    prefs = json.load(f)
                    self.debug_print(f"Loaded preferences: {prefs}")
                    return prefs
        except Exception as e:
            self.debug_print(f"Error loading preferences: {e}")
            
        # Return default preferences if file doesn't exist or has errors
        self.debug_print("Using default preferences")
        return DEFAULT_PREFS.copy()
    
    def save_preferences(self):
        """Save preferences to file"""
        prefs_path = os.path.join(self.script_dir, PREFS_FILENAME)
        self.debug_print(f"Saving preferences to: {prefs_path}")
        
        try:
            with open(prefs_path, 'w') as f:
                json.dump(self.preferences, f, indent=4)
                self.debug_print("Preferences saved successfully")
        except Exception as e:
            self.debug_print(f"Error saving preferences: {e}")
            self.status_var.set(f"Error saving preferences: {e}")
    
    def get_current_directory(self):
        """Get the current directory with a consistent format"""
        return os.path.abspath(".")
        
    def update_directory_listbox(self):
        """Update the directory listbox with current preferences"""
        self.dir_listbox.delete(0, tk.END)
        
        # Add current directory if not excluded
        current_dir = self.get_current_directory()
        if not self.preferences.get("exclude_current_dir", False):
            self.dir_listbox.insert(tk.END, current_dir + " (Current)")
        
        # Add all directories from preferences
        for directory in self.preferences.get("directories", []):
            # Don't add the current directory twice
            if directory != current_dir:
                self.dir_listbox.insert(tk.END, directory)
                
        # If the listbox is empty, force add current directory and update preferences
        if self.dir_listbox.size() == 0:
            self.dir_listbox.insert(tk.END, current_dir + " (Current)")
            self.preferences["exclude_current_dir"] = False
            self.save_preferences()
            self.debug_print("No directories left, re-added current directory")
    
    def add_directory(self):
        """Open file dialog to add directories to preferences"""
        # Add a button to select multiple directories
        root = tk.Toplevel(self.root)
        root.title("Add Media Directories")
        root.geometry("500x400")
        
        # Create frame to hold the listbox and buttons
        frame = ttk.Frame(root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Instructions label
        ttk.Label(frame, text="Select one or more directories that contain your media:").pack(pady=5, anchor="w")
        ttk.Label(frame, text="Note: You can remove the current directory if you have other directories added.").pack(pady=2, anchor="w")
        
        # Listbox to display selected directories
        listbox_frame = ttk.Frame(frame)
        listbox_frame.pack(fill="both", expand=True, pady=5)
        
        listbox = tk.Listbox(listbox_frame, selectmode="extended", height=10)
        listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # Function to add a directory
        def select_dir():
            new_dir = filedialog.askdirectory(
                title="Select Media Directory",
                initialdir=self.script_dir,
                mustexist=True
            )
            if new_dir and new_dir not in [listbox.get(i) for i in range(listbox.size())]:
                listbox.insert(tk.END, new_dir)
        
        # Function to remove selected directories from the listbox
        def remove_selected():
            selected = list(listbox.curselection())
            selected.reverse()  # Reverse to delete from bottom to top
            for i in selected:
                listbox.delete(i)
        
        # Function to save and close
        def save_and_close():
            # Get all directories from the listbox
            new_dirs = [listbox.get(i) for i in range(listbox.size())]
            
            # Update preferences
            if "directories" not in self.preferences:
                self.preferences["directories"] = []
                
            # Keep track of current directory
            current_dir = self.get_current_directory()
                
            # Add new directories if not already in preferences, excluding current directory
            # since it's always added automatically
            added_count = 0
            for new_dir in new_dirs:
                if new_dir != current_dir and new_dir not in self.preferences["directories"]:
                    self.preferences["directories"].append(new_dir)
                    added_count += 1
            
            # Save and update
            if added_count > 0:
                self.save_preferences()
                self.update_directory_listbox()
                
                # Reload shows and remap files
                self.load_shows()
                self.map_subtitles_to_videos()
                
                self.status_var.set(f"Added {added_count} new directories to preferences")
            
            # Close the dialog
            root.destroy()
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=10)
        
        # Buttons
        select_btn = ttk.Button(button_frame, text="Add Directory", command=select_dir)
        select_btn.pack(side="left", padx=5)
        
        remove_btn = ttk.Button(button_frame, text="Remove Selected", command=remove_selected)
        remove_btn.pack(side="left", padx=5)
        
        # Bottom button frame
        bottom_button_frame = ttk.Frame(frame)
        bottom_button_frame.pack(fill="x", pady=10)
        
        cancel_btn = ttk.Button(bottom_button_frame, text="Cancel", command=root.destroy)
        cancel_btn.pack(side="right", padx=5)
        
        save_btn = ttk.Button(bottom_button_frame, text="Save and Close", command=save_and_close)
        save_btn.pack(side="right", padx=5)
        
        # Make the dialog modal
        root.transient(self.root)
        root.grab_set()
        self.root.wait_window(root)
    
    def remove_directory(self):
        """Remove selected directory from preferences"""
        selected_indices = self.dir_listbox.curselection()
        
        if not selected_indices:
            self.status_var.set("No directory selected")
            return
        
        selected_index = selected_indices[0]
        selected_dir = self.dir_listbox.get(selected_index)
        current_dir = self.get_current_directory()
        
        # Check if it's the "Current" directory
        is_current = "(Current)" in selected_dir
        
        # Count total non-current directories
        custom_dirs_count = len(self.preferences.get("directories", []))
        if current_dir in self.preferences.get("directories", []):
            custom_dirs_count -= 1  # Don't count current dir if it's in the list
            
        # Only allow removing current directory if we have other directories
        if is_current:
            if custom_dirs_count == 0:
                self.status_var.set("Cannot remove current directory when it's the only one")
                return
            else:
                # Mark current directory as excluded in preferences
                self.preferences["exclude_current_dir"] = True
                self.save_preferences()
                self.update_directory_listbox()
                
                # Reload shows and remap files
                self.load_shows()
                self.map_subtitles_to_videos()
                
                self.status_var.set("Current directory excluded from search")
                return
            
        # Handle removal of non-current directory
        if selected_dir in self.preferences.get("directories", []):
            self.preferences["directories"].remove(selected_dir)
            self.save_preferences()
            
            # Reload shows and remap files
            self.load_shows()
            self.map_subtitles_to_videos()
            
            # Update the listbox
            self.update_directory_listbox()
            
            self.status_var.set(f"Removed directory: {selected_dir}")
        else:
            self.status_var.set("Directory not found in preferences")

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Rapid Moment Navigator - Search subtitles and navigate to moments in videos")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")
    args = parser.parse_args()
    
    root = tk.Tk()
    app = RapidMomentNavigator(root, debug=args.debug)
    # Force debug output to be flushed immediately if debug is enabled
    if args.debug:
        sys.stdout.flush()
    root.mainloop() 