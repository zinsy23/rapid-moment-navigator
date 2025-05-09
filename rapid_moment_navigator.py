import os
import re
import glob
import subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, Label, Frame, filedialog, messagebox
import threading
from pathlib import Path
import sys
import argparse
import json

# Constants
PREFS_FILENAME = "rapid_navigator_prefs.json"
DEFAULT_PREFS = {
    "directories": [],
    "exclude_current_dir": False,
    "selected_editor": "None"
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

class ClickableImport(Label):
    """A clickable label widget for import buttons"""
    def __init__(self, parent, text, result, callback, **kwargs):
        super().__init__(parent, text=text, cursor="hand2", fg="blue", **kwargs)
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
        
        # Create editor selection dropdown
        ttk.Label(self.search_frame, text="Editor:").pack(side="left", padx=5)
        self.editor_var = tk.StringVar()
        self.editor_dropdown = ttk.Combobox(self.search_frame, textvariable=self.editor_var, state="readonly", width=15)
        self.editor_dropdown.pack(side="left", padx=5)
        
        # Set editor dropdown options
        self.editor_dropdown['values'] = ["None", "DaVinci Resolve"]
        
        # Set default value from preferences
        selected_editor = self.preferences.get("selected_editor", "None")
        self.editor_var.set(selected_editor)
        
        # Bind editor dropdown change
        self.editor_dropdown.bind("<<ComboboxSelected>>", self._on_editor_changed)
        
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
        self.debug_print("Initializing shows and mapping...")
        self.load_shows()
        self.map_subtitles_to_videos()
        
        # Store the search results for later reference
        self.search_results = []
        
        # Debug print
        self.debug_print(f"Application initialized with {len(self.show_name_to_path_map)} shows")
    
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
    
    def update_show_dropdown(self):
        """Update the show dropdown with current show names"""
        # Get sorted list of show names for the dropdown
        show_names = sorted(list(self.show_name_to_path_map.keys()))
        
        # Update dropdown with show names (not full paths)
        self.debug_print(f"Updating dropdown with {len(show_names)} shows: {show_names}")
        
        # Directly set the 'values' property of the dropdown
        self.show_dropdown['values'] = show_names
        
        # Select first item if available
        if show_names:
            self.show_dropdown.current(0)
        
        # Force immediate update
        self.root.update_idletasks()
            
    def load_shows(self):
        """Load the available shows from the directory structure"""
        shows_paths = []
        self.show_name_to_path_map.clear()  # Clear the mapping
        
        # Prepare search directories
        current_dir = self.get_current_directory()
        search_dirs = []
        
        # Handle current directory - if included, add subdirectories as individual shows
        if not self.preferences.get("exclude_current_dir", False):
            self.debug_print(f"Load shows - including current directory: {current_dir}")
            
            # Get all immediate subdirectories in the current directory
            try:
                current_dir_subdirs = [os.path.join(current_dir, d) for d in os.listdir(current_dir) 
                                     if os.path.isdir(os.path.join(current_dir, d)) 
                                     and not d.startswith('.') 
                                     and d not in ['.git']]
                
                self.debug_print(f"Load shows - found {len(current_dir_subdirs)} subdirectories in current directory")
                
                # For each subdirectory in current directory, check if it has subtitle files
                for subdir in current_dir_subdirs:
                    # Check if this subdirectory has any SRT files (anywhere in its tree)
                    has_srt_files = False
                    
                    try:
                        for dirpath, dirnames, filenames in os.walk(subdir):
                            # Skip hidden directories
                            dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != '.git']
                            
                            # Check if any SRT files exist in this directory
                            if any(f.lower().endswith('.srt') for f in filenames):
                                has_srt_files = True
                                break
                        
                        if has_srt_files:
                            # This is a valid show directory
                            shows_paths.append(subdir)
                            
                            # Use the basename of the subdirectory as show name
                            show_name = os.path.basename(subdir)
                            
                            # Handle duplicates by appending parent directory if needed
                            count = 1
                            original_name = show_name
                            while show_name in self.show_name_to_path_map:
                                show_name = f"{original_name} ({count})"
                                count += 1
                                if count > 10:  # Safety to prevent infinite loop
                                    show_name = f"{original_name} ({subdir})"
                                    break
                            
                            # Add to the mapping
                            self.show_name_to_path_map[show_name] = subdir
                            self.debug_print(f"Load shows - added current dir show: {show_name} -> {subdir}")
                    
                    except Exception as e:
                        self.debug_print(f"Load shows - error scanning subdirectory {subdir}: {e}")
            
            except Exception as e:
                self.debug_print(f"Load shows - error listing current directory contents: {e}")
        else:
            self.debug_print(f"Load shows - current directory is excluded: {current_dir}")
        
        # Handle custom directories from preferences (each is a complete show)
        custom_dirs = list(self.preferences.get("directories", []))
        self.debug_print(f"Load shows - custom directories from preferences: {custom_dirs}")
        
        # Add custom directories (each directory is treated as a complete show)
        for directory in custom_dirs:
            # Don't duplicate the current directory
            if directory != current_dir:
                if os.path.exists(directory) and os.path.isdir(directory):
                    search_dirs.append(directory)
                    self.debug_print(f"Load shows - added directory to search: {directory}")
                else:
                    self.debug_print(f"Load shows - ignoring non-existent directory: {directory}")
        
        self.debug_print(f"Load shows - custom search directories ({len(search_dirs)}): {search_dirs}")
        
        # If no directories and no shows from current directory, force include current directory
        if not shows_paths and not search_dirs:
            self.debug_print(f"Load shows - no shows or directories found, adding current directory")
            self.preferences["exclude_current_dir"] = False
            self.save_preferences()
            
            # Re-run the method to include current directory
            return self.load_shows()
        
        # Process each custom directory as a complete show
        for directory in search_dirs:
            # Check if this directory has any SRT files
            subtitle_files = []
            
            self.debug_print(f"Load shows - recursively scanning for SRT files in: {directory}")
            
            # Walk through all subdirectories to find SRT files
            try:
                for dirpath, dirnames, filenames in os.walk(directory):
                    # Skip hidden directories
                    dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != '.git']
                    
                    # Find all SRT files in this directory
                    srt_files = [os.path.join(dirpath, f) for f in filenames if f.lower().endswith('.srt')]
                    if srt_files:
                        self.debug_print(f"Load shows - found {len(srt_files)} SRT files in: {dirpath}")
                        subtitle_files.extend(srt_files)
            
            except Exception as e:
                self.debug_print(f"Load shows - error scanning directory {directory}: {e}")
                continue
            
            # If we found subtitle files, add this as a show
            if subtitle_files:
                self.debug_print(f"Load shows - found total of {len(subtitle_files)} SRT files in {os.path.basename(directory)}")
                
                # Add this directory as a show
                shows_paths.append(directory)
                
                # Use the root directory name for display in the dropdown
                show_name = os.path.basename(directory)
                
                # Handle duplicates by appending parent directory if needed
                count = 1
                original_name = show_name
                while show_name in self.show_name_to_path_map:
                    parent_dir = os.path.basename(os.path.dirname(directory))
                    show_name = f"{original_name} ({parent_dir})"
                    count += 1
                    if count > 10:  # Safety to prevent infinite loop
                        show_name = f"{original_name} ({directory})"
                        break
                
                # Add to the mapping
                self.show_name_to_path_map[show_name] = directory
                
                self.debug_print(f"Load shows - added custom dir show: {show_name} -> {directory}")
            else:
                self.debug_print(f"Load shows - no SRT files found in {directory}")
        
        # Update the dropdown with show names
        self.debug_print(f"Load shows - completed. Found {len(self.show_name_to_path_map)} shows from all sources.")
        self.update_show_dropdown()
        
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
            
            # Find all subtitle files (anywhere in the directory tree)
            subtitle_files = []
            # Find all video files (anywhere in the directory tree)
            video_files = []
            
            self.debug_print(f"Mapping - scanning for subtitle and video files in: {show_path}")
            
            # Walk through all subdirectories
            try:
                for dirpath, dirnames, filenames in os.walk(show_path):
                    # Skip hidden directories
                    dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != '.git']
                    
                    # Find all SRT files in this directory
                    srt_files = [os.path.join(dirpath, f) for f in filenames if f.lower().endswith('.srt')]
                    if srt_files:
                        subtitle_files.extend(srt_files)
                        
                    # Find all video files in this directory
                    video_files_here = [os.path.join(dirpath, f) for f in filenames 
                                       if any(f.lower().endswith(ext) for ext in video_extensions)]
                    if video_files_here:
                        video_files.extend(video_files_here)
            
            except Exception as e:
                self.debug_print(f"Mapping - error scanning directory {show_path}: {e}")
                continue
            
            self.debug_print(f"Mapping - found {len(subtitle_files)} subtitle files and {len(video_files)} video files in {show_name}")
            
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
                        self.debug_print(f"Mapping - exact match: {subtitle_basename} -> {video_basename}")
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
                            self.debug_print(f"Mapping - partial match: {subtitle_basename} -> {video_basename}")
                            matched = True
                            break
        
        self.debug_print(f"Mapping - completed. Mapped {len(self.subtitle_to_video_map)} subtitle files to videos")
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
        show_name = os.path.basename(selected_show_path)
        self.status_var.set(f"Searching for '{keyword}' in {show_name}...")
        
        # Find all SRT files in the selected show directory
        subtitle_files = []
        
        self.debug_print(f"Search - recursively scanning for SRT files in: {selected_show_path}")
        
        # Walk through all subdirectories to find SRT files
        try:
            for dirpath, dirnames, filenames in os.walk(selected_show_path):
                # Skip hidden directories
                dirnames[:] = [d for d in dirnames if not d.startswith('.') and d != '.git']
                
                # Find all SRT files in this directory
                srt_files = [os.path.join(dirpath, f) for f in filenames 
                            if f.lower().endswith('.srt') or f.lower().endswith('.txt')]
                
                if srt_files:
                    self.debug_print(f"Search - found {len(srt_files)} subtitle files in: {dirpath}")
                    subtitle_files.extend(srt_files)
        
        except Exception as e:
            self.debug_print(f"Search - error scanning directory {selected_show_path}: {e}")
            self.status_var.set(f"Error scanning directory: {e}")
            return
        
        if not subtitle_files:
            self.status_var.set(f"No subtitle files found in {show_name}")
            return
            
        self.debug_print(f"Search - found total of {len(subtitle_files)} subtitle files to search")
        
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
        self.debug_print(f"Found {total_results} matches in {show_name}")
        self.root.after(0, lambda: self.status_var.set(f"Found {total_results} matches in {show_name}"))
    
    def _update_results_ui(self, subtitle_file, file_results):
        """Update the UI with search results (called from main thread)"""
        # Add file header
        file_basename = os.path.basename(subtitle_file)
        
        # Find the show name this subtitle belongs to
        show_path = None
        for show_name, path in self.show_name_to_path_map.items():
            if subtitle_file.startswith(path):
                show_path = path
                break
        
        # Get relative path from show root if possible
        if show_path:
            relative_path = os.path.relpath(subtitle_file, show_path)
            header_text = f"File: {relative_path}"
        else:
            header_text = f"File: {file_basename}"
        
        # Create a frame for the file header
        header_frame = ttk.Frame(self.results_container)
        header_frame.pack(fill="x", padx=5, pady=2)
        
        # Add file header label
        file_header = ttk.Label(
            header_frame, 
            text=header_text, 
            font=("TkDefaultFont", 10, "bold"), 
            foreground="green"
        )
        file_header.pack(side="left", anchor="w")
        
        # Check if we should show import buttons
        selected_editor = self.editor_var.get()
        show_import_buttons = selected_editor != "None"
        
        # Add each result
        for result in file_results:
            # Create a frame for this result
            result_frame = ttk.Frame(self.results_container)
            result_frame.pack(fill="x", padx=5, pady=2, anchor="w")
            
            # Create import buttons frame at the top right
            import_buttons_frame = ttk.Frame(result_frame)
            # Store reference to the import buttons frame for later visibility updates
            result_frame.import_buttons_frame = import_buttons_frame
            
            if show_import_buttons:
                import_buttons_frame.pack(side="right", padx=5, anchor="ne")
                
                # Add Import Media button
                import_media_btn = ClickableImport(
                    import_buttons_frame, 
                    "Import Media", 
                    result, 
                    self._handle_import_media_click
                )
                import_media_btn.pack(side="left", padx=5)
                
                # Add Import Clip button
                import_clip_btn = ClickableImport(
                    import_buttons_frame, 
                    "Import Clip", 
                    result, 
                    self._handle_import_clip_click
                )
                import_clip_btn.pack(side="left", padx=5)
            
            # Create content frame (with timecode and text) that fills the remaining space
            content_frame = ttk.Frame(result_frame)
            content_frame.pack(side="left", fill="both", expand=True, anchor="w")
            
            # Create clickable timecode label
            timecode_text = f"{result['start_time']} --> {result['end_time']}"
            timecode_label = ClickableTimecode(
                content_frame, 
                timecode_text, 
                result, 
                self._handle_timecode_click
            )
            timecode_label.pack(anchor="w")
            
            # Add text label
            text_label = ttk.Label(content_frame, text=result['clean_text'], wraplength=700)
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

    def _on_editor_changed(self, event):
        """Handle editor dropdown change"""
        selected_editor = self.editor_var.get()
        self.debug_print(f"Editor changed to: {selected_editor}")
        
        # Update preferences
        self.preferences["selected_editor"] = selected_editor
        self.save_preferences()
        
        # Update UI to show/hide import buttons
        self._update_import_buttons_visibility()
    
    def _update_import_buttons_visibility(self):
        """Update visibility of import buttons based on selected editor"""
        selected_editor = self.editor_var.get()
        show_import_buttons = selected_editor != "None"
        
        # Loop through all frames and update import buttons visibility
        for widget in self.results_container.winfo_children():
            if isinstance(widget, ttk.Frame) and hasattr(widget, "import_buttons_frame"):
                if show_import_buttons:
                    widget.import_buttons_frame.pack(side="right", padx=5)
                else:
                    widget.import_buttons_frame.pack_forget()
        
        # Update the canvas scroll region
        self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
    
    def _handle_import_media_click(self, result):
        """Handle click on Import Media button"""
        selected_editor = self.editor_var.get()
        subtitle_file = result['file']
        
        if subtitle_file in self.subtitle_to_video_map:
            video_file = self.subtitle_to_video_map[subtitle_file]
            self.debug_print(f"Import Media clicked for {os.path.basename(video_file)} with editor {selected_editor}")
            
            # Call the appropriate import function based on selected editor
            if selected_editor == "DaVinci Resolve":
                self._import_media_to_davinci_resolve(video_file)
            # Add more editors as needed
            
            self.status_var.set(f"Importing {os.path.basename(video_file)} to {selected_editor}")
        else:
            self.debug_print(f"No matching video file found for {os.path.basename(subtitle_file)}")
            self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def _handle_import_clip_click(self, result):
        """Handle click on Import Clip button"""
        selected_editor = self.editor_var.get()
        subtitle_file = result['file']
        start_time = result['mpc_start_time']
        end_time = result['end_time'].replace(',', '.').split('.')[0]  # Convert to MPC format
        
        if subtitle_file in self.subtitle_to_video_map:
            video_file = self.subtitle_to_video_map[subtitle_file]
            self.debug_print(f"Import Clip clicked for {os.path.basename(video_file)} at {start_time}-{end_time} with editor {selected_editor}")
            
            # Call the appropriate import function based on selected editor
            if selected_editor == "DaVinci Resolve":
                self._import_clip_to_davinci_resolve(video_file, start_time, end_time)
            # Add more editors as needed
            
            self.status_var.set(f"Importing clip from {os.path.basename(video_file)} at {start_time}-{end_time} to {selected_editor}")
        else:
            self.debug_print(f"No matching video file found for {os.path.basename(subtitle_file)}")
            self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def _import_media_to_davinci_resolve(self, video_file):
        """Import full media file to DaVinci Resolve"""
        # This is a placeholder for future implementation
        self.debug_print(f"TODO: Implement importing {video_file} to DaVinci Resolve")
        # Future implementation will go here
    
    def _import_clip_to_davinci_resolve(self, video_file, start_time, end_time):
        """Import clip with time range to DaVinci Resolve"""
        # This is a placeholder for future implementation
        self.debug_print(f"TODO: Implement importing {video_file} from {start_time} to {end_time} to DaVinci Resolve")
        # Future implementation will go here
    
    def load_preferences(self):
        """Load preferences from file or create default preferences if file doesn't exist"""
        prefs_path = os.path.join(self.script_dir, PREFS_FILENAME)
        self.debug_print(f"Loading preferences from: {prefs_path}")
        
        try:
            if os.path.exists(prefs_path):
                with open(prefs_path, 'r') as f:
                    prefs = json.load(f)
                    self.debug_print(f"Loaded preferences: {prefs}")
                    
                    # Ensure all expected keys are present
                    for key in DEFAULT_PREFS.keys():
                        if key not in prefs:
                            prefs[key] = DEFAULT_PREFS[key]
                    
                    # Validate directory paths
                    valid_dirs = []
                    for dir_path in prefs.get("directories", []):
                        if os.path.exists(dir_path) and os.path.isdir(dir_path):
                            valid_dirs.append(dir_path)
                        else:
                            self.debug_print(f"Ignoring non-existent directory from preferences: {dir_path}")
                    
                    prefs["directories"] = valid_dirs
                    
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
        
        # Add existing directories to the listbox
        existing_dirs = self.preferences.get("directories", [])
        for dir_path in existing_dirs:
            listbox.insert(tk.END, dir_path)
        
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
            self.debug_print(f"Directory dialog - new directories: {new_dirs}")
            
            # Keep track of current directory
            current_dir = self.get_current_directory()
            self.debug_print(f"Directory dialog - current directory: {current_dir}")
            
            # Get existing directories from preferences
            existing_dirs = self.preferences.get("directories", [])
            self.debug_print(f"Directory dialog - existing directories: {existing_dirs}")
            
            # Track if we've made any changes
            changes_made = False
            
            # Add new directories if not already in preferences
            added_count = 0
            for new_dir in new_dirs:
                # Skip current directory since it's handled separately
                if new_dir == current_dir:
                    continue
                    
                # Check if this is a new directory to add
                if new_dir not in existing_dirs:
                    self.debug_print(f"Directory dialog - adding new directory: {new_dir}")
                    if "directories" not in self.preferences:
                        self.preferences["directories"] = []
                    self.preferences["directories"].append(new_dir)
                    added_count += 1
                    changes_made = True
            
            # Save and update if changes were made
            if changes_made:
                self.save_preferences()
                self.debug_print(f"Directory dialog - saved preferences: {self.preferences}")
                self.update_directory_listbox()
                
                # Clear existing show map and reload everything
                self.debug_print("Directory dialog - clearing show map and reloading shows")
                self.show_name_to_path_map.clear()
                
                # Reload shows and remap files
                self.debug_print("Directory dialog - reloading shows after directory changes")
                shows_paths = self.load_shows()
                self.debug_print(f"Directory dialog - loaded shows: {len(shows_paths)}, names: {list(self.show_name_to_path_map.keys())}")
                self.map_subtitles_to_videos()
                
                # Force the dropdown to update with the new values
                self.debug_print(f"Directory dialog - updating dropdown with {len(self.show_name_to_path_map)} shows")
                self.update_show_dropdown()
                
                self.status_var.set(f"Added {added_count} directories. Found {len(self.show_name_to_path_map)} shows")
            else:
                self.status_var.set("No changes made to media directories")
            
            # Close the dialog
            root.destroy()
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=10)
        
        # Buttons
        select_btn = ttk.Button(button_frame, text="Add Directory", command=select_dir)
        select_btn.pack(side="left", padx=5)
        
        # Function to add multiple directories at once using a directory selection dialog
        def select_multiple_dirs():
            # Show info message about selecting multiple directories
            self.debug_print("Opening directory selection dialog for multiple directories")
            root.withdraw()  # Hide the current dialog temporarily
            
            # Use a message box to instruct the user
            tk.messagebox.showinfo(
                "Multiple Directory Selection",
                "Please select directories one by one. Click Cancel when done."
            )
            
            # Keep track of added directories
            added_dirs = []
            
            # Loop to select multiple directories
            while True:
                new_dir = filedialog.askdirectory(
                    title="Select Media Directory (Cancel when done)",
                    initialdir=self.script_dir if not added_dirs else added_dirs[-1],  # Start from last selected directory
                    mustexist=True
                )
                
                if not new_dir:  # User clicked Cancel
                    break
                    
                added_dirs.append(new_dir)
                
                # Add to listbox if not already there
                if new_dir not in [listbox.get(i) for i in range(listbox.size())]:
                    listbox.insert(tk.END, new_dir)
            
            # Show the dialog again
            root.deiconify()
            
            # Report how many directories were added
            if added_dirs:
                self.debug_print(f"Added {len(added_dirs)} directories through multiple selection")
        
        # Add button for multiple directory selection
        multi_select_btn = ttk.Button(button_frame, text="Add Multiple Directories", command=select_multiple_dirs)
        multi_select_btn.pack(side="left", padx=5)
        
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
        
        self.debug_print(f"Remove directory - selected: {selected_dir}, current: {current_dir}")
        
        # Check if it's the "Current" directory
        is_current = "(Current)" in selected_dir
        
        # Count total non-current directories
        custom_dirs_count = len(self.preferences.get("directories", []))
        if current_dir in self.preferences.get("directories", []):
            custom_dirs_count -= 1  # Don't count current dir if it's in the list
            
        self.debug_print(f"Remove directory - is current: {is_current}, custom dirs count: {custom_dirs_count}")
            
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
                
                # Clear existing show map and reload everything
                self.show_name_to_path_map.clear()
                
                # Reload shows and remap files
                shows_paths = self.load_shows()
                self.map_subtitles_to_videos()
                
                self.status_var.set(f"Current directory excluded. Found {len(self.show_name_to_path_map)} shows")
                return
            
        # Handle removal of non-current directory
        if selected_dir in self.preferences.get("directories", []):
            self.preferences["directories"].remove(selected_dir)
            self.save_preferences()
            
            # Clear existing show map and reload everything
            self.show_name_to_path_map.clear()
            
            # Update the listbox
            self.update_directory_listbox()
            
            # Reload shows and remap files
            shows_paths = self.load_shows()
            self.map_subtitles_to_videos()
            
            self.status_var.set(f"Removed directory: {selected_dir}. Found {len(self.show_name_to_path_map)} shows")
        else:
            self.status_var.set("Directory not found in preferences")

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Rapid Moment Navigator - Search subtitles and navigate to moments in videos")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")
    args = parser.parse_args()
    
    if args.debug:
        print(f"DEBUG: Starting application in debug mode")
        print(f"DEBUG: Current directory: {os.path.abspath('.')}")
        print(f"DEBUG: Script directory: {os.path.dirname(os.path.abspath(__file__))}")
    
    root = tk.Tk()
    app = RapidMomentNavigator(root, debug=args.debug)
    # Force debug output to be flushed immediately if debug is enabled
    if args.debug:
        sys.stdout.flush()
    root.mainloop() 