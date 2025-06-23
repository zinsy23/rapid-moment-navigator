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
import logging
import time
import traceback
import multiprocessing
import tempfile

# Constants
PREFS_FILENAME = "rapid_navigator_prefs.json"
RESOLVE_PATHS_FILENAME = "resolve_paths.json"
DEFAULT_PREFS = {
    "directories": [],
    "exclude_current_dir": False,
    "selected_editor": "None",
    "min_duration_enabled": True,  # Changed from False to True
    "min_duration_seconds": 10.0
}

def find_module_locations(base_path):
    """Find possible locations of DaVinciResolveScript.py based on a base path.
    Only checks the standard location and directly in the specified path."""
    locations = []
    module_paths = []
    
    # Check standard location (base_path/Modules/DaVinciResolveScript.py)
    standard_location = os.path.join(base_path, "Modules", "DaVinciResolveScript.py")
    if os.path.isfile(standard_location):
        locations.append(os.path.dirname(standard_location))
        module_paths.append(standard_location)
        
    # Check directly in base path (base_path/DaVinciResolveScript.py)
    direct_location = os.path.join(base_path, "DaVinciResolveScript.py")
    if os.path.isfile(direct_location):
        locations.append(base_path)
        module_paths.append(direct_location)
    
    return {
        "locations": locations,  # Directories containing the module
        "module_paths": module_paths  # Full paths to the module files
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
    def __init__(self, parent, text, result, callback, tooltip=None, **kwargs):
        super().__init__(parent, text=text, cursor="hand2", fg="blue", **kwargs)
        self.result = result
        self.callback = callback
        self.bind("<Button-1>", self._on_click)
        # Add underline
        self.config(font=("TkDefaultFont", 10, "underline"))
        
        # Add tooltip functionality
        self.tooltip_text = tooltip
        if tooltip:
            self.bind("<Enter>", self._on_enter)
            self.bind("<Leave>", self._on_leave)
            self.tooltip = None
            self.tooltip_timer = None
        
    def _on_click(self, event):
        """Handle click event"""
        self.callback(self.result)
        
    def _on_enter(self, event):
        """Start timer to show tooltip when mouse enters the widget"""
        if self.tooltip_text:
            # Cancel any existing timer
            self._cancel_timer()
            # Start a new timer - 1000ms = 1 second
            self.tooltip_timer = self.after(1000, self._show_tooltip)
    
    def _show_tooltip(self):
        """Display the tooltip after delay"""
        x, y, _, _ = self.bbox("insert")
        x += self.winfo_rootx() + 25
        y += self.winfo_rooty() + 25
        
        # Create a toplevel window
        self.tooltip = tk.Toplevel(self)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        # Create tooltip content
        frame = tk.Frame(self.tooltip, background="#ffffe0", borderwidth=1, relief="solid")
        frame.pack(ipadx=3, ipady=2)
        
        label = tk.Label(frame, text=self.tooltip_text, justify="left",
                      background="#ffffe0", fg="#000000", 
                      wraplength=250, font=("TkDefaultFont", 9))
        label.pack()
    
    def _on_leave(self, event):
        """Hide tooltip and cancel timer when mouse leaves the widget"""
        self._cancel_timer()
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
            
    def _cancel_timer(self):
        """Cancel any pending tooltip timer"""
        if self.tooltip_timer:
            self.after_cancel(self.tooltip_timer)
            self.tooltip_timer = None

class RapidMomentNavigator:
    def __init__(self, root, debug=False):
        self.root = root
        self.root.title("Rapid Moment Navigator")
        
        # Calculate the centered position for the main window
        window_width = 800
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # Set both size and position in one geometry call
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Debug mode setting
        self.debug = debug
        
        # Initialize debug window to None
        self.debug_window = None
        
        # Setup exception handling for Tkinter
        self.setup_exception_handler()
        
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
        shows_paths = self.load_shows()
        self.map_subtitles_to_videos()
        
        # Store the search results for later reference
        self.search_results = []
        
        # Update status based on whether shows were found
        if len(self.show_name_to_path_map) > 0:
            self.debug_print(f"Application initialized with {len(self.show_name_to_path_map)} shows")
            self.status_var.set(f"Ready. Found {len(self.show_name_to_path_map)} shows with {len(self.subtitle_to_video_map)} mapped videos.")
        else:
            self.debug_print("No shows found during initialization")
            self.status_var.set("No media found. Please add directories containing subtitle files and videos.")
        
        # Initialize safe mode flag for editors
        self.resolve_in_safe_mode = False
        
        # Track if guidance dialog is currently showing
        self.guidance_dialog_showing = False
        
        # Store the need to show guidance dialog
        self.needs_guidance_dialog = len(self.show_name_to_path_map) == 0
        
        # Wait for window to stabilize before showing guidance dialog
        # This helps prevent the "jump" effect where the window moves after initial rendering
        if self.needs_guidance_dialog:
            self.root.after(300, self._delayed_show_guidance)
    
    def position_window(self, window, x=None, y=None, parent=None, offset_x=0, offset_y=0):
        """
        Position a window at specific coordinates or centered, with optional offsets
        
        Args:
            window: The window to position
            x, y: Specific coordinates (if None, will center)
            parent: Parent window to center relative to (if provided)
            offset_x, offset_y: Additional offsets to apply
            
        Returns:
            The positioned window
        """
        # Force the window to update and process all pending events
        window.update()
        
        # Ensure window size is updated before calculating positions
        window.update_idletasks()
        
        # Get window dimensions
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        
        # If window size is still 1x1, it's not properly initialized yet
        # This can happen with newly created windows
        if window_width <= 1 or window_height <= 1:
            self.debug_print(f"Window size not initialized yet: {window_width}x{window_height}, forcing geometry")
            # Try to force the window to its requested size
            if hasattr(window, '_w') and window._w == '.':  # Main window
                # For main window, use the initial geometry we set
                window.geometry("800x600")
                window_width = 800
                window_height = 600
            window.update_idletasks()
            window_width = window.winfo_width()
            window_height = window.winfo_height()
            self.debug_print(f"After forcing geometry: {window_width}x{window_height}")
        
        # Store original values for debugging
        orig_x, orig_y = x, y
        
        # Calculate X position independently
        if x is None:
            if parent is None:
                # Center on screen horizontally
                screen_width = window.winfo_screenwidth()
                # x = (screen_width // 2) - (window_width // 2)
                x = (screen_width - window_width) // 2
                self.debug_print(f"Calculated X center: {x} = ({screen_width} - {window_width}) // 2")
            else:
                # Center relative to parent horizontally
                parent_x = parent.winfo_x()
                parent_width = parent.winfo_width()
                x = parent_x + (parent_width - window_width) // 2
                self.debug_print(f"Calculated X center relative to parent: {x}")
        
        # Calculate Y position independently
        if y is None:
            if parent is None:
                # Center on screen vertically
                screen_height = window.winfo_screenheight()
                # y = (screen_height // 2) - (window_height // 2)
                y = (screen_height - window_height) // 2
                self.debug_print(f"Calculated Y center: {y} = ({screen_height} - {window_height}) // 2")
            else:
                # Center relative to parent vertically
                parent_y = parent.winfo_y()
                parent_height = parent.winfo_height()
                y = parent_y + (parent_height - window_height) // 2
                self.debug_print(f"Calculated Y center relative to parent: {y}")
        
        # Apply offsets
        x += offset_x
        y += offset_y
        
        # Ensure coordinates are not negative
        x = max(0, x)
        y = max(0, y)
        
        # Log screen dimensions if we calculated either position
        if orig_x is None or orig_y is None:
            if parent is None:
                screen_width = window.winfo_screenwidth()
                screen_height = window.winfo_screenheight()
                self.debug_print(f"Screen dimensions: {screen_width}x{screen_height}")
        
        # Set window position
        window.geometry(f"+{x}+{y}")
        self.debug_print(f"Positioned window at ({x},{y}) with size {window_width}x{window_height}")
        
        # Force window to update again to ensure position takes effect
        window.update_idletasks()
        
        # Verify final position
        actual_x = window.winfo_x()
        actual_y = window.winfo_y()
        if actual_x != x or actual_y != y:
            self.debug_print(f"Warning: Window position changed by window manager: ({x},{y}) → ({actual_x},{actual_y})")
        
        return window
    
    def center_window(self, window, parent=None):
        """Center a window on screen or relative to parent window"""
        # Use the positioning method with centered coordinates
        return self.position_window(window, x=None, y=None, parent=parent)
    
    def setup_exception_handler(self):
        """Setup global exception handler to catch and display errors"""
        # Store the original exception handler
        self.original_exception_handler = sys.excepthook
        
        # Create a new exception handler
        def exception_handler(exc_type, exc_value, exc_traceback):
            # Format the exception and traceback
            error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            
            # Log the error
            self.debug_print(f"UNCAUGHT EXCEPTION: {error_msg}")
            
            # Show in GUI
            self.show_error_in_gui("Application Error", 
                                  f"An unexpected error occurred:\n\n{str(exc_value)}\n\nSee debug window for details.")
            
            # Show in debug window
            self.ensure_debug_window()
            if self.debug_window:
                self.debug_window.insert_text(error_msg)
            
            # Call the original handler
            self.original_exception_handler(exc_type, exc_value, exc_traceback)
        
        # Set our custom exception handler
        sys.excepthook = exception_handler
    
    def ensure_debug_window(self):
        """Create debug window if it doesn't exist"""
        try:
            if not hasattr(self, 'debug_window') or self.debug_window is None or not self.debug_window.winfo_exists():
                self.debug_window = DebugWindow(self.root, self.debug)
                self.debug_print("Debug window created")
        except Exception as e:
            print(f"Error creating debug window: {e}", file=sys.stderr)
    
    def show_error_in_gui(self, title, message):
        """Display an error message in the GUI"""
        # Update status bar
        first_line = message.split('\n')[0]
        self.status_var.set(f"Error: {first_line}")
        
        # Show error dialog (in a separate thread to avoid blocking)
        threading.Thread(target=lambda: messagebox.showerror(title, message)).start()

    def debug_print(self, message):
        """Print debug messages and add to debug window if available"""
        if self.debug:
            print(f"DEBUG: {message}", flush=True)
            
            # Only try to use the debug window if it's properly initialized
            try:
                if hasattr(self, 'debug_window') and self.debug_window is not None and self.debug_window.winfo_exists():
                    self.debug_window.insert_text(f"DEBUG: {message}\n")
            except Exception:
                # Silently ignore any errors with the debug window
                pass
    
    def _configure_scroll_region(self, event):
        """Configure the scroll region of the canvas"""
        self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
    
    def _configure_canvas_width(self, event):
        """Make the canvas width match its container"""
        self.results_canvas.itemconfig(self.results_container_id, width=event.width)
    
    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling"""
        self.results_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
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

    def filter_nested_directories(self, directories):
        """
        Filter out nested directories if their parent directory is already in the list.
        
        Args:
            directories (list): List of directory paths
            
        Returns:
            list: Filtered list with no nested directories
        """
        if not directories:
            return []
            
        # Normalize all paths first
        normalized_dirs = [os.path.normpath(d) for d in directories]
        
        # Sort directories by path length (descending) to process most nested first
        # This ensures we compare subdirectories against parent directories
        sorted_dirs = sorted(normalized_dirs, key=len, reverse=True)
        
        # List to hold directories that aren't nested within others
        filtered_dirs = []
        
        for dir_path in sorted_dirs:
            # Check if this directory is contained within any directory we've already approved
            is_nested = False
            for approved_dir in filtered_dirs:
                # Check if dir_path is a subdirectory of approved_dir
                # dir_path starts with approved_dir followed by os separator
                if dir_path.startswith(approved_dir + os.sep):
                    is_nested = True
                    break
                    
            # Add to filtered list if not nested
            if not is_nested:
                filtered_dirs.append(dir_path)
        
        return filtered_dirs
            
    def load_shows(self):
        """Load the available shows from the directory structure"""
        shows_paths = []
        self.show_name_to_path_map.clear()  # Clear the mapping
        
        # Get manually added directories from preferences
        manual_dirs = list(self.preferences.get("directories", []))
        self.debug_print(f"Load shows - manual directories from preferences: {manual_dirs}")
        
        # Filter out nested directories from manually added directories
        filtered_manual_dirs = self.filter_nested_directories(manual_dirs)  # Use self. to call the class method
        if len(filtered_manual_dirs) != len(manual_dirs):
            self.debug_print(f"Load shows - filtered {len(manual_dirs) - len(filtered_manual_dirs)} nested manually added directories")
            self.debug_print(f"Load shows - using filtered manual directories: {filtered_manual_dirs}")
        
        # Create a set of all manual directories and their subdirectories to exclude from current dir scanning
        manual_dirs_and_subdirs = set()
        for dir_path in filtered_manual_dirs:
            if os.path.exists(dir_path) and os.path.isdir(dir_path):
                # Add the directory itself
                manual_dirs_and_subdirs.add(os.path.normpath(dir_path))
                
                # Add all subdirectories
                try:
                    for dirpath, dirnames, _ in os.walk(dir_path):
                        for dirname in dirnames:
                            subdir = os.path.normpath(os.path.join(dirpath, dirname))
                            manual_dirs_and_subdirs.add(subdir)
                except Exception as e:
                    self.debug_print(f"Load shows - error scanning subdirectories of {dir_path}: {e}")
        
        self.debug_print(f"Load shows - found {len(manual_dirs_and_subdirs)} directories to exclude from current dir scanning")
        
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
                    # Skip if this directory is in the manual_dirs_and_subdirs set
                    normalized_subdir = os.path.normpath(subdir)
                    if normalized_subdir in manual_dirs_and_subdirs:
                        self.debug_print(f"Load shows - skipping {subdir} as it's already in manual directories")
                        continue
                    
                    # Check if this is a parent directory of any manual directory
                    is_parent_of_manual = False
                    for manual_dir in filtered_manual_dirs:
                        normalized_manual = os.path.normpath(manual_dir)
                        if normalized_manual.startswith(normalized_subdir + os.sep):
                            is_parent_of_manual = True
                            self.debug_print(f"Load shows - skipping {subdir} as it's a parent of manual directory {manual_dir}")
                            break
                    
                    if is_parent_of_manual:
                        continue
                    
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
        
        # Add filtered custom directories (each directory is treated as a complete show)
        for directory in filtered_manual_dirs:
            # Don't duplicate the current directory
            if directory != current_dir:
                if os.path.exists(directory) and os.path.isdir(directory):
                    search_dirs.append(directory)
                    self.debug_print(f"Load shows - added directory to search: {directory}")
                else:
                    self.debug_print(f"Load shows - ignoring non-existent directory: {directory}")
        
        self.debug_print(f"Load shows - custom search directories ({len(search_dirs)}): {search_dirs}")
        
        # If no directories and no shows from current directory, mark for showing guidance dialog
        if not shows_paths and not search_dirs:
            self.debug_print(f"Load shows - no shows or directories found, showing guidance")
            # Set flag to show guidance dialog after main window is positioned
            self.needs_guidance_dialog = True
            return []
        
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
    
    def _show_no_shows_guidance(self):
        """Show guidance dialog when no shows are found"""
        # This method is deprecated and replaced by _delayed_show_guidance
        # We'll call the new method instead
        self._delayed_show_guidance()
    
    def map_subtitles_to_videos(self):
        """Map subtitle files to their corresponding video files"""
        self.status_var.set("Mapping subtitle files to videos...")
        
        # Clear previous mappings
        self.subtitle_to_video_map = {}
        
        # Common video file extensions
        video_extensions = ['.mp4', '.mkv', '.avi', '.mov', '.wmv', '.m4v', '.flv', '.webm', '.mp3', '.m4a', '.wav', '.aac', '.ogg']
        
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
                        # Store only the path - framerate will be detected when needed
                        self.subtitle_to_video_map[subtitle_file] = {
                            "path": video_file,
                            "fps": None  # Initialize as None, will detect when needed
                        }
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
                            
                            # Store only the path - framerate will be detected when needed
                            self.subtitle_to_video_map[subtitle_file] = {
                                "path": video_file,
                                "fps": None  # Initialize as None, will detect when needed
                            }
                            self.debug_print(f"Mapping - partial match: {subtitle_basename} -> {video_basename}")
                            matched = True
                            break
        
        self.debug_print(f"Mapping - completed. Mapped {len(self.subtitle_to_video_map)} subtitle files to videos")
        self.status_var.set(f"Ready. Mapped {len(self.subtitle_to_video_map)} subtitle files to videos.")
    
    def detect_video_framerate(self, video_path):
        """
        Detect the framerate of a video file using FFprobe.
        
        Args:
            video_path (str): Path to the video file
            
        Returns:
            float: Framerate of the video (defaults to 24.0 if detection fails)
        """
        try:
            # Get absolute path to the video file
            abs_video_path = self.get_absolute_path(video_path)
            
            # Check if ffprobe is available
            try:
                # First try ffprobe directly
                cmd = ["ffprobe", "-version"]
                subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                ffprobe_cmd = "ffprobe"
            except (subprocess.SubprocessError, FileNotFoundError):
                # If not available directly, try looking for it in common locations
                ffprobe_locations = [
                    r"C:\Program Files\FFmpeg\bin\ffprobe.exe",
                    r"C:\Program Files (x86)\FFmpeg\bin\ffprobe.exe",
                    os.path.join(self.script_dir, "ffprobe.exe")
                ]
                
                ffprobe_found = False
                for location in ffprobe_locations:
                    if os.path.exists(location):
                        ffprobe_cmd = location
                        ffprobe_found = True
                        break
                    
                if not ffprobe_found:
                    self.debug_print("FFprobe not found, using default framerate")
                    return 24.0
            
            # Construct the FFprobe command
            cmd = [
                ffprobe_cmd,
                "-v", "error",
                "-select_streams", "v:0",
                "-show_entries", "stream=r_frame_rate",
                "-of", "default=noprint_wrappers=1:nokey=1",
                abs_video_path
            ]
            
            # Run the command
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            if result.returncode != 0:
                self.debug_print(f"FFprobe error: {result.stderr}")
                return 24.0
                
            # Parse the output (usually in the form "num/den")
            fps_str = result.stdout.strip()
            
            if '/' in fps_str:
                num, den = map(int, fps_str.split('/'))
                fps = num / den if den != 0 else 24.0
            else:
                try:
                    fps = float(fps_str)
                except ValueError:
                    fps = 24.0
                    
            # Validate the framerate (ensure it's reasonable)
            if fps < 1 or fps > 300:
                self.debug_print(f"Invalid framerate detected ({fps}), using default")
                return 24.0
                
            self.debug_print(f"Detected framerate for {os.path.basename(video_path)}: {fps} fps")
            return float(fps)
        
        except Exception as e:
            self.debug_print(f"Error detecting framerate: {str(e)}")
            return 24.0  # Default fallback framerate
    
    def detect_video_framerate_from_resolve(self, video_path):
        """
        Detect the framerate of a video file using the DaVinci Resolve API.
        
        Args:
            video_path (str): Path to the video file
            
        Returns:
            float: Framerate of the video (defaults to 24.0 if detection fails)
        """
        try:
            # Check if Resolve API is initialized
            if not hasattr(self, 'resolve_initialized') or not self.resolve_initialized:
                self.debug_print("Resolve API not initialized, using fallback method")
                return self.detect_video_framerate(video_path)
            
            # Get absolute path to the video file
            abs_video_path = self.get_absolute_path(video_path)
            
            # Get Resolve
            resolve = dvr_script.scriptapp("Resolve")
            if not resolve:
                self.debug_print("Failed to get Resolve object")
                return self.detect_video_framerate(video_path)
            
            # Get project manager
            project_manager = resolve.GetProjectManager()
            if not project_manager:
                self.debug_print("Failed to get project manager")
                return self.detect_video_framerate(video_path)
            
            # Get current project
            project = project_manager.GetCurrentProject()
            if not project:
                self.debug_print("No project is currently open")
                return self.detect_video_framerate(video_path)
            
            # Get media pool
            media_pool = project.GetMediaPool()
            if not media_pool:
                self.debug_print("Failed to get media pool")
                return self.detect_video_framerate(video_path)
                
            # Import the media file temporarily to get its properties
            current_folder = media_pool.GetCurrentFolder()
            temp_folder = None
            
            try:
                # Create a temporary folder in the media pool to avoid cluttering the main pool
                root_folder = media_pool.GetRootFolder()
                folder_list = root_folder.GetSubFolderList()
                
                # Check if we already have a temp folder
                temp_folder_name = "_RapidNavigator_Temp"
                temp_folder = None
                
                for folder in folder_list:
                    if folder.GetName() == temp_folder_name:
                        temp_folder = folder
                        break
                
                # Create temp folder if it doesn't exist
                if not temp_folder:
                    temp_folder = media_pool.AddSubFolder(root_folder, temp_folder_name)
                
                # Set the temp folder as the current folder
                media_pool.SetCurrentFolder(temp_folder)
                
                # Import the media
                imported_media = media_pool.ImportMedia([abs_video_path])
                
                if not imported_media or len(imported_media) == 0:
                    self.debug_print("Failed to import media for framerate detection")
                    return self.detect_video_framerate(video_path)
                
                media_item = imported_media[0]
                
                # Get the frame rate from clip properties
                try:
                    # Try to get the frame rate property directly
                    fps_str = media_item.GetClipProperty("FPS")
                    
                    if not fps_str:
                        # If direct property failed, try getting properties snapshot
                        all_props = media_item.GetClipProperty()
                        fps_str = all_props.get("FPS", "")
                        
                    if fps_str:
                        # Parse the FPS string
                        try:
                            fps = float(fps_str)
                            self.debug_print(f"Detected framerate from Resolve API: {fps} fps")
                            
                            # Validate the framerate (ensure it's reasonable)
                            if 1 <= fps <= 300:
                                return fps
                        except ValueError:
                            self.debug_print(f"Could not parse Resolve FPS value: {fps_str}")
                except Exception as e:
                    self.debug_print(f"Error getting clip properties from Resolve: {str(e)}")
                    
                # If we've reached here, try to use the timeline frame rate
                try:
                    timeline_fps_str = project.GetSetting("timelineFrameRate")
                    if timeline_fps_str:
                        # Parse the timeline FPS string (might be in format like "29.97 DF")
                        timeline_fps_str = timeline_fps_str.split()[0]  # Remove "DF" if present
                        fps = float(timeline_fps_str)
                        self.debug_print(f"Using timeline framerate from Resolve: {fps} fps")
                        return fps
                except Exception as e:
                    self.debug_print(f"Error getting timeline frame rate from Resolve: {str(e)}")
                    
                # Fallback to FFprobe method
                return self.detect_video_framerate(video_path)
                
            finally:
                # Restore the original media pool folder
                if current_folder:
                    media_pool.SetCurrentFolder(current_folder)
                    
                # Delete imported media item if it exists
                if 'media_item' in locals() and media_item:
                    try:
                        media_pool.DeleteClips([media_item])
                    except:
                        self.debug_print("Couldn't clean up temporary media item")
                        
        except Exception as e:
            self.debug_print(f"Error detecting framerate from Resolve: {str(e)}")
            return self.detect_video_framerate(video_path)
    
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
            
            # Always create the import buttons, but only show the frame if editor is selected
            if show_import_buttons:
                import_buttons_frame.pack(side="right", padx=5, anchor="ne")
            
            # Add Import Media button (always create, will be visible only if frame is visible)
            import_media_btn = ClickableImport(
                import_buttons_frame, 
                "Import Media", 
                result, 
                self._handle_import_media_click,
                tooltip="Import the entire video file to the DaVinci Resolve timeline"
            )
            import_media_btn.pack(side="left", padx=5)
            
            # Add Import Clip button (always create, will be visible only if frame is visible)
            import_clip_btn = ClickableImport(
                import_buttons_frame, 
                "Import Clip", 
                result, 
                self._handle_import_clip_click,
                tooltip="Import only the time range from this subtitle entry to the DaVinci Resolve timeline"
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
            video_info = self.subtitle_to_video_map[subtitle_file]
            video_file = video_info["path"]
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
        try:
            selected_editor = self.editor_var.get()
            self.debug_print(f"Editor changed to: {selected_editor}")
            
            # Update preferences
            self.preferences["selected_editor"] = selected_editor
            self.save_preferences()
            
            # Just update the UI and status - no API loading here
            self._update_import_buttons_visibility()
            
            # Update status
            if selected_editor == "None":
                self.status_var.set("Editor integration disabled")
            else:
                self.status_var.set(f"{selected_editor} selected. API will be initialized when needed.")
            
        except Exception as e:
            error_msg = f"Error changing editor: {str(e)}"
            self.debug_print(error_msg)
            self.status_var.set(f"Error: {error_msg}")

    def _update_import_buttons_visibility(self):
        """Update visibility of import buttons based on selected editor"""
        selected_editor = self.editor_var.get()
        show_import_buttons = selected_editor != "None"
        
        # Loop through all result frames and update import buttons visibility
        for widget in self.results_container.winfo_children():
            # Check if this is a result frame with import_buttons_frame attribute
            if isinstance(widget, ttk.Frame) and hasattr(widget, "import_buttons_frame"):
                # Get the import buttons frame
                import_buttons_frame = widget.import_buttons_frame
                
                if show_import_buttons:
                    # Show the buttons frame
                    import_buttons_frame.pack(side="right", padx=5, anchor="ne")
                else:
                    # Hide the buttons frame
                    import_buttons_frame.pack_forget()
        
        # Update the canvas scroll region
        self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))
    
    def _handle_import_media_click(self, result):
        """Handle click on Import Media button"""
        selected_editor = self.editor_var.get()
        subtitle_file = result['file']
        
        if subtitle_file in self.subtitle_to_video_map:
            video_info = self.subtitle_to_video_map[subtitle_file]
            video_file = video_info["path"]
            
            # We don't need framerate for full media import, but update it if needed
            if video_info["fps"] is None and selected_editor == "DaVinci Resolve":
                try:
                    if hasattr(self, 'resolve_initialized') and self.resolve_initialized:
                        fps = self.detect_video_framerate_from_resolve(video_file)
                        video_info["fps"] = fps
                except Exception:
                    # Silently ignore framerate detection errors for full media import
                    pass
                
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
        
        # Improve timecode handling for better accuracy
        start_time = result['mpc_start_time']
        
        # Parse end time properly - keep milliseconds for more accurate calculations
        end_time = result['end_time']
        if ',' in end_time:
            # Convert comma to period for consistent processing
            end_time = end_time.replace(',', '.')
        
        # Ensure proper formatting for Resolve (no milliseconds in the string)
        display_end_time = end_time.split('.')[0]
        
        self.debug_print(f"Original timecodes: {start_time} to {end_time}")
        
        if subtitle_file in self.subtitle_to_video_map:
            video_info = self.subtitle_to_video_map[subtitle_file]
            video_file = video_info["path"]
            
            # Check if we need to detect framerate
            if video_info["fps"] is None:
                # Detect framerate now - use Resolve API if available
                try:
                    if self.editor_var.get() == "DaVinci Resolve" and hasattr(self, 'resolve_initialized') and self.resolve_initialized:
                        fps = self.detect_video_framerate_from_resolve(video_file)
                    else:
                        fps = self.detect_video_framerate(video_file)
                    
                    # Store the detected framerate for future use
                    video_info["fps"] = fps
                    
                except Exception as e:
                    self.debug_print(f"Error detecting framerate for {os.path.basename(video_file)}: {str(e)}")
                    fps = 24.0  # Default fallback
                    video_info["fps"] = fps
            else:
                # Use cached framerate
                fps = video_info["fps"]
            
            # Apply minimum duration if enabled
            adjusted_start, adjusted_end = self._apply_minimum_duration(start_time, display_end_time, fps)
            
            # Calculate duration in frames to verify timing
            start_frame = self.timecode_to_frames(adjusted_start, fps)
            end_frame = self.timecode_to_frames(adjusted_end, fps)
            duration_frames = end_frame - start_frame
            duration_secs = duration_frames / fps
            
            self.debug_print(f"Import Clip clicked for {os.path.basename(video_file)} at {adjusted_start}-{adjusted_end} with editor {selected_editor}, FPS: {fps}")
            self.debug_print(f"Frame calculation: {start_frame} to {end_frame} ({duration_frames} frames, {duration_secs:.2f} seconds)")
            
            # Call the appropriate import function based on selected editor
            if selected_editor == "DaVinci Resolve":
                self._import_clip_to_davinci_resolve(video_file, adjusted_start, adjusted_end, fps)
            # Add more editors as needed
            
            self.status_var.set(f"Importing clip from {os.path.basename(video_file)} at {adjusted_start}-{adjusted_end} to {selected_editor}")
        else:
            self.debug_print(f"No matching video file found for {os.path.basename(subtitle_file)}")
            self.status_var.set(f"No matching video file found for {os.path.basename(subtitle_file)}")
    
    def _ensure_resolve_ready(self):
        """
        Ensure DaVinci Resolve is ready for use by checking editor selection, 
        initializing API if needed, and handling safety checks.
        
        Returns:
            bool: True if Resolve is ready for use, False otherwise
        """
        # Check if DaVinci Resolve is selected as the editor
        selected_editor = self.editor_var.get()
        if selected_editor != "DaVinci Resolve":
            self.debug_print(f"Wrong editor selected: {selected_editor}")
            self.status_var.set("Please select DaVinci Resolve as the editor first")
            return False
        
        try:
            # If we haven't initialized yet or previously failed, initialize now
            if not hasattr(self, 'resolve_in_safe_mode') or not hasattr(self, 'resolve_initialized'):
                self.resolve_in_safe_mode = False
                self.resolve_initialized = False
            
            # Don't attempt to use the API if we're in safe mode
            if self.resolve_in_safe_mode:
                self.debug_print("Resolve is in safe mode - import functionality disabled")
                self.status_var.set("Cannot import: DaVinci Resolve integration is in safe mode")
                self.show_error_in_gui("DaVinci Resolve Safe Mode", 
                                     "The integration is running in safe mode due to initialization errors.\n\n"
                                     "Import functionality is disabled to prevent crashes.\n\n"
                                     "Please check that DaVinci Resolve is properly installed and running.")
                return False
            
            # Check if API is initialized, if not initialize it now
            if not self.resolve_initialized:
                self.status_var.set("Testing DaVinci Resolve API safety...")
                self.debug_print("Initializing DaVinci Resolve API on first use...")
                
                # First, test in subprocess for safety
                try:
                    subprocess_test_result = self._test_resolve_import_in_subprocess()
                    if not subprocess_test_result:
                        self.resolve_in_safe_mode = True
                        self.status_var.set("DaVinci Resolve API failed safety test - import disabled")
                        self.show_error_in_gui("DaVinci Resolve Error",
                                            "The DaVinci Resolve API failed the safety test.\n\n"
                                            "Import functionality has been disabled for safety.\n\n"
                                            "This usually happens if there is an incompatibility between the\n"
                                            "Python version and the DaVinci Resolve API.")
                        return False
                        
                    self.debug_print("Safety test passed, attempting actual initialization")
                    success = self._init_davinci_resolve_api()
                    
                    if success:
                        self.resolve_initialized = True
                        self.status_var.set("DaVinci Resolve API initialized")
                        return True
                    else:
                        self.resolve_in_safe_mode = True
                        self.status_var.set("Failed to initialize DaVinci Resolve API")
                        self.show_error_in_gui("DaVinci Resolve Error",
                                            "Failed to initialize DaVinci Resolve API.\n\n"
                                            "Please ensure DaVinci Resolve is installed correctly and running.")
                        return False
                        
                except Exception as init_error:
                    self.resolve_in_safe_mode = True
                    error_msg = f"Error initializing DaVinci Resolve API: {str(init_error)}"
                    self.debug_print(error_msg)
                    self.status_var.set(f"Error: {error_msg}")
                    self.show_error_in_gui("DaVinci Resolve Error",
                                         f"Error initializing DaVinci Resolve API:\n\n{str(init_error)}")
                    return False
            
            # If we get here, everything is ready
            return True
            
        except Exception as e:
            self.resolve_in_safe_mode = True
            error_msg = f"Error ensuring DaVinci Resolve readiness: {str(e)}"
            self.debug_print(error_msg)
            self.status_var.set(f"Error: {error_msg}")
            self.show_error_in_gui("DaVinci Resolve Error", f"Error preparing DaVinci Resolve:\n\n{str(e)}")
            return False

    def _import_media_to_davinci_resolve(self, video_file):
        """Import full media file to DaVinci Resolve"""
        # Ensure DaVinci Resolve is ready for use
        if not self._ensure_resolve_ready():
            return
        
        try:
            # Get absolute path to the video file
            abs_video_path = self.get_absolute_path(video_file)
            
            # Call the timeline import function without time range parameters
            success = self.import_clip_to_timeline(abs_video_path)
            
            if success:
                self.debug_print(f"Successfully imported media to DaVinci Resolve timeline")
                self.status_var.set("Media successfully imported to DaVinci Resolve timeline")
            else:
                self.debug_print("Failed to import media to DaVinci Resolve timeline")
                self.status_var.set("Failed to import media to DaVinci Resolve timeline")
        except Exception as e:
            error_msg = f"Error importing media to DaVinci Resolve: {str(e)}"
            self.debug_print(error_msg)
            self.status_var.set(f"Error: {error_msg}")
            self.show_error_in_gui("DaVinci Resolve Error", f"Error importing media:\n\n{str(e)}")
            
            # Enable safe mode to prevent further crashes
            self.resolve_in_safe_mode = True
    
    def _test_resolve_import_in_subprocess(self):
        """Test importing DaVinciResolveScript in a separate process for safety"""
        self.debug_print("Testing DaVinci Resolve import in a separate process...")
        
        # Get the current API path
        api_path = os.environ.get("RESOLVE_SCRIPT_API", "")
        lib_path = os.environ.get("RESOLVE_SCRIPT_LIB", "")
        
        # Find all possible module locations
        module_info = find_module_locations(api_path)
        module_locations = module_info["locations"]
        module_files = module_info["module_paths"]
        
        # Print found paths for debugging
        self.debug_print("Found module locations:")
        for path in module_files:
            self.debug_print(f"  - {path}")
        
        # Add the API path itself and its parent to search paths
        search_paths = []
        if os.path.exists(api_path):
            search_paths.append(api_path)
        if os.path.exists(os.path.dirname(api_path)):
            search_paths.append(os.path.dirname(api_path))
            
        # Add module locations to search paths
        for loc in module_locations:
            if loc not in search_paths and os.path.exists(loc):
                search_paths.append(loc)
        
        self.debug_print(f"Search paths to be used: {search_paths}")
        
        # Create a temporary Python file with the import test
        with tempfile.NamedTemporaryFile(suffix='.py', delete=False, mode='w') as f:
            test_script = f.name
            
            # Write a script that attempts the import and exits with code 0 if successful
            f.write(f'''
import os
import sys
import glob

# Set required environment variables
os.environ["RESOLVE_SCRIPT_API"] = r"{api_path}"
os.environ["RESOLVE_SCRIPT_LIB"] = r"{lib_path}"

# Add all possible module paths to sys.path
search_paths = {search_paths!r}
for path in search_paths:
    if path and path not in sys.path:
        sys.path.append(path)
        print(f"Added {{path}} to Python path")

# Known module files
module_files = {module_files!r}
for module_file in module_files:
    print(f"Known module file: {{module_file}}")

# Print sys.path for debugging
print(f"Python sys.path: {{sys.path}}")

# Try direct import first
try:
    import DaVinciResolveScript
    print("Successfully imported DaVinciResolveScript in test process")
    sys.exit(0)
except ImportError as e:
    print(f"Standard import failed: {{e}}")

# Try alternate import approaches
for module_file in module_files:
    module_dir = os.path.dirname(module_file)
    module_name = os.path.splitext(os.path.basename(module_file))[0]
    
    try:
        # Try adding module dir directly to path and importing
        if module_dir not in sys.path:
            sys.path.append(module_dir)
            print(f"Added module directory {{module_dir}} directly to path")
        
        # Try to import again
        import DaVinciResolveScript
        print(f"Successfully imported DaVinciResolveScript after adding {{module_dir}}")
        sys.exit(0)
    except ImportError as e:
        print(f"Import still failed with {{module_dir}} in path: {{e}}")

# If all previous attempts failed, try more aggressive approaches
# Try to load the module directly using importlib
print("Trying direct module loading with importlib...")
try:
    import importlib.util
    
    for module_file in module_files:
        try:
            print(f"Trying to load {{module_file}} with importlib...")
            spec = importlib.util.spec_from_file_location("DaVinciResolveScript", module_file)
            if spec:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                print(f"Successfully loaded {{module_file}} with importlib")
                sys.path.insert(0, os.path.dirname(module_file))
                sys.exit(0)
        except Exception as e:
            print(f"Importlib loading failed for {{module_file}}: {{e}}")
except Exception as e:
    print(f"Importlib approach failed: {{e}}")

print("All import attempts failed")
sys.exit(1)
''')
        
        try:
            # Run the test script in a separate process
            self.debug_print(f"Running import test script: {test_script}")
            result = subprocess.run(
                [sys.executable, test_script],
                capture_output=True, 
                text=True,
                timeout=10  # Set a timeout to avoid hanging
            )
            
            # Check the result
            if result.returncode == 0:
                self.debug_print("Import test succeeded in subprocess")
                self.debug_print(f"Subprocess stdout: {result.stdout}")
                return True
            else:
                self.debug_print(f"Import test failed in subprocess with exit code {result.returncode}")
                self.debug_print(f"Subprocess stderr: {result.stderr}")
                self.debug_print(f"Subprocess stdout: {result.stdout}")
                return False
                
        except subprocess.TimeoutExpired:
            self.debug_print("Import test timed out - this suggests the import would hang or crash")
            return False
        except Exception as e:
            self.debug_print(f"Error running import test: {e}")
            return False
        finally:
            # Clean up the temporary file
            try:
                os.unlink(test_script)
            except:
                pass

    def import_clip_to_timeline(self, clip_path, start_tc=None, end_tc=None, start_frame=None, end_frame=None, fps=24.0):
        """
        Import a clip to the current timeline with specified in/out points
        
        Args:
            clip_path (str): Path to the clip to import
            start_tc (str, optional): Start timecode in HH:MM:SS format
            end_tc (str, optional): End timecode in HH:MM:SS format
            start_frame (int, optional): Start frame (alternative to start_tc)
            end_frame (int, optional): End frame (alternative to end_frame)
            fps (float, optional): Frames per second of the clip (default: 24.0)
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self.debug_print(f"Importing clip: {clip_path} with FPS: {fps}")
            
            # Convert timecodes to frames if provided
            if start_tc:
                start_frame = self.timecode_to_frames(start_tc, fps)
                self.debug_print(f"Converted start timecode {start_tc} to frame {start_frame} at {fps} fps")
            
            if end_tc:
                end_frame = self.timecode_to_frames(end_tc, fps)
                self.debug_print(f"Converted end timecode {end_tc} to frame {end_frame} at {fps} fps")
            
            # Calculate and log the duration for verification
            if start_frame is not None and end_frame is not None:
                duration_frames = end_frame - start_frame
                duration_seconds = duration_frames / fps
                self.debug_print(f"Clip duration: {duration_frames} frames ({duration_seconds:.2f} seconds)")
            
            # Get Resolve
            resolve = dvr_script.scriptapp("Resolve")
            if not resolve:
                self.debug_print("Failed to get Resolve object")
                return False
            
            # Get project manager
            project_manager = resolve.GetProjectManager()
            if not project_manager:
                self.debug_print("Failed to get project manager")
                return False
            
            # Get current project
            project = project_manager.GetCurrentProject()
            if not project:
                self.debug_print("No project is currently open")
                return False
            
            # Get current project framerate
            try:
                project_fps_str = project.GetSetting("timelineFrameRate")
                if project_fps_str:
                    # Remove "DF" if present
                    if ' ' in project_fps_str:
                        project_fps_str = project_fps_str.split()[0]
                    project_fps = float(project_fps_str)
                    self.debug_print(f"Project timeline framerate: {project_fps} fps")
                    
                    # Log warning if project and clip framerates don't match
                    if abs(project_fps - fps) > 0.01:  # Allow for small floating-point differences
                        self.debug_print(f"WARNING: Project framerate ({project_fps}) does not match clip framerate ({fps})")
            except Exception as e:
                self.debug_print(f"Could not get project framerate: {e}")
            
            self.debug_print(f"Current project: {project.GetName()}")
            
            # Make sure we're on the Edit page
            current_page = resolve.GetCurrentPage()
            if current_page != "edit":
                resolve.OpenPage("edit")
                time.sleep(0.5)
            
            # Get media pool
            media_pool = project.GetMediaPool()
            if not media_pool:
                self.debug_print("Failed to get media pool")
                return False
            
            # Get current timeline
            timeline = project.GetCurrentTimeline()
            if not timeline:
                self.debug_print("No timeline is currently open")
                return False
            
            self.debug_print(f"Current timeline: {timeline.GetName()}")
            
            # Normalize path
            abs_path = os.path.abspath(clip_path)
            if not os.path.exists(abs_path):
                self.debug_print(f"File not found: {abs_path}")
                return False
            
            # Import media
            imported_media = media_pool.ImportMedia([abs_path])
            if not imported_media or len(imported_media) == 0:
                self.debug_print("Failed to import media")
                return False
            
            media_item = imported_media[0]
            self.debug_print(f"Successfully imported: {media_item.GetName()}")
            
            # Get actual media properties for verification
            try:
                media_fps = media_item.GetClipProperty("FPS")
                if media_fps:
                    self.debug_print(f"Imported media actual FPS: {media_fps}")
                    
                    # If the detected FPS is significantly different from the imported media's FPS
                    try:
                        media_fps_float = float(media_fps)
                        if abs(media_fps_float - fps) > 0.5:  # If difference is more than 0.5 fps
                            self.debug_print(f"WARNING: Detected FPS ({fps}) differs from media's actual FPS ({media_fps_float})")
                    except ValueError:
                        pass
            except Exception as e:
                self.debug_print(f"Could not get media FPS: {e}")
            
            # Create clip info dictionary with explicit frame ranges
            clip_info = {
                "mediaPoolItem": media_item
            }
            
            # Add in/out points if specified
            if start_frame is not None:
                clip_info["startFrame"] = start_frame
            
            if end_frame is not None:
                clip_info["endFrame"] = end_frame
            
            self.debug_print(f"Appending clip with time range {start_frame if start_frame is not None else 0} to {end_frame if end_frame is not None else 'end'}")
            
            # Append to timeline with clip info dictionary
            appended_items = media_pool.AppendToTimeline([clip_info])
            
            if appended_items and len(appended_items) > 0:
                self.debug_print(f"Successfully appended clip to timeline with specified time range")
                
                # Get the appended timeline item for verification
                try:
                    timeline_item = appended_items[0]
                    start = timeline_item.GetStart()
                    self.debug_print(f"Timeline item start frame: {start}")
                    
                    # Get the left and right offset capabilities
                    left_offset = timeline_item.GetLeftOffset()
                    right_offset = timeline_item.GetRightOffset()
                    self.debug_print(f"Item can be extended: {left_offset} frames left, {right_offset} frames right")
                except Exception as e:
                    self.debug_print(f"Could not get timeline item details: {e}")
                    
                return True
            else:
                # Fallback to just appending the media item if the clip info approach fails
                self.debug_print("Failed to append clip with time range, trying without time range...")
                appended_items = media_pool.AppendToTimeline([media_item])
                
                if appended_items and len(appended_items) > 0:
                    self.debug_print("Successfully appended clip to timeline (without time range)")
                    return True
                
            self.debug_print("Failed to append clip to timeline")
            return False
                
        except Exception as e:
            self.debug_print(f"Error importing clip: {e}")
            return False
        
    def timecode_to_frames(self, timecode, fps=24.0):
        """
        Convert HH:MM:SS or HH:MM:SS:FF or HH:MM:SS,MMM timecode to frames
        with improved millisecond handling and offset compensation
        
        Args:
            timecode (str): Timecode in HH:MM:SS, HH:MM:SS.MS, HH:MM:SS,MS or HH:MM:SS:FF format
            fps (float): Frames per second (default: 24.0)
        
        Returns:
            int: Frame number
        """
        try:
            self.debug_print(f"Converting timecode {timecode} using {fps} fps")
            
            # Check if we have a DaVinci Resolve style timecode with HH:MM:SS:FF
            if len(timecode.split(':')) == 4:
                hours, minutes, seconds, frames = map(int, timecode.split(':'))
                total_frames = (hours * 3600 + minutes * 60 + seconds) * fps + frames
                return int(total_frames)
            
            # Handle milliseconds which could be comma or period separated
            ms = 0
            if '.' in timecode:
                time_parts, ms_part = timecode.split('.')
                ms = int(ms_part) if ms_part else 0
            elif ',' in timecode:
                time_parts, ms_part = timecode.split(',')
                ms = int(ms_part) if ms_part else 0
            else:
                time_parts = timecode
            
            # Split time parts
            parts = time_parts.split(':')
            if len(parts) == 3:  # HH:MM:SS
                hours, minutes, seconds = map(int, parts)
            elif len(parts) == 2:  # MM:SS
                hours = 0
                minutes, seconds = map(int, parts)
            else:
                self.debug_print(f"Invalid timecode format: {timecode}")
                return 0
            
            # Calculate total seconds with proper millisecond handling
            if ms > 0:
                if ms < 100:  # If ms is less than 100, assume it's already in frames rather than milliseconds
                    frame_portion = ms
                else:  # Convert milliseconds to frames
                    frame_portion = (ms / 1000.0) * fps
            else:
                frame_portion = 0
                
            # Calculate total frames
            total_seconds = hours * 3600 + minutes * 60 + seconds
            total_frames = (total_seconds * fps) + frame_portion
            
            # Apply frame offset compensation for better sync
            # Resolve tends to need a slight offset for accurate positioning
            offset_frames = self._get_timecode_offset(fps)
            compensated_frames = int(total_frames) + offset_frames
            
            self.debug_print(f"Calculated frame position: {int(total_frames)} → {compensated_frames} (with offset {offset_frames})")
            return max(0, compensated_frames)  # Ensure non-negative frame number
        except Exception as e:
            self.debug_print(f"Invalid timecode format: {timecode} - Error: {e}")
            return 0

    def _get_timecode_offset(self, fps):
        """
        Get the appropriate frame offset compensation based on framerate
        
        Args:
            fps (float): Framerate of the video
            
        Returns:
            int: Frame offset to apply
        """
        # These offsets are for fine-tuning the frame calculations
        # Different framerates may need different offsets for perfect accuracy
        if fps >= 59 and fps <= 60:    # 59.94/60 fps
            return 0
        elif fps >= 29 and fps <= 30:  # 29.97/30 fps
            return 0  
        elif fps >= 23 and fps <= 24:  # 23.976/24 fps
            return 0
        elif fps >= 25 and fps <= 25.1: # 25 fps (PAL)
            return 0
        else:
            return 0
    
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
                
        # If the listbox is empty, we don't force add current directory anymore
        # This prevents recursive load_shows() calls when no valid shows exist
        if self.dir_listbox.size() == 0:
            self.debug_print("No directories in listbox, but not forcing current directory")

    def add_directory_warning_logic(self, directory_to_add, existing_directories):
        """
        Check if a directory would be filtered due to intersection rules
        
        Args:
            directory_to_add (str): Directory being considered for addition
            existing_directories (list): Already added directories
            
        Returns:
            tuple: (will_be_filtered, parent_dir) where parent_dir is the existing directory
                that would cause directory_to_add to be filtered out, or None
        """
        normalized_new = os.path.normpath(directory_to_add)
        
        # Check if directory is a subdirectory of an existing directory
        for existing in existing_directories:
            normalized_existing = os.path.normpath(existing)
            if normalized_new.startswith(normalized_existing + os.sep):
                return True, existing
        
        # Check if directory is a parent of any existing directory
        child_dirs = []
        for existing in existing_directories:
            normalized_existing = os.path.normpath(existing)
            if normalized_existing.startswith(normalized_new + os.sep):
                child_dirs.append(existing)
        
        if child_dirs:
            return True, child_dirs
        
        return False, None 
    
    def add_directory(self):
        """Open file dialog to add directories to preferences"""
        # Get main window position and size
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # Calculate centered position for the dialog relative to main window
        dialog_width = 525
        dialog_height = 450  # Increased height to accommodate the current directory status display
        x = root_x + (root_width - dialog_width) // 2
        y = root_y + (root_height - dialog_height) // 2
        
        # Create dialog window
        root = tk.Toplevel(self.root)
        root.title("Add Media Directories")
        root.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        root.transient(self.root)
        root.grab_set()
        
        # Create frame to hold the listbox and buttons
        frame = ttk.Frame(root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Instructions label
        ttk.Label(frame, text="Select one or more directories that contain your media:").pack(pady=5, anchor="w")
        
        # Current directory status frame
        current_dir_frame = ttk.LabelFrame(frame, text="Current Directory Status")
        current_dir_frame.pack(fill="x", pady=5)
        
        # Current directory path display
        current_dir = self.get_current_directory()
        ttk.Label(current_dir_frame, text=f"Current directory: {current_dir}").pack(anchor="w", padx=5, pady=2)
        
        # Current directory inclusion status
        self.current_dir_status_var = tk.StringVar()
        self.current_dir_status_label = ttk.Label(current_dir_frame, textvariable=self.current_dir_status_var, font=("TkDefaultFont", 10))
        self.current_dir_status_label.pack(anchor="w", padx=5, pady=2)
        
        # Function to update current directory status display
        def update_current_dir_status():
            if self.preferences.get("exclude_current_dir", False):
                self.current_dir_status_var.set("❌ Current directory is EXCLUDED from automatic scanning")
                self.current_dir_status_label.config(foreground="red")
            else:
                self.current_dir_status_var.set("✓ Current directory is INCLUDED in automatic scanning")
                self.current_dir_status_label.config(foreground="green")
        
        # Initial update of status
        update_current_dir_status()
        
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
            
            if new_dir:
                # Get existing directories from the listbox
                existing_dirs = [listbox.get(i) for i in range(listbox.size())]
                
                # Check if it's the current directory - if so, suggest using the toggle instead
                if os.path.normpath(new_dir) == os.path.normpath(current_dir):
                    message = ("This is your current working directory.\n\n"
                            "Instead of adding it manually, please use the 'Toggle Current Directory' button "
                            "to include/exclude it from automatic scanning.")
                    messagebox.showinfo("Current Directory Selected", message)
                    return
                
                # Check if the directory is already in the list
                if new_dir in existing_dirs:
                    self.debug_print(f"Directory already exists in list: {new_dir}")
                    messagebox.showinfo("Directory Exists", "This directory is already in the list.")
                    return
                    
                # Check if this directory would be filtered
                will_be_filtered, reason = self.add_directory_warning_logic(new_dir, existing_dirs)
                
                if will_be_filtered:
                    if isinstance(reason, list):
                        # This is a parent directory of existing directory/directories
                        child_dirs_str = "\n• ".join(reason)
                        message = (f"The directory you selected is a parent of one or more existing directories:\n\n• {child_dirs_str}\n\n"
                                f"If you add this directory, the existing directories will be filtered out since they're already covered by the parent.")
                        add_anyway = messagebox.askyesno("Directory Conflict", message + "\n\nAdd parent directory anyway?")
                        
                        if add_anyway:
                            # First remove the child directories
                            for child in reason:
                                for i in range(listbox.size()):
                                    if listbox.get(i) == child:
                                        listbox.delete(i)
                                        self.debug_print(f"Removed child directory: {child}")
                                        break
                            
                            # Now add the new directory
                            listbox.insert(tk.END, new_dir)
                            self.debug_print(f"Added parent directory: {new_dir}")
                    else:
                        # This is a subdirectory of an existing directory
                        message = (f"The directory you selected is inside an existing directory:\n\n• {reason}\n\n"
                                f"It will be filtered out during loading since its parent directory will already be scanned.")
                        add_anyway = messagebox.askyesno("Directory Conflict", message + "\n\nAdd subdirectory anyway?")
                        
                        if add_anyway:
                            listbox.insert(tk.END, new_dir)
                            self.debug_print(f"Added subdirectory despite warning: {new_dir}")
                else:
                    # No conflicts, add normally
                    listbox.insert(tk.END, new_dir)
                    self.debug_print(f"Added directory without conflicts: {new_dir}")
        
        # Function to toggle current directory inclusion
        def toggle_current_dir():
            # Toggle the exclude_current_dir preference
            current_exclude = self.preferences.get("exclude_current_dir", False)
            self.preferences["exclude_current_dir"] = not current_exclude
            
            # Update the status display
            update_current_dir_status()
            
            # Log the change
            if current_exclude:
                self.debug_print("Current directory inclusion enabled")
            else:
                self.debug_print("Current directory inclusion disabled")
        
        # Function to remove selected directories from the listbox
        def remove_selected():
            selected = list(listbox.curselection())
            selected.reverse()  # Reverse to delete from bottom to top
            for i in selected:
                self.debug_print(f"Removed directory: {listbox.get(i)}")
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
            
            # Check if the list of directories has changed
            if set(new_dirs) != set(existing_dirs):
                changes_made = True
                self.preferences["directories"] = new_dirs
            
            # Save and update if changes were made
            if changes_made or self.preferences.get("exclude_current_dir", False) != self.original_exclude_current_dir:
                self.save_preferences()
                self.debug_print(f"Directory dialog - saved preferences: {self.preferences}")
                self.update_directory_listbox()
                
                # Clear existing show map and reload everything
                self.debug_print("Directory dialog - clearing show map and reloading shows")
                self.show_name_to_path_map.clear()
                
                # Reload shows and remap files
                self.debug_print("Directory dialog - reloading shows after directory changes")
                shows_paths = self.load_shows()
                
                # Check if we loaded any shows
                if len(self.show_name_to_path_map) > 0:
                    self.debug_print(f"Directory dialog - loaded shows: {len(shows_paths)}, names: {list(self.show_name_to_path_map.keys())}")
                    self.map_subtitles_to_videos()
                    
                    # Force the dropdown to update with the new values
                    self.debug_print(f"Directory dialog - updating dropdown with {len(self.show_name_to_path_map)} shows")
                    self.update_show_dropdown()
                    
                    self.status_var.set(f"Added {len(new_dirs) - len(existing_dirs)} directories. Found {len(self.show_name_to_path_map)} shows")
                else:
                    self.debug_print("Directory dialog - no shows found after adding directories")
                    self.status_var.set("No shows found in selected directories. Please ensure they contain subtitle files.")
                    
                    # If we still don't have shows after selecting directories, show guidance again
                    if len(self.show_name_to_path_map) == 0:
                        self.root.after(500, self._delayed_show_guidance)
            else:
                self.status_var.set("No changes made to media directories")
                
                # If no changes were made and we still have no shows, show guidance again
                if len(self.show_name_to_path_map) == 0:
                    self.root.after(500, self._delayed_show_guidance)
            
            # Close the dialog
            root.destroy()
        
        # Store original exclude_current_dir value to detect changes
        self.original_exclude_current_dir = self.preferences.get("exclude_current_dir", False)
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=10)
        
        # Buttons - REORDERED as requested
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
                    
                # Check if it's the current directory
                if os.path.normpath(new_dir) == os.path.normpath(current_dir):
                    message = ("This is your current working directory.\n\n"
                            "Instead of adding it manually, please use the 'Toggle Current Directory' button "
                            "to include/exclude it from automatic scanning.")
                    messagebox.showinfo("Current Directory Selected", message)
                    continue
                    
                # Check if the directory is already in the list
                existing_dirs = [listbox.get(i) for i in range(listbox.size())]
                if new_dir in existing_dirs:
                    self.debug_print(f"Directory already exists in list: {new_dir}")
                    messagebox.showinfo("Directory Exists", "This directory is already in the list.")
                    continue
                    
                # Check if this directory would be filtered
                will_be_filtered, reason = self.add_directory_warning_logic(new_dir, existing_dirs)
                
                if will_be_filtered:
                    if isinstance(reason, list):
                        # This is a parent directory of existing directory/directories
                        child_dirs_str = "\n• ".join(reason)
                        message = (f"The directory you selected is a parent of one or more existing directories:\n\n• {child_dirs_str}\n\n"
                                f"If you add this directory, the existing directories will be filtered out since they're already covered by the parent.")
                        add_anyway = messagebox.askyesno("Directory Conflict", message + "\n\nAdd parent directory anyway?")
                        
                        if add_anyway:
                            # First remove the child directories
                            for child in reason:
                                for i in range(listbox.size()):
                                    if listbox.get(i) == child:
                                        listbox.delete(i)
                                        self.debug_print(f"Removed child directory: {child}")
                                        break
                            
                            # Now add the new directory
                            listbox.insert(tk.END, new_dir)
                            added_dirs.append(new_dir)
                            self.debug_print(f"Added parent directory: {new_dir}")
                    else:
                        # This is a subdirectory of an existing directory
                        message = (f"The directory you selected is inside an existing directory:\n\n• {reason}\n\n"
                                f"It will be filtered out during loading since its parent directory will already be scanned.")
                        add_anyway = messagebox.askyesno("Directory Conflict", message + "\n\nAdd subdirectory anyway?")
                        
                        if add_anyway:
                            listbox.insert(tk.END, new_dir)
                            added_dirs.append(new_dir)
                            self.debug_print(f"Added subdirectory despite warning: {new_dir}")
                else:
                    # No conflicts, add normally
                    listbox.insert(tk.END, new_dir)
                    added_dirs.append(new_dir)
                    self.debug_print(f"Added directory without conflicts: {new_dir}")
                    
            # Show the dialog again
            root.deiconify()
            
            # Report how many directories were added
            if added_dirs:
                self.debug_print(f"Added {len(added_dirs)} directories through multiple selection")
        
        # Add button for multiple directory selection - REORDERED
        multi_select_btn = ttk.Button(button_frame, text="Add Multiple Directories", command=select_multiple_dirs)
        multi_select_btn.pack(side="left", padx=5)
        
        # Add button for toggling current directory - REORDERED and RENAMED
        current_dir_btn = ttk.Button(button_frame, text="Toggle Current Directory", command=toggle_current_dir)
        current_dir_btn.pack(side="left", padx=5)
        
        remove_btn = ttk.Button(button_frame, text="Remove Selected", command=remove_selected)
        remove_btn.pack(side="left", padx=5)
        
        # Bottom button frame
        bottom_button_frame = ttk.Frame(frame)
        bottom_button_frame.pack(fill="x", pady=10)
        
        # Function to handle cancel button
        def on_cancel():
            root.destroy()
            # Only show guidance dialog if no shows exist and not already showing
            if len(self.show_name_to_path_map) == 0 and not self.guidance_dialog_showing:
                self.root.after(500, self._delayed_show_guidance)
        
        cancel_btn = ttk.Button(bottom_button_frame, text="Cancel", command=on_cancel)
        cancel_btn.pack(side="right", padx=5)
        
        save_btn = ttk.Button(bottom_button_frame, text="Save and Close", command=save_and_close)
        save_btn.pack(side="right", padx=5)
        
        # Log the position for debugging
        self.debug_print(f"Created directory dialog at position ({x},{y}) with size {dialog_width}x{dialog_height}")
        
        # Make the dialog modal
        root.transient(self.root)
        root.grab_set()
        self.root.wait_window(root)
    
    def remove_directory(self):
        """Remove selected directory from preferences or toggle current directory inclusion"""
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
        
        # Handle removal of current directory by toggling the exclude_current_dir preference
        if is_current:
            # Toggle the exclude_current_dir preference
            self.preferences["exclude_current_dir"] = True
            self.save_preferences()
            self.update_directory_listbox()
            
            # Clear existing show map and reload everything
            self.show_name_to_path_map.clear()
            
            # Reload shows and remap files
            shows_paths = self.load_shows()
            self.map_subtitles_to_videos()
            
            # Update status message
            if len(self.show_name_to_path_map) > 0:
                self.status_var.set(f"Current directory excluded. Found {len(self.show_name_to_path_map)} shows")
            else:
                self.status_var.set("Current directory excluded. No shows found.")
                # If no shows are found, show the guidance dialog
                self.root.after(500, self._delayed_show_guidance)
                
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
            
            # Update status and show guidance if needed
            if len(self.show_name_to_path_map) > 0:
                self.status_var.set(f"Removed directory: {selected_dir}. Found {len(self.show_name_to_path_map)} shows")
            else:
                self.status_var.set(f"Removed directory: {selected_dir}. No shows found.")
                # If no shows are found, show the guidance dialog
                self.root.after(500, self._delayed_show_guidance)
        else:
            self.status_var.set("Directory not found in preferences")

    def _init_davinci_resolve_api(self):
        """Initialize the DaVinci Resolve API"""
        try:
            self.debug_print("Initializing DaVinci Resolve API...")
            
            # Get configuration file path
            config_file = os.path.join(self.script_dir, RESOLVE_PATHS_FILENAME)
            config = {}
            modified = False
            
            # Load saved paths if available
            if os.path.exists(config_file):
                try:
                    with open(config_file, 'r') as f:
                        config = json.load(f)
                        self.debug_print(f"Loaded config from: {config_file}")
                        if "RESOLVE_SCRIPT_API" in config:
                            self.debug_print(f"Config API path: {config['RESOLVE_SCRIPT_API']}")
                        if "RESOLVE_SCRIPT_LIB" in config:
                            self.debug_print(f"Config LIB path: {config['RESOLVE_SCRIPT_LIB']}")
                except Exception as e:
                    self.debug_print(f"Failed to load config file: {str(e)}")
            
            # Get default API path based on OS
            if sys.platform.startswith("win"):  # Windows
                default_api_path = r"C:\ProgramData\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting"
            elif sys.platform == "darwin":  # macOS
                default_api_path = r"/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
            elif sys.platform.startswith("linux"):  # Linux
                default_api_path = r"/opt/resolve/Developer/Scripting"
            else:
                default_api_path = ""
            
            # Get default LIB path based on OS
            if sys.platform.startswith("win"):  # Windows
                default_lib_path = r"C:\Program Files\Blackmagic Design\DaVinci Resolve\fusionscript.dll"
            elif sys.platform == "darwin":  # macOS
                default_lib_path = r"/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"
            elif sys.platform.startswith("linux"):  # Linux
                default_lib_path = r"/opt/resolve/libs/Fusion/fusionscript.so"
            else:
                default_lib_path = ""
            
            self.debug_print(f"Default API path: {default_api_path}")
            self.debug_print(f"Default LIB path: {default_lib_path}")
            
            # Check if default paths are valid
            default_module_paths = find_module_locations(default_api_path)
            default_api_valid = len(default_module_paths["module_paths"]) > 0
            default_lib_valid = os.path.isfile(default_lib_path)
            
            self.debug_print(f"Default module file exists: {default_api_valid}")
            self.debug_print(f"Default library file exists: {default_lib_valid}")
            
            # Use a dialog to get custom paths if needed
            need_custom_paths = False
            
            # Set initial API path from config or env
            api_path = None
            if "RESOLVE_SCRIPT_API" in config:
                api_path = config["RESOLVE_SCRIPT_API"]
                os.environ["RESOLVE_SCRIPT_API"] = api_path
                self.debug_print(f"Using API path from config: {api_path}")
            elif os.getenv("RESOLVE_SCRIPT_API"):
                api_path = os.getenv("RESOLVE_SCRIPT_API")
                self.debug_print(f"Using existing API path from env: {api_path}")
            else:
                api_path = default_api_path
                os.environ["RESOLVE_SCRIPT_API"] = api_path
                self.debug_print(f"Using default API path: {api_path}")
                
            # Set initial LIB path from config or env
            lib_path = None
            if "RESOLVE_SCRIPT_LIB" in config:
                lib_path = config["RESOLVE_SCRIPT_LIB"]
                os.environ["RESOLVE_SCRIPT_LIB"] = lib_path
                self.debug_print(f"Using LIB path from config: {lib_path}")
            elif os.getenv("RESOLVE_SCRIPT_LIB"):
                lib_path = os.getenv("RESOLVE_SCRIPT_LIB")
                self.debug_print(f"Using existing LIB path from env: {lib_path}")
            else:
                lib_path = default_lib_path
                os.environ["RESOLVE_SCRIPT_LIB"] = lib_path
                self.debug_print(f"Using default LIB path: {lib_path}")
            
            # Check if module file exists at any possible location
            module_info = find_module_locations(api_path)
            module_exists = len(module_info["module_paths"]) > 0
            
            # Print all found locations
            self.debug_print("Checking possible module locations:")
            for path in module_info["module_paths"]:
                self.debug_print(f"  - {path}: Found")
            
            # Check if library file exists
            lib_file_exists = os.path.isfile(lib_path)
            self.debug_print(f"Library exists at {lib_path}: {lib_file_exists}")
            
            # If the module or library is missing, show dialog to input paths
            if not module_exists or not lib_file_exists:
                need_custom_paths = True
                
                # Create a dialog to get the custom paths
                self.debug_print("Need to get custom paths from user")
                # Show dialog to get paths
                result = self._show_resolve_paths_dialog(
                    api_path, lib_path, 
                    default_api_path, default_lib_path,
                    default_api_valid, default_lib_valid,
                    module_exists, lib_file_exists
                )
                
                # Process the result
                if not result['success']:
                    self.debug_print("User cancelled path configuration")
                    return False
                
                # Update paths if changed
                if result['api_path'] != api_path:
                    api_path = result['api_path']
                    os.environ["RESOLVE_SCRIPT_API"] = api_path
                    config["RESOLVE_SCRIPT_API"] = api_path
                    modified = True
                    self.debug_print(f"Updated API path: {api_path}")
                
                if result['lib_path'] != lib_path:
                    lib_path = result['lib_path']
                    os.environ["RESOLVE_SCRIPT_LIB"] = lib_path
                    config["RESOLVE_SCRIPT_LIB"] = lib_path
                    modified = True
                    self.debug_print(f"Updated LIB path: {lib_path}")
            
            # Add module paths to sys.path
            self.debug_print("========== FINAL PATH CONFIGURATION ==========")
            self.debug_print(f"Using RESOLVE_SCRIPT_API: {api_path}")
            self.debug_print(f"Using RESOLVE_SCRIPT_LIB: {lib_path}")
            
            # Add all module locations to sys.path
            module_info = find_module_locations(api_path)
            for path in module_info["locations"]:
                if path not in sys.path:
                    sys.path.append(path)
                    self.debug_print(f"Added to Python path: {path}")
                    
            # Also add the API path itself if not already added
            if api_path and api_path not in sys.path and os.path.exists(api_path):
                sys.path.append(api_path)
                self.debug_print(f"Added API path to Python path: {api_path}")
            
            self.debug_print("=============================================")
            
            # Save config if modified
            if modified:
                try:
                    # If config is empty, delete the file instead
                    if not config:
                        if os.path.exists(config_file):
                            os.remove(config_file)
                            self.debug_print(f"Removed empty config file: {config_file}")
                    else:
                        with open(config_file, 'w') as f:
                            json.dump(config, f, indent=2)
                            self.debug_print(f"Saved custom paths to {config_file}")
                except Exception as e:
                    self.debug_print(f"Failed to save config file: {str(e)}")
            
            # Attempt to import the module
            try:
                global dvr_script
                self.debug_print("Attempting to import DaVinciResolveScript...")
                import DaVinciResolveScript as dvr_script
                self.debug_print("Successfully imported DaVinciResolveScript")
                return True
            except ImportError as e:
                error_msg = f"Failed to import DaVinciResolveScript: {str(e)}"
                self.debug_print(error_msg)
                self.show_error_in_gui("DaVinci Resolve Import Error", 
                                      f"Failed to import DaVinciResolveScript module.\n\n{str(e)}\n\nPlease check your DaVinci Resolve installation.")
                return False
                
        except Exception as e:
            error_msg = f"Error initializing DaVinci Resolve API: {str(e)}"
            self.debug_print(error_msg)
            self.show_error_in_gui("DaVinci Resolve Error", 
                                  f"Error initializing DaVinci Resolve API:\n\n{str(e)}")
            return False
            
    def _show_resolve_paths_dialog(self, current_api, current_lib, default_api, default_lib, 
                                 default_api_valid, default_lib_valid, module_exists, lib_exists):
        """Show a dialog to get custom paths for DaVinci Resolve scripting"""
        dialog = tk.Toplevel(self.root)
        dialog.title("DaVinci Resolve Scripting Setup")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set the result
        result = {
            'success': False,
            'api_path': current_api,
            'lib_path': current_lib
        }
        
        # Create main frame
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Title and explanation
        ttk.Label(main_frame, text="DaVinci Resolve Scripting Setup", font=("TkDefaultFont", 14, "bold")).pack(pady=5)
        
        explanation_text = "The application needs to know where DaVinci Resolve scripting files are located."
        if not module_exists:
            explanation_text += "\n\nThe script module file (DaVinciResolveScript.py) was not found."
        if not lib_exists:
            explanation_text += "\n\nThe script library file (fusionscript.dll) was not found."
        
        explanation = ttk.Label(main_frame, text=explanation_text, wraplength=550, justify="left")
        explanation.pack(pady=10, fill="x")
        
        # Create API path input frame
        api_frame = ttk.LabelFrame(main_frame, text="API Script Directory", padding=10)
        api_frame.pack(fill="x", expand=False, pady=5)
        
        ttk.Label(api_frame, text="Directory containing DaVinciResolveScript.py (or its Modules subfolder):").pack(anchor="w")
        
        api_path_frame = ttk.Frame(api_frame)
        api_path_frame.pack(fill="x", expand=True, pady=5)
        
        api_path_var = tk.StringVar(value=current_api)
        api_path_entry = ttk.Entry(api_path_frame, textvariable=api_path_var, width=50)
        api_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        def browse_api():
            path = filedialog.askdirectory(
                title="Select API Directory",
                initialdir=os.path.dirname(api_path_var.get()) if os.path.exists(os.path.dirname(api_path_var.get())) else None
            )
            if path:
                api_path_var.set(path)
                check_api_path_valid()
        
        api_browse_btn = ttk.Button(api_path_frame, text="Browse...", command=browse_api)
        api_browse_btn.pack(side="right")
        
        # Add a label to show if the path is valid
        api_valid_var = tk.StringVar(value="")
        api_valid_label = ttk.Label(api_frame, textvariable=api_valid_var)
        api_valid_label.pack(anchor="w", pady=5)
        
        def check_api_path_valid():
            path = api_path_var.get()
            if not path:
                api_valid_var.set("Error: No path specified")
                api_valid_label.config(foreground="red")
                return False
                
            # Check if path exists
            if not os.path.exists(path):
                api_valid_var.set("Error: Directory does not exist")
                api_valid_label.config(foreground="red")
                return False
                
            # Check if module exists
            module_info = find_module_locations(path)
            if len(module_info["module_paths"]) > 0:
                found_at = module_info["module_paths"][0]
                api_valid_var.set(f"Found at: {found_at}")
                api_valid_label.config(foreground="green")
                return True
            else:
                api_valid_var.set("Error: DaVinciResolveScript.py not found at this location")
                api_valid_label.config(foreground="red")
                return False
        
        # Default button
        def use_default_api():
            api_path_var.set(default_api)
            check_api_path_valid()
        
        api_default_btn = ttk.Button(api_frame, text="Use Default Path", command=use_default_api)
        api_default_btn.pack(anchor="w", pady=5)
        api_default_btn.config(state="normal" if default_api_valid else "disabled")
        
        # Create Library path input frame
        lib_frame = ttk.LabelFrame(main_frame, text="Script Library File", padding=10)
        lib_frame.pack(fill="x", expand=False, pady=5)
        
        ttk.Label(lib_frame, text="Path to fusionscript.dll/.so:").pack(anchor="w")
        
        lib_path_frame = ttk.Frame(lib_frame)
        lib_path_frame.pack(fill="x", expand=True, pady=5)
        
        lib_path_var = tk.StringVar(value=current_lib)
        lib_path_entry = ttk.Entry(lib_path_frame, textvariable=lib_path_var, width=50)
        lib_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        def browse_lib():
            path = filedialog.askopenfilename(
                title="Select Library File",
                initialdir=os.path.dirname(lib_path_var.get()) if os.path.exists(os.path.dirname(lib_path_var.get())) else None,
                filetypes=[("Library Files", "*.dll *.so"), ("All Files", "*.*")]
            )
            if path:
                lib_path_var.set(path)
                check_lib_path_valid()
        
        lib_browse_btn = ttk.Button(lib_path_frame, text="Browse...", command=browse_lib)
        lib_browse_btn.pack(side="right")
        
        # Add a label to show if the path is valid
        lib_valid_var = tk.StringVar(value="")
        lib_valid_label = ttk.Label(lib_frame, textvariable=lib_valid_var)
        lib_valid_label.pack(anchor="w", pady=5)
        
        def check_lib_path_valid():
            path = lib_path_var.get()
            if not path:
                lib_valid_var.set("Error: No path specified")
                lib_valid_label.config(foreground="red")
                return False
                
            # Check if path exists
            if os.path.isfile(path):
                lib_valid_var.set("File exists")
                lib_valid_label.config(foreground="green")
                return True
            else:
                lib_valid_var.set("Error: File not found")
                lib_valid_label.config(foreground="red")
                return False
        
        # Default button
        def use_default_lib():
            lib_path_var.set(default_lib)
            check_lib_path_valid()
        
        lib_default_btn = ttk.Button(lib_frame, text="Use Default Path", command=use_default_lib)
        lib_default_btn.pack(anchor="w", pady=5)
        lib_default_btn.config(state="normal" if default_lib_valid else "disabled")
        
        # Check initial validity
        check_api_path_valid()
        check_lib_path_valid()
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x", pady=10)
        
        def on_cancel():
            dialog.destroy()
            
        def on_ok():
            # Validate both paths
            api_valid = check_api_path_valid()
            lib_valid = check_lib_path_valid()
            
            if not api_valid or not lib_valid:
                messagebox.showerror("Invalid Paths", 
                                    "One or both paths are invalid. Please provide valid paths or cancel.")
                return
            
            # Set result
            result['success'] = True
            result['api_path'] = api_path_var.get()
            result['lib_path'] = lib_path_var.get()
            dialog.destroy()
        
        cancel_btn = ttk.Button(buttons_frame, text="Cancel", command=on_cancel)
        cancel_btn.pack(side="right", padx=5)
        
        ok_btn = ttk.Button(buttons_frame, text="OK", command=on_ok)
        ok_btn.pack(side="right", padx=5)
        
        # Wait for dialog to close
        self.root.wait_window(dialog)
        
        return result
        
    def _import_clip_to_davinci_resolve(self, video_file, start_time, end_time, fps=24.0):
        """Import clip with time range to DaVinci Resolve"""
        # Ensure DaVinci Resolve is ready for use
        if not self._ensure_resolve_ready():
            return
            
        try:
            # Get absolute path to the video file
            abs_video_path = self.get_absolute_path(video_file)
            
            # Call the timeline import function with the detected framerate
            success = self.import_clip_to_timeline(
                abs_video_path, 
                start_tc=start_time, 
                end_tc=end_time, 
                fps=fps
            )
            
            if success:
                self.debug_print(f"Successfully imported clip to DaVinci Resolve timeline")
                self.status_var.set("Clip successfully imported to DaVinci Resolve timeline")
            else:
                self.debug_print("Failed to import clip to DaVinci Resolve timeline")
                self.status_var.set("Failed to import clip to DaVinci Resolve timeline")
        except Exception as e:
            error_msg = f"Error importing clip to DaVinci Resolve: {str(e)}"
            self.debug_print(error_msg)
            self.status_var.set(f"Error: {error_msg}")
            self.show_error_in_gui("DaVinci Resolve Error", f"Error importing clip:\n\n{str(e)}")
            
            # Enable safe mode to prevent further crashes
            self.resolve_in_safe_mode = True

    def _apply_minimum_duration(self, start_time, end_time, fps):
        """
        Apply minimum duration to a clip by extending it if needed
        
        Args:
            start_time (str): Start timecode in HH:MM:SS format
            end_time (str): End timecode in HH:MM:SS format
            fps (float): Frames per second
            
        Returns:
            tuple: (adjusted_start_time, adjusted_end_time) as strings
        """
        if not self.preferences.get("min_duration_enabled", True):  # Changed default to True
            return start_time, end_time
            
        min_seconds = self.preferences.get("min_duration_seconds", 10.0)
        
        # Convert times to seconds for easier math
        start_seconds = self._timecode_to_seconds(start_time)
        end_seconds = self._timecode_to_seconds(end_time)
        
        # Calculate current duration
        current_duration = end_seconds - start_seconds
        
        # If already meeting minimum, return unchanged
        if current_duration >= min_seconds:
            return start_time, end_time
        
        # Calculate how much time to add
        additional_time_needed = min_seconds - current_duration
        
        # Distribute the additional time evenly on both sides
        time_to_add_each_side = additional_time_needed / 2.0
        
        # Adjust start and end times
        new_start_seconds = max(0, start_seconds - time_to_add_each_side)
        new_end_seconds = end_seconds + time_to_add_each_side
        
        # If we couldn't add enough at the start (because we hit 0), add more to the end
        if new_start_seconds > 0 and start_seconds - time_to_add_each_side < 0:
            shortfall = abs(start_seconds - time_to_add_each_side)
            new_end_seconds += shortfall
            self.debug_print(f"Hit start boundary, adding extra {shortfall:.1f}s to end")
        
        # Convert back to timecodes
        new_start_time = self._seconds_to_timecode(new_start_seconds)
        new_end_time = self._seconds_to_timecode(new_end_seconds)
        
        self.debug_print(f"Applied minimum duration: {start_time}-{end_time} ({current_duration:.1f}s) → {new_start_time}-{new_end_time} ({min_seconds:.1f}s)")
        
        return new_start_time, new_end_time

    def _timecode_to_seconds(self, timecode):
        """Convert HH:MM:SS or HH:MM:SS.MS timecode to seconds"""
        try:
            # Handle milliseconds which could be comma or period separated
            if '.' in timecode:
                time_parts, ms_part = timecode.split('.')
                ms = int(ms_part) if ms_part else 0
            elif ',' in timecode:
                time_parts, ms_part = timecode.split(',')
                ms = int(ms_part) if ms_part else 0
            else:
                time_parts = timecode
                ms = 0
            
            # Split time parts
            parts = time_parts.split(':')
            if len(parts) == 3:  # HH:MM:SS
                hours, minutes, seconds = map(int, parts)
            elif len(parts) == 2:  # MM:SS
                hours = 0
                minutes, seconds = map(int, parts)
            else:
                return 0
            
            # Calculate total seconds
            total_seconds = hours * 3600 + minutes * 60 + seconds
            if ms > 0:
                total_seconds += ms / 1000.0
                
            return total_seconds
        except Exception:
            return 0

    def _seconds_to_timecode(self, total_seconds):
        """Convert seconds to HH:MM:SS format"""
        try:
            # Ensure non-negative
            total_seconds = max(0, total_seconds)
            
            # Calculate hours, minutes, seconds
            hours = int(total_seconds // 3600)
            minutes = int((total_seconds % 3600) // 60)
            seconds = int(total_seconds % 60)
            
            # Format as HH:MM:SS
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        except Exception:
            return "00:00:00"

    # Add a method to show the settings dialog
    def _show_settings_dialog(self):
        """Show a dialog with minimum duration settings"""
        settings_dialog = tk.Toplevel(self.root)
        settings_dialog.title("Minimum Import Duration")
        settings_dialog.geometry("400x250")
        settings_dialog.transient(self.root)
        settings_dialog.grab_set()
        
        # Make dialog modal
        settings_dialog.focus_set()
        
        # Create main frame with padding
        main_frame = ttk.Frame(settings_dialog, padding=15)
        main_frame.pack(fill="both", expand=True)
        
        # Title label
        ttk.Label(main_frame, text="Minimum Clip Duration Settings", 
                 font=("TkDefaultFont", 12, "bold")).pack(anchor="w", pady=(0, 10))
        
        # Enable checkbox
        self.min_duration_var = tk.BooleanVar(value=self.preferences.get("min_duration_enabled", True))
        min_duration_cb = ttk.Checkbutton(
            main_frame, 
            text="Enable minimum duration for imported clips", 
            variable=self.min_duration_var
        )
        min_duration_cb.pack(anchor="w", pady=5)
        
        # Duration input frame
        duration_input_frame = ttk.Frame(main_frame)
        duration_input_frame.pack(fill="x", pady=5)
        
        ttk.Label(duration_input_frame, text="Minimum duration:").pack(side="left")
        
        self.min_duration_seconds_var = tk.StringVar(value=str(self.preferences.get("min_duration_seconds", 10.0)))
        min_duration_entry = ttk.Entry(
            duration_input_frame, 
            textvariable=self.min_duration_seconds_var,
            width=5
        )
        min_duration_entry.pack(side="left", padx=5)
        
        ttk.Label(duration_input_frame, text="seconds").pack(side="left")
        
        # Description label
        description_frame = ttk.LabelFrame(main_frame, text="Description", padding=10)
        description_frame.pack(fill="x", pady=10, expand=True)
        
        description_text = ("Clips shorter than the minimum duration will be extended equally on both sides when imported.\n\n"
                           "If a clip cannot be extended to the full minimum duration due to reaching the start or end of "
                           "the source media, it will be extended as much as possible.")
        
        ttk.Label(description_frame, text=description_text, wraplength=350, justify="left").pack(fill="both")
        
        # Buttons frame
        buttons_frame = ttk.Frame(settings_dialog)
        buttons_frame.pack(fill="x", padx=15, pady=15)
        
        # Cancel button
        cancel_btn = ttk.Button(
            buttons_frame, 
            text="Cancel", 
            command=settings_dialog.destroy
        )
        cancel_btn.pack(side="right", padx=5)
        
        # Apply button
        apply_btn = ttk.Button(
            buttons_frame, 
            text="Apply", 
            command=lambda: self._apply_settings(settings_dialog)
        )
        apply_btn.pack(side="right", padx=5)
        
        # Center the dialog on the main window
        settings_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() / 2) - (settings_dialog.winfo_width() / 2)
        y = self.root.winfo_y() + (self.root.winfo_height() / 2) - (settings_dialog.winfo_height() / 2)
        settings_dialog.geometry(f"+{int(x)}+{int(y)}")

    # Add a method to apply the settings
    def _apply_settings(self, dialog):
        """Apply settings from the dialog"""
        try:
            # Get the minimum duration settings
            enabled = self.min_duration_var.get()
            
            # Parse and validate seconds value
            seconds_str = self.min_duration_seconds_var.get()
            try:
                seconds = float(seconds_str)
                if seconds < 0:
                    seconds = 0.0
                    self.min_duration_seconds_var.set("0.0")
                elif seconds > 60:
                    seconds = 60.0
                    self.min_duration_seconds_var.set("60.0")
            except ValueError:
                # If not a valid float, reset to default
                seconds = 10.0
                self.min_duration_seconds_var.set("10.0")
            
            # Update preferences
            self.preferences["min_duration_enabled"] = enabled
            self.preferences["min_duration_seconds"] = seconds
            self.save_preferences()
            
            self.debug_print(f"Minimum duration settings updated - enabled: {enabled}, seconds: {seconds}")
            
            # Close the dialog
            dialog.destroy()
            
            # Update status
            self.status_var.set("Settings updated successfully")
        except Exception as e:
            self.debug_print(f"Error updating settings: {e}")
            self.status_var.set(f"Error updating settings: {e}")

    # Add a new method for general settings
    def _show_general_settings_dialog(self):
        """Show a dialog with general application settings"""
        settings_dialog = tk.Toplevel(self.root)
        settings_dialog.title("General Settings")
        settings_dialog.geometry("400x300")
        settings_dialog.transient(self.root)
        settings_dialog.grab_set()
        
        # Make dialog modal
        settings_dialog.focus_set()
        
        # Create main frame with padding
        main_frame = ttk.Frame(settings_dialog, padding=15)
        main_frame.pack(fill="both", expand=True)
        
        # Title label
        ttk.Label(main_frame, text="General Application Settings", 
                 font=("TkDefaultFont", 12, "bold")).pack(anchor="w", pady=(0, 10))
        
        # Placeholder for future general settings
        placeholder_frame = ttk.LabelFrame(main_frame, text="Application Settings", padding=10)
        placeholder_frame.pack(fill="both", expand=True, pady=10)
        
        ttk.Label(placeholder_frame, 
                 text="General application settings will be added in future updates.",
                 wraplength=350, justify="center").pack(pady=20)
        
        # Buttons frame
        buttons_frame = ttk.Frame(settings_dialog)
        buttons_frame.pack(fill="x", padx=15, pady=15)
        
        # Close button
        close_btn = ttk.Button(
            buttons_frame, 
            text="Close", 
            command=settings_dialog.destroy
        )
        close_btn.pack(side="right", padx=5)
        
        # Center the dialog on the main window
        settings_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() / 2) - (settings_dialog.winfo_width() / 2)
        y = self.root.winfo_y() + (self.root.winfo_height() / 2) - (settings_dialog.winfo_height() / 2)
        settings_dialog.geometry(f"+{int(x)}+{int(y)}")

    def _delayed_show_guidance(self):
        """Show guidance dialog after initial rendering"""
        # Check if guidance dialog is already showing
        if hasattr(self, 'guidance_dialog_showing') and self.guidance_dialog_showing:
            self.debug_print("Guidance dialog already showing, not creating another one")
            return
            
        # Set flag to indicate dialog is showing
        self.guidance_dialog_showing = True
        
        # Get main window position and size
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # Calculate centered position for the dialog relative to main window
        dialog_width = 600
        dialog_height = 575
        x = root_x + (root_width - dialog_width) // 2
        y = root_y + (root_height - dialog_height) // 2
        
        # Create the dialog window with position already set
        guidance_dialog = tk.Toplevel(self.root)
        guidance_dialog.title("Welcome to Rapid Moment Navigator")
        guidance_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        guidance_dialog.transient(self.root)
        guidance_dialog.grab_set()
        
        # Function to handle dialog close
        def on_dialog_close():
            self.guidance_dialog_showing = False
            guidance_dialog.destroy()
            
        # Set the close protocol
        guidance_dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        
        # Create main frame with padding
        main_frame = ttk.Frame(guidance_dialog, padding=15)
        main_frame.pack(fill="both", expand=True)
        
        # Title label
        title_label = ttk.Label(main_frame, text="No Media Found!", font=("TkDefaultFont", 14, "bold"))
        title_label.pack(anchor="center", pady=(0, 20))
        
        # Explanation text
        explanation_text = (
            "Rapid Moment Navigator needs to know where your media is located.\n\n"
            "You need to add directories that contain subtitle (.SRT) files and their corresponding video files.\n\n"
            "How it works:\n"
            "1. The app will scan directories for subtitle files (.SRT) and video files with matching names\n"
            "2. It can scan nested folder structures - media can be anywhere in the folder hierarchy\n"
            "3. Video files should have the same base name as their subtitle files\n"
            "4. You can add multiple media directories\n\n"
            "You have two options:"
        )
        
        explanation_label = ttk.Label(main_frame, text=explanation_text, wraplength=550, justify="left")
        explanation_label.pack(pady=10, fill="x")
        
        # Options frame
        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill="both", expand=True, pady=10)
        
        # Option 1: Current directory
        option1_frame = ttk.LabelFrame(options_frame, text="Option 1: Create show folders in current directory", padding=10)
        option1_frame.pack(fill="x", pady=5)
        
        # Current directory text
        current_dir_text = f"Current directory: {self.get_current_directory()}"
        current_dir_label = ttk.Label(option1_frame, text=current_dir_text, wraplength=550)
        current_dir_label.pack(anchor="w")
        
        # Instructions for Option 1
        option1_instructions = ttk.Label(option1_frame, 
                                       text="Create subfolders in this directory and add your media files + subtitles to them.",
                                       wraplength=550)
        option1_instructions.pack(anchor="w", pady=5)
        
        # Button frame for Option 1
        option1_btn_frame = ttk.Frame(option1_frame)
        option1_btn_frame.pack(anchor="w", pady=5)
        
        # Open current directory button
        def open_current_dir():
            try:
                current_dir = self.get_current_directory()
                if sys.platform.startswith('win'):
                    os.startfile(current_dir)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.Popen(['open', current_dir])
                else:  # Linux
                    subprocess.Popen(['xdg-open', current_dir])
            except Exception as e:
                self.debug_print(f"Error opening directory: {e}")
                self.status_var.set(f"Error opening directory: {e}")
        
        open_dir_btn = ttk.Button(option1_btn_frame, text="Open Current Directory", command=open_current_dir)
        open_dir_btn.pack(side="left", padx=5)
        
        # Refresh shows button
        def refresh_current_dir():
            self.debug_print("Refreshing shows from current directory")
            
            # Ensure current directory is included
            if self.preferences.get("exclude_current_dir", False):
                self.preferences["exclude_current_dir"] = False
                self.save_preferences()
                self.update_directory_listbox()
            
            # Clear existing show map and reload everything
            self.show_name_to_path_map.clear()
            
            # Reload shows and remap files
            shows_paths = self.load_shows()
            
            # Check if we found any shows
            if len(self.show_name_to_path_map) > 0:
                self.map_subtitles_to_videos()
                self.update_show_dropdown()
                self.status_var.set(f"Found {len(self.show_name_to_path_map)} shows with {len(self.subtitle_to_video_map)} mapped videos.")
                on_dialog_close()  # Close the dialog if shows were found
            else:
                self.status_var.set("No shows found in current directory. Please create show folders with subtitle files.")
        
        refresh_btn = ttk.Button(option1_btn_frame, text="Refresh Shows", command=refresh_current_dir)
        refresh_btn.pack(side="left", padx=5)
        
        # Option 2: Add existing directories
        option2_frame = ttk.LabelFrame(options_frame, text="Option 2: Add existing media directories", padding=10)
        option2_frame.pack(fill="x", pady=5)
        
        # Create a horizontal frame to hold both text and button side by side
        option2_content_frame = ttk.Frame(option2_frame)
        option2_content_frame.pack(fill="x", pady=5)
        
        # Instructions for Option 2 - now in the horizontal frame
        option2_instructions = ttk.Label(option2_content_frame, 
                                       text="Select directories that already contain media files and subtitles.",
                                       wraplength=400)  # Reduced width to make room for button
        option2_instructions.pack(side="left", anchor="w")
        
        # Function for Add Directory button
        def call_add_directory():
            on_dialog_close()  # Close the guidance dialog
            self.root.after(100, self.add_directory)
        
        # Create Add Directory button directly in the horizontal frame
        add_dir_btn = ttk.Button(option2_content_frame, text="Add Directory", command=call_add_directory)
        add_dir_btn.pack(side="right", padx=10)
        
        # Bottom buttons frame
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill="x", pady=15)
        
        # Help button
        def show_help():
            help_text = (
                "Detailed Help:\n\n"
                "Subtitle files (.SRT) should match the video file names.\n"
                "Example:\n"
                "- video_file.mp4 → video_file.srt\n\n"
                "The app will try to match files even with slight differences in naming.\n"
                "Each directory you add is treated as a separate 'show' in the dropdown menu.\n\n"
                "When you add a directory, the entire folder structure is scanned for matching\n"
                "subtitle and video files. Files can be in subdirectories.\n\n"
                "After adding directories, use the search box to find specific dialog,\n"
                "then click on the timecode to play the video at that exact moment."
            )
            messagebox.showinfo("How It Works", help_text)
        
        help_btn = ttk.Button(bottom_frame, text="Detailed Help", command=show_help)
        help_btn.pack(side="left", padx=5)
        
        # Close button
        close_btn = ttk.Button(bottom_frame, text="Close", command=on_dialog_close)
        close_btn.pack(side="right", padx=5)
        
        # Ensure the dialog is visible and focused
        guidance_dialog.lift()
        guidance_dialog.focus_force()
        
        # Log the position for debugging
        self.debug_print(f"Created guidance dialog at position ({x},{y}) with size {dialog_width}x{dialog_height}")
        
        # Force update the UI to ensure all elements are rendered
        guidance_dialog.update()

class DebugWindow:
    """A debug window to display errors and debug information"""
    def __init__(self, parent, auto_show=False):
        self.parent = parent
        
        # Get parent window position and size
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Set debug window dimensions
        debug_width = 800
        debug_height = 425  # Increased from 400 to 425 to show all buttons
        
        # Calculate position (centered relative to parent)
        x = parent_x + (parent_width - debug_width) // 2
        y = parent_y + (parent_height - debug_height) // 2
        
        # Create the window with position already set
        self.window = tk.Toplevel(parent)
        self.window.title("Debug Console")
        self.window.geometry(f"{debug_width}x{debug_height}+{x}+{y}")
        self.window.transient(parent)
        
        # Create a frame for the text area and scrollbars
        frame = ttk.Frame(self.window)
        frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create text area with scrollbars
        self.text_area = scrolledtext.ScrolledText(frame, wrap=tk.WORD, font=("Consolas", 10))
        self.text_area.pack(fill="both", expand=True)
        
        # Create a frame for buttons
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        # Add clear button
        clear_btn = ttk.Button(button_frame, text="Clear", command=self.clear_text)
        clear_btn.pack(side="left", padx=5)
        
        # Add close button
        close_btn = ttk.Button(button_frame, text="Close", command=self.window.withdraw)
        close_btn.pack(side="right", padx=5)
        
        # Add save button
        save_btn = ttk.Button(button_frame, text="Save Log", command=self.save_log)
        save_btn.pack(side="right", padx=5)
        
        # If auto_show is False, hide the window initially
        if not auto_show:
            self.window.withdraw()
            
        # Add help text
        self.insert_text("Debug Console - Errors and detailed messages will appear here.\n")
        self.insert_text("If you encounter problems, please save this log and include it in any bug reports.\n\n")
    
    def winfo_exists(self):
        """Check if window still exists"""
        try:
            return self.window.winfo_exists()
        except:
            return False
    
    def insert_text(self, text):
        """Insert text into the debug window"""
        try:
            self.text_area.insert(tk.END, text)
            self.text_area.see(tk.END)  # Auto-scroll to the end
            
            # Make the window visible if it's hidden
            if not self.window.winfo_viewable():
                self.window.deiconify()
        except:
            pass  # Silently fail if window is closed
    
    def clear_text(self):
        """Clear all text from the window"""
        try:
            self.text_area.delete(1.0, tk.END)
        except:
            pass
    
    def save_log(self):
        """Save the log contents to a file"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title="Save Debug Log"
            )
            if filename:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(self.text_area.get(1.0, tk.END))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save log: {e}")

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Rapid Moment Navigator - Search subtitles and navigate to moments in videos")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")
    args = parser.parse_args()
    
    if args.debug:
        print(f"DEBUG: Starting application in debug mode")
        print(f"DEBUG: Current directory: {os.path.abspath('.')}")
        print(f"DEBUG: Script directory: {os.path.dirname(os.path.abspath(__file__))}")
    
    # Create the main window
    root = tk.Tk()
    
    try:
        # Initialize and run the application
        app = RapidMomentNavigator(root, debug=args.debug)
        
        # Force debug output to be flushed immediately if debug is enabled
        if args.debug:
            sys.stdout.flush()
            
        # Create menu bar
        menu_bar = tk.Menu(root)
        
        # Add Settings menu with separate items
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Minimum Import Duration...", 
                                 command=app._show_settings_dialog)
        settings_menu.add_command(label="General Settings...", 
                                 command=app._show_general_settings_dialog)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
            
        # Add Debug menu
        debug_menu = tk.Menu(menu_bar, tearoff=0)
        debug_menu.add_command(label="Show Debug Window", 
                              command=lambda: app.ensure_debug_window() or 
                                             (app.debug_window and app.debug_window.window.deiconify()))
        menu_bar.add_cascade(label="Debug", menu=debug_menu)
        
        # Apply menu bar to root window
        root.config(menu=menu_bar)
        
        # Start the main loop
        root.mainloop()
        
    except Exception as e:
        # Handle any uncaught exceptions during initialization
        print(f"CRITICAL ERROR: {e}", file=sys.stderr)
        traceback.print_exc()
        messagebox.showerror("Critical Error", 
                           f"The application encountered a critical error and cannot start:\n\n{str(e)}")
        sys.exit(1) 