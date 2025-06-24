# Rapid Moment Navigator

A Python application for searching subtitle files and quickly navigating to specific moments in your media files. See my [video](https://www.youtube.com/watch?v=jP1fWYjkZ5s) where I try to make this in an hour with an AI coding agent. I also did a [follow-up video](https://www.youtube.com/watch?v=s2W8Wnd-ARs) where I show some interesting lessons I'm learning as I develop it further after the original challenge.

## Features

- Search for keywords across subtitle files (.srt)
- Select different shows from a dropdown menu
- View search results with clickable timecodes
- Click on a result to open the corresponding video at the exact timestamp
- Automatic mapping between subtitle files and video files
- Support for various video formats (MP4, MKV, AVI, MOV, WMV, etc.)
- Generic directory structure support - works with any organization pattern
- Ctrl+Backspace support in the search field for faster text editing
- Multi-directory support with preferences system
- Add and remove media directories through a user-friendly interface
- 🎬 Multi-Editor Support: Extensible editor registry system for video editing software
- 📝 DaVinci Resolve Integration: Import media and clips directly to DaVinci Resolve timelines
- 🔧 Plugin Architecture: Easy-to-extend system for adding new video editors

## Requirements

- Python 3.6 or above
- Media Player Classic (MPC-HC) installed (preferably at one of the standard locations)
- Tkinter (usually comes with Python installation)

## Installation

1. Make sure you have Python installed. If not, download and install from [python.org](https://python.org)
2. No additional packages are needed beyond the standard library

## Usage

1. Place the `rapid_moment_navigator.py` script in any directory
2. Run the script with Python:
   ```
   python rapid_moment_navigator.py
   ```
3. By default, the application will scan the current directory for show folders
4. To add additional media directories, click the "Add Directory" button
5. Select a show from the dropdown menu
6. Enter a search term
7. Click "Find All" or press Enter to search
8. Results will display with clickable timecodes
9. Click on any timecode to open the corresponding video at that timestamp

### Media Directory Management

The application now supports managing multiple media directories:

1. **Adding Directories**: Click the "Add Directory" button to open the directory manager
   - You can add multiple directories that contain your show folders
   - Existing directories are displayed in the dialog for easy management
   - Use the "Add Multiple Directories" button to select multiple directories in sequence
   - Each directory added will be scanned for SRT files in any subfolder
   - Preferences are automatically saved to `rapid_navigator_prefs.json`

2. **Removing Directories**: Select a directory in the list and click "Remove Directory"
   - You can remove the current directory, but only if you've added other directories
   - If you remove all directories, the current directory will be automatically re-added
   - The application ensures there is always at least one directory to search

3. **Directory Structure**: The application is flexible and works with various folder structures
   - SRT files can be located anywhere within the directory tree
   - Video files can be anywhere within the directory tree
   - The application will automatically find and map subtitle files to video files

### Command-line Options

- `--debug`: Enable debug output to console for troubleshooting
   ```
   python rapid_moment_navigator.py --debug
   ```

## File Structure

The application is flexible and can handle various directory structures:

1. Subtitle files should be in a folder named "Subtitles" within each show's directory
2. Video files can be anywhere within the show's directory structure
3. The application will automatically find and map subtitle files to video files based on filename similarity
4. Supports common video formats including MP4, MKV, AVI, MOV, WMV, and more

Example structures that are supported:

```
Parent Directory
│
├── Show1
│   ├── Subtitles
│   │   ├── Show1 - 1x01.srt
│   │   └── ...
│   │
│   ├── Season 1
│   │   ├── Show1 - 1x01.mp4
│   │   └── ...
│   │
│   └── Season 2
│       └── ...
│
├── Show2
│   ├── Subtitles
│   │   ├── SHOW2_S01E01.srt
│   │   └── ...
│   │
│   ├── S01
│   │   ├── SHOW2_S01E01.mkv
│   │   └── ...
│   │
│   └── S02
│       └── ...
│
└── Show3
    ├── Subtitles
    │   ├── S01_DISC1_Title1.srt
    │   └── ...
    │
    └── S01
        ├── D1
        │   ├── S01_DISC1_Title1.mov
        │   └── ...
        └── D2
            └── ...
```

## Preferences System

The application stores your preferences in a JSON file named `rapid_navigator_prefs.json` located in the same directory as the script. The preferences include:

- Media directories: Paths to directories containing show folders

If the preferences file doesn't exist or no directories are defined, the application will use the current directory by default.

## Filename Matching Logic

The application attempts to match subtitle and video files using the following strategies:

1. Exact filename match (ignoring extensions)
2. Partial matching with cleaned filenames (removing separators, common words, etc.)

This approach handles various naming conventions and formats automatically.

## Running Globally

To run the application globally using an AutoHotkey mapping, follow these steps:

1. Install AutoHotkey from [https://www.autohotkey.com/](https://www.autohotkey.com/)
2. If you want to use the included `rapid_moment_navigator.ahk` script directly, you can add an AutoHotkey shortcut mapping to it using standard syntax. I personally have a master AutoHotkey script for all my shortcuts, so to make it run globally across the system, I've created a `.lnk` shortcut file to the script and run that from the master script using the following syntax in the same directory as my master script:

```
#+s::
run, launch-rapid-moment-navigator.ahk.lnk,,hide
return
```

**Note:** The reason I'm doing it this way with my setup is because I want to keep this script in my GitHub repository on my system, but access it from my AutoHotkey scripts on my system. It works great if I have the `launch-rapid-moment-navigator.ahk` file in the repository, and then have a Windows `.lnk` shortcut to that in my AutoHotkey scripts directory where I have all my personal scripts, including the one that runs this using the above syntax. The `,,hide` prevents a console window from appearing when the script is run.

## Editor Registry System

The application features a powerful Editor Registry System that makes it easy to add support for new video editing software. The system provides a clean, modular architecture for integrating different editors without hardcoding editor-specific logic throughout the codebase.

### Supported Editors

Currently supported editors:
- 🎥 DaVinci Resolve: Full integration with media/clip import, framerate detection, and timeline management

### Adding New Editors

The editor registry system makes adding new editors straightforward. Here's how to extend the application:

#### 1. 📋 Registry Configuration

Add your editor to the `EDITOR_REGISTRY` dictionary in the `RapidMomentNavigator` class:

```python
EDITOR_REGISTRY = {
    "DaVinci Resolve": {
        "ensure_ready_method": "_ensure_resolve_ready",
        "import_media_method": "_import_media_to_davinci_resolve", 
        "import_clip_method": "_import_clip_to_davinci_resolve",
        "framerate_detection_method": "detect_video_framerate_from_resolve",
        "supports_advanced_framerate": True,
        "timecode_format": "no_milliseconds"
    },
    # Add your new editor here:
    "Adobe Premiere Pro": {
        "ensure_ready_method": "_ensure_premiere_ready",
        "import_media_method": "_import_media_to_premiere",
        "import_clip_method": "_import_clip_to_premiere", 
        "framerate_detection_method": "detect_video_framerate_from_premiere",
        "supports_advanced_framerate": True,
        "timecode_format": "with_milliseconds"
    }
}
```

#### 2. 🛠️ Required Methods

Implement the following methods for your editor:

##### **Readiness Check Method**
```python
def _ensure_premiere_ready(self):
    """
    Ensure Adobe Premiere Pro API is ready for use.
    
    Returns:
        bool: True if ready, False if failed or in safe mode
    """
    # Check if already initialized
    if hasattr(self, 'premiere_initialized') and self.premiere_initialized:
        return True
        
    # Check if in safe mode (previous failures)
    if hasattr(self, 'premiere_in_safe_mode') and self.premiere_in_safe_mode:
        return False
    
    try:
        # Initialize Premiere Pro API
        # Your initialization code here
        self.premiere_initialized = True
        return True
    except Exception as e:
        self.premiere_in_safe_mode = True
        self.debug_print(f"❌ Premiere Pro initialization failed: {e}")
        return False
```

##### **Media Import Method**
```python
def _import_media_to_premiere(self, subtitle_file, start_time, end_time):
    """
    Import media file to Adobe Premiere Pro.
    
    Args:
        subtitle_file (str): Path to subtitle file (used to find corresponding video)
        start_time (str): Start timecode (optional, for reference)
        end_time (str): End timecode (optional, for reference)
    """
    # Get corresponding video file
    video_file = self.subtitle_to_video_map.get(subtitle_file)
    if not video_file:
        self.debug_print(f"❌ No video file found for subtitle: {subtitle_file}")
        return
        
    try:
        # Your Premiere Pro import logic here
        self.debug_print(f"📥 Importing to Premiere Pro: {video_file}")
        # Example: premiere_api.import_media(video_file)
        
    except Exception as e:
        self.debug_print(f"❌ Premiere Pro import failed: {e}")
```

##### **Clip Import Method**
```python
def _import_clip_to_premiere(self, subtitle_file, start_time, end_time):
    """
    Import specific clip segment to Adobe Premiere Pro timeline.
    
    Args:
        subtitle_file (str): Path to subtitle file
        start_time (str): Start timecode for clip
        end_time (str): End timecode for clip
    """
    video_file = self.subtitle_to_video_map.get(subtitle_file)
    if not video_file:
        return
        
    try:
        # Your clip import logic here
        self.debug_print(f"✂️ Importing clip to Premiere Pro: {start_time} - {end_time}")
        # Example: premiere_api.import_clip(video_file, start_time, end_time)
        
    except Exception as e:
        self.debug_print(f"❌ Premiere Pro clip import failed: {e}")
```

##### **Framerate Detection Method** (Optional)
```python
def detect_video_framerate_from_premiere(self, video_path):
    """
    Detect video framerate using Adobe Premiere Pro API.
    
    Args:
        video_path (str): Path to video file
        
    Returns:
        float: Detected framerate, or None if detection fails
    """
    try:
        # Your Premiere Pro framerate detection logic
        # Example: return premiere_api.get_framerate(video_path)
        return 24.0  # Fallback
    except Exception as e:
        self.debug_print(f"❌ Premiere Pro framerate detection failed: {e}")
        return None
```

#### 3. 🎯 Registry Configuration Options

Configure your editor's capabilities in the registry:

| Option | Description | Values |
|--------|-------------|--------|
| `ensure_ready_method` | Method name for initialization check | `"_ensure_your_editor_ready"` |
| `import_media_method` | Method name for media import | `"_import_media_to_your_editor"` |
| `import_clip_method` | Method name for clip import | `"_import_clip_to_your_editor"` |
| `framerate_detection_method` | Method name for framerate detection | `"detect_video_framerate_from_your_editor"` |
| `supports_advanced_framerate` | Whether editor supports advanced framerate detection | `true` / `false` |
| `timecode_format` | Timecode format preference | `"with_milliseconds"` / `"no_milliseconds"` |

#### 4. 🚀 Automatic Integration

Once you've added the registry entry and implemented the required methods:

1. **Import buttons automatically appear** when your editor is selected
2. **Generic dispatching handles routing** to your editor-specific methods  
3. **Fallback systems work automatically** for framerate detection and error handling
4. **No changes needed** to existing UI or import handler code

#### 5. 🧪 Testing Your Integration

Test your new editor integration:

```python
# The application will automatically:
# 1. Show your editor in the dropdown
# 2. Display import buttons when selected
# 3. Route import clicks to your methods
# 4. Handle errors gracefully with fallbacks
```

### Architecture Benefits

The editor registry system provides several key advantages:

- 🔌 Zero Hardcoding: No editor names hardcoded in import handlers
- 🎯 Single Source of Truth: All editor capabilities defined in one place
- ⚡ Easy Extension: Add new editors with minimal code changes
- 🛡️ Robust Error Handling: Built-in fallbacks and safety checks
- 🔄 Consistent Interface: Same import workflow for all editors
- 📈 Future-Proof: Easy to add editor-specific features and capabilities

### Advanced Features

The registry system supports advanced editor-specific features:

- **Custom Timecode Formats**: Different editors prefer different timecode formats
- **Capability Flags**: Mark which features each editor supports
- **Fallback Chains**: Automatic fallback to generic methods when editor-specific ones fail
- **State Management**: Smart initialization that only runs when needed
- **Performance Optimization**: Lazy loading and caching of editor connections

## Notes

- If Media Player Classic (MPC-HC) isn't found at the default locations, the application will try alternative locations or fall back to the default media player
- The search is case-insensitive
- HTML tags in subtitles are ignored during both search and display
- The application supports a wide range of video formats - if MPC-HC can play it, the application can usually handle it 

## Future Improvements

- Verify support for other video players and formats in Windows, Mac, and Linux
   - Provide a moment offset mechanism since keywords likely take the user to the middle of a moment and the user generally rewinds to go to the beginning of the moment
   - If possible, add support for marking a Plex video at a timecode to view a moment on another device connected to the Plex server
   - Allow users to add a custom media player based on criteria, like a path or a shell command, or utilizing an API to connect to the media player
- Add a smarter search mechanism so keyword matches don't have to be exact, like a hyphen vs space, or a space vs no space, etc.
- Better duplicate file handling, particularly where corresponding file matches are inaccurate.
- Integrate editor subtitle generation to index videos for searching around various directory structures and automatically figure out what's not indexed.
   - Or if something is poorly/not properly indexed, can be marked in the interface to re-index that automatically.
- Add support to control the GUI via command line arguments, which would be useful for automation and even using Python-based voice assistants to control the application, like Neon AI or OVOS.
- Make it easier to import specific moments in editors not from keywords when a file and rough timecode is known.
- Add support for importing media player favorites/bookmarks into editors
- Add bookmark mechanism for use with this software
- Allow the user to modify the timecode range of search results
   - Improve the scroll bar behavior when new results are rendered
- Add the ability to open specific moments remotely on other computers (via SSH or network sockets on multiple computers, depending on their capabilities)
- Add AI capabilities for finding moments in shows, assuming efficient tokenization can be done with good support of models
   - The AI could also mostly automate creating compilations based on criteria (i. e. moments in Silicon Valley where Gavin Belson is demonstrating sociopathic tendencies or it's implied by other people mentioning it or moments where Leon makes Seinfeld references in Mr. Robot)
- Improve support for older versions of DaVinci Resolve utilizing older python scripts.
