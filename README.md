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
