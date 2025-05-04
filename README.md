# Rapid Moment Navigator

A Python application for searching subtitle files and quickly navigating to specific moments in your media files.

## Features

- Search for keywords across subtitle files (.srt and .txt)
- Select different shows from a dropdown menu
- View search results with timecodes
- Click on a result to open the corresponding video at the exact timestamp
- Automatic mapping between subtitle files and video files

## Requirements

- Python 3.6 or above
- Media Player Classic (MPC-HC) installed at the default location (`C:\Program Files\MPC-HC\mpc-hc64.exe`)
- Tkinter (usually comes with Python installation)

## Installation

1. Make sure you have Python installed. If not, download and install from [python.org](https://python.org)
2. No additional packages are needed beyond the standard library

## Usage

1. Place the `rapid_moment_navigator.py` script in the parent directory of your show folders
2. Run the script with Python:
   ```
   python rapid_moment_navigator.py
   ```
3. Select a show from the dropdown menu
4. Enter a search term
5. Click "Find All" or press Enter to search
6. Results will display with timecodes
7. Click on any timecode to open the corresponding video at that timestamp

## File Structure Requirements

The application expects your media to be organized as follows:

```
Parent Directory (where the script is located)
│
├── Show1
│   ├── Subtitles
│   │   ├── Show1 - 1x01 - Episode_Name.mp4.srt
│   │   └── ...
│   │
│   ├── Season 1
│   │   ├── Show1 - 1x01 - Episode_Name.mp4
│   │   └── ...
│   │
│   └── Season 2
│       └── ...
│
└── Show2
    └── ...
```

The application will automatically map subtitle files to video files based on their filenames.

## Notes

- If Media Player Classic (MPC-HC) isn't found at the default location, the application will fall back to using the default media player
- The search is case-insensitive
- HTML tags in subtitles are ignored during both search and display 