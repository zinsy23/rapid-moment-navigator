#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Defines the hotkey Ctrl+Alt+R to launch the Rapid Moment Navigator
^!r::
    Run, pythonw.exe "%A_ScriptDir%\rapid_moment_navigator.py"
    return

; Optionally add another hotkey to focus the app if it's already running
^!f::
    if WinExist("Rapid Moment Navigator")
        WinActivate
    else
        Run, pythonw.exe "%A_ScriptDir%\rapid_moment_navigator.py"
    return 