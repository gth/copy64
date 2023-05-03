# copy64
Powershell script to copy files to/from paths longer than 260 characters

Draft objectives (to get this repository into a workable state)
- central storage of various examples and works-in-progress
- collate main code into a single working source file
- establish version control

General objectives:
- use Powershell functions and, where possible, operate on as many Powershell versions as possible
- overcome the Windows bugs around a directory path longer than 260 characters (e.g. the kind where Windows Explorer drag-and-drop copying fails)
- a graphical interface, for simplicity and nice overall clarity
- supports multiple sources
- a trace log on screen, for detailed status and troubleshooting
- a log file, to keep a record of what happened (do you remember if you already copied that really large file, successfully?)
- currently using the built-in Copy-Item, which handles long paths (good) but locks up the GUI until it has finished (bad), 
- handles local or network drives

Next release goals:
- multiple destination(s)
- launch the copy task(s) as a separate process, to keep the GUI responsive
- stream-based copying, to provide:
-- progress indication (progress gauge)
-- time elapsed
-- estimated time remaining
- check prior to running copy task, that

Future goals:
- save current source(s) and destination(s) values
- command-line options, for possible automation
- retry functionality (e.g. if network links drop and manual intervention is needed to fix them)
- simple verification (e.g. based on file attributes)
- hash-based verification (e.g. MD5, SHA1, SHA256)
- move files, not just copy
- functionality to resume previous execution (e.g. Windows crashed, or system power loss)
- multi-threaded, if possible

Reference Links
Some Powershell, .NET and related github links:
- https://github.com/PowerShell/PowerShell/issues/2581 - Improve Copy-Item to provide parity with Robocopy (has a good list with similar goals to those listed above!)
- https://github.com/dotnet/runtime/issues/60903 - .NET API Proposal: DirectoryInfo.Copy / Directory.Copy
- https://github.com/dotnet/runtime/issues/20695 - .NET Provide overloads of File.Copy that support cancellation and progress
- https://github.com/chinhdo/txFileManager/issues/30 - Transactional File Manager .NET library: Directory/File Copy/Move Cancel Support

Code Sample Links
- https://stackoverflow.com/questions/2339313/how-to-use-correctly-copyfileex-and-copyprogressroutine-functions - large-number math
- https://stackoverflow.com/questions/51162410/powershell-add-text-to-progressbar-gui - just a label overlay
- Using color for row status:
-- https://social.technet.microsoft.com/Forums/en-US/05144c37-4991-46fa-b2d8-dcb0a7f266df/wpf-powershell-listbox-change-foreground-color-of-item-if-it-matches-expression?forum=winserverpowershell
-- https://social.technet.microsoft.com/Forums/ie/en-US/7f7d8c8c-729e-40b8-89ed-624e251fce4f/textbox-listbox-color-separate-rows?forum=winserverpowershell
- Waiting/Progress updates
-- https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/wait-event?view=powershell-7.3
-- https://social.technet.microsoft.com/Forums/en-US/2541cf38-798d-4381-bf35-edae8c11ccba/forms-in-powershell-use-systemwindowsformsprogressbar-to-display-progress-in-realtime?forum=winserverpowershell
-- https://stackoverflow.com/questions/2434133/progress-during-large-file-copy-copy-item-write-progress
