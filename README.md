# copy64
Powershell script to copy files to/from paths longer than 260 characters

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
