
##################################################################################
###  Copy64.ps1                                  https://github.com/gth/copy64 ###
###  ------------------------------------------------------------------------- ###
###  Simple Windows PowerShell GUI that takes a source and destination and     ###
###  copies from one to the other.  A trace log window helps view progress.    ###
###  Modern copy commands such as PowerShell's Copy-Item and/or Windows APIs   ### 
###  CopyFileEx() and CopyFile2() are used and, as a result, various issues    ###
###  are avoided including:                                                    ###
###  - long-filename limitations (260 characters)                              ###
###  - maximum copy count of ~1,500 files (BITS-based copying)                 ###
###  - maximum filesize issues (2GB/4GB etc), with various causes              ###
###  - 32bit vs 64bit issues                                                   ###
###                                                                            ###
###  -- BETA RELEASE STATUS --                                                 ###
###  Always (at least) compare folder properties to ensure files have been     ###
###  copied successfully. Until the script can automatically verify the copied ###
###  data, a successful copy should NOT be assumed.                            ###
###  !! USE AT OWN RISK! YOU HAVE BEEN WARNED! !!                              ###
###                                                                            ###
###  -- ACKNOWLEDGEMENT --                                                     ###
###  Credit to @FranciscoNabas for making CopyFileEx useable in PowerShell:    ###
###  github.com/FranciscoNabas/PowerShellPublic/blob/main/Copy-File.ps1        ###
###  stackoverflow.com/questions/2434133/progress-during-large-file-copy-copy-item-write-progress/76049235#76049235
###                                                                            ###
###  ------------------------------------------------------------------------- ###
###  Usage after downloading the file to your PC:                              ###
###   1. Right-click the script and select "Run with Powershell"               ###
###   2. Drag-and-drop file(s) or folder(s) to the source box at the top of    ###
###      the form.                                                             ###
###   3. Drag-and-drop a folder to the destination box (cannot copy to a file) ###
###   4. Click the COPY button                                                 ###
###   5. Manually compare folder properties to confirm copy was successful.    ###
###                                                                            ###
###  ------------------------------------------------------------------------- ###
###  Release log:                                                              ###
###  v1.0 Initial version                                                      ###
###  v1.1 Minor improvements                                                   ###
###  v1.2 Wider form; DPI aware rendering; debug messages; timestamps; copy    ###
###       durations.                                                           ###
###  v1.3 Log file now exists and saved as:                                    ###
###       (script_path)\(date-time)_(hostname)_(process_id).log                ###
###  v1.4 Relaunch if in a 32bit Powershell; Checks longpathsenabled registry  ###
###       setting; Many-to-many copying re-enabled; Stream-based file copying  ###
###       with progress added; check for admin rights.                         ###
###                                                                            ###
###  ------------------------------------------------------------------------- ###
###  KNOWN BUGS:                                                               ###
###  - No error checking before copying: source is presumed to be accessible,  ###
###    destination folder is presumed to be writable and source files are      ###
###    presumed not to already exist in the destination folder.                ###
###  - Although PowerShell's Copy-Item command supports long filenames,        ###
###    Windows Explorer drag-and-drop function does NOT.  Work-around: Ensure  ###
###    any dragged and dropped files have a paths less than 260 characters.    ###
###  - No detection of errors during copying.                                  ###
###  - No detection/verification copy was successful.                          ###
###  - (New, due to the addition of streaming copy providing progress updates) ###
###     --- progress callback doesn't take place during a *directory* copy.    ###
###         Added new roadmap goal regarding directory sources.                ###
###  - If PowerShell is running as an Administrator, windows security settings ###
###    block drag-and-drop - i.e. users cannot drag-and-drop from a user-level ###
###    program like Windows Explorer to an administrator-level program, which  ###
###    COPY64 would become if we requested elevated rights. For now,           ###
###    Administrator rights are NOT requested (see entry below).               ###
###  - Administrator elevation IS required to copy to restricted folders such  ###
###    as C:\ or C:\WINDOWS and so on. For now, such destinations will fail.   ###
###                                                                            ###
###  ------------------------------------------------------------------------- ###
###  RELEASE ROADMAP - 1.x:                                                    ###
###  - detect & display error messages after a failed copy.                    ###
###  - Abort due to any incurable error conditions, before command runs.       ###
###  - indicate copy speed (MB/sec).                                           ###
###  - add support for basic verification:                                     ###
###    -- filesize                                                             ###
###    -- first 300 bytes                                                      ###
###    -- last 300 bytes                                                       ###
###  - check for write permission in target folders.                           ###
###  - Recursively copy directories - i.e. this script copies files and        ###
###    continues to do so should it find any directories, by traversing down   ###
###    into them. Thorough testing of long folders, very large file-counts,    ###
###    long filenames and files well over 4GB in size is required.             ###
###                                                                            ###
###  RELEASE ROADMAP - 2.x:                                                    ###
###  - check for and prevent/correct as many error conditions as possible,     ###
###    before copy command runs.                                               ###
###  - add support for hash-based verification:                                ###
###    -- MD5                                                                  ###
###    -- SHA1                                                                 ###
###    -- SHA256                                                               ###
###  - Add "simulate only" button to see if long filename issue is present in  ###
###    any of the source items (don't actually do the copy).                   ###
###                                                                            ###
###  RELEASE ROADMAP - 3.x:                                                    ###
###  - detect when when script crashes + offer to redo it from scratch         ###
###    -- will need to be able to reload prior source & destination lists.     ###
###  - add resume support (copy from where we left off)                        ###
###    -- will need to detect incomplete copy (see CopyFile2 parameter)        ###
###    -- needs opt-in checkbox, due to recoverable copying performance impact ###
###  - graph showing performance                                               ###
###                                                                            ###
##################################################################################


### Global variables ###

$global:debug = $False           # Enable to see very detailed program progress messages (there are many!)
$global:hide_console = $True     # Disable to ensure any diagnostics / errors / debug messages are visible.
$global:source_total = 0
$global:dest_total = 0
$global:copy_count = 0
$global:last_log_entry
[uint64]$global:total = 0
$my_logdir = Get-ScriptPath
$my_pid = $PID
$my_datestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss_")
$global:my_logfile = $my_logdir + "\" + $my_datestamp + $env:COMPUTERNAME + "-" + $my_pid + ".log"
write-host ("[INFO] Logging to " + $my_logfile ) -Foregroundcolor cyan

debugMsg "Global variables defined."

### Functions ###

### ------------ MAIN COPY FUNCTION ------------ ###
function Copy-File {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$Path,

        [Parameter(Mandatory, Position = 1)]
        [string]$Destination
    )

    $signature = @'
    namespace Utilities {

        using System;
        using System.Runtime.InteropServices;
    
        public class FileSystem {
            
            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool CopyFileEx(
                string lpExistingFileName,
                string lpNewFileName,
                CopyProgressRoutine lpProgressRoutine,
                IntPtr lpData,
                ref Int32 pbCancel,
                CopyFileFlags dwCopyFlags
            );
        
            delegate CopyProgressResult CopyProgressRoutine(
            long TotalFileSize,
            long TotalBytesTransferred,
            long StreamSize,
            long StreamBytesTransferred,
            uint dwStreamNumber,
            CopyProgressCallbackReason dwCallbackReason,
            IntPtr hSourceFile,
            IntPtr hDestinationFile,
            IntPtr lpData);
        
            int pbCancel;
        
            public enum CopyProgressResult : uint
            {
                PROGRESS_CONTINUE = 0,
                PROGRESS_CANCEL = 1,
                PROGRESS_STOP = 2,
                PROGRESS_QUIET = 3
            }
        
            public enum CopyProgressCallbackReason : uint
            {
                CALLBACK_CHUNK_FINISHED = 0x00000000,
                CALLBACK_STREAM_SWITCH = 0x00000001
            }
        
            [Flags]
            enum CopyFileFlags : uint
            {
                COPY_FILE_FAIL_IF_EXISTS = 0x00000001,
                COPY_FILE_RESTARTABLE = 0x00000002,
                COPY_FILE_OPEN_SOURCE_FOR_WRITE = 0x00000004,
                COPY_FILE_ALLOW_DECRYPTED_DESTINATION = 0x00000008
            }
        
            public void CopyWithProgress(string oldFile, string newFile, Func<long, long, long, long, uint, CopyProgressCallbackReason, System.IntPtr, System.IntPtr, System.IntPtr, CopyProgressResult> callback)
            {
                CopyFileEx(oldFile, newFile, new CopyProgressRoutine(callback), IntPtr.Zero, ref pbCancel, CopyFileFlags.COPY_FILE_RESTARTABLE);
            }
        }
    }
'@

    Add-Type -TypeDefinition $signature
    [Func      [long,        long,        long,             long,   System.UInt32, Utilities.FileSystem+CopyProgressCallbackReason, System.IntPtr,     System.IntPtr, System.IntPtr, Utilities.FileSystem+CopyProgressResult]]$copyProgressDelegate = {
        param($global:total, $transferred, $streamSize, $streamByteTrans, $dwStreamNumber,                                         $reason,  $hSourceFile, $hDestinationFile, $lpData)
        
        # GUI Percent progress bar
        $myunit = "bytes"
        $mytotal = $global:total
        $mytransferred = $transferred
        if ($mytotal -gt 1024) {
            $myunit = "KB"
            $mytransferred = $([Math]::Round(($mytransferred/1KB), 0))
            $mytotal = $([Math]::Round(($mytotal/1KB), 0))
            if ($mytotal -gt 1024) {
                $myunit = "MB"
                $mytransferred = $([Math]::Round(($mytransferred/1KB), 0))
                $mytotal = $([Math]::Round(($mytotal/1KB), 0))
                if ($mytotal -gt 2048) {
                    $myunit = "GB"
                    $mytransferred = $([Math]::Round(($mytransferred/1KB), 0))
                    $mytotal = $([Math]::Round(($mytotal/1KB), 0))
                    if ($mytotal -gt 2048) {
                        $myunit = "TB"
                        $mytransferred = $([Math]::Round(($mytransferred/1KB), 0))
                        $mytotal = $([Math]::Round(($mytotal/1KB), 0))
                    }
                }
            }
        }
        $my_ProgressBar_Bytes.Maximum = $mytotal
        $my_ProgressBar_Bytes.Value = $mytransferred
        $objLabelProgressNote.Text = $mytransferred.ToString() + $myunit + " of " + $mytotal + $myunit

        # Help Windows update the progress bar, and avoid an unrepsonsive GUI -
        [System.Windows.Forms.Application]::DoEvents()
    }

    $fileName = [System.IO.Path]::GetFileName($Path)
    $destFileName = [System.IO.Path]::GetFileName($Destination)
    if ([string]::IsNullOrEmpty($destFileName) -or $destFileName -notlike '*.*') {
        if ($Destination.EndsWith('\')) {
            $destFullName = "$Destination$fileName"
        }
        else {
            $destFullName = "$Destination\$fileName"
        }
    }
    $wrapper = New-Object Utilities.FileSystem
    $wrapper.CopyWithProgress($Path, $destFullName, $copyProgressDelegate)
}
### ------------ END OF MAIN COPY FUNCTION ------------ ###

function Get-ScriptPath() {
    # If using PowerShell ISE
    if ($psISE) {
        $ScriptPath = Split-Path -Parent -Path $psISE.CurrentFile.FullPath
    }
    # If using PowerShell 3.0 or greater
    elseif($PSVersionTable.PSVersion.Major -gt 3) {
        $ScriptPath = $PSScriptRoot
    }

    # If using PowerShell 2.0 or lower
    else {
        $ScriptPath = split-path -parent $MyInvocation.MyCommand.Path
    }

    # If still not found (e.g. running an exe created using PS2EXE module)
    if(-not $ScriptPath) {
        $ScriptPath = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
    }

    # Return result
    return $ScriptPath
}

function Test-RegistryValue ($regkey, $name) {
    if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
        $true
    } else {
        $false
    }
}

function LogfileMsg {
    param(
        [string]$Message
    )
    
    if($checkboxLOG.Checked -eq $True) {
        $timestamp = (Get-Date).ToString("dd-MM-yy HH:mm:ss")
        Add-Content -Path $global:my_logfile -Value ($timestamp + "  " + $Message)
    }
}

function debugMsg {
    param(
        [string]$Message
    )
    if ($global:debug) { 
        write-host "$Message" -ForegroundColor Yellow 
        $timestamp = (Get-Date).ToString("dd-MM-yy HH:mm:ss")
        LogfileMsg ("[DBG] * " + $Message)
    }
}

function logMsg {
    param(
        [string]$Message
    )
    $timestamp = (Get-Date).ToString("dd-MM-yy HH:mm:ss")
    $dummy = $listBoxTRACE.Items.Add("$timestamp  $Message")
    LogfileMsg ("[INF] " + $Message)
}

function format_duration {
    param(
        [timespan]$Duration
    )

    $Day = switch ($Duration.Days) {
        0 { $null; break }
        1 { "{0} day, " -f $Duration.Days; break }
        Default {"{0} days, " -f $Duration.Days}
    }
    $Hour = switch ($Duration.Hours) {
        0 { $null; break }
        1 { "{0} hour, " -f $Duration.Hours; break }
        Default { "{0} hours, " -f $Duration.Hours }
    }
    $Minute = switch ($Duration.Minutes) {
        0 { $null; break }
        1 { "{0} minute, " -f $Duration.Minutes; break }
        Default { "{0} minutes, " -f $Duration.Minutes }
    }
    $Second = switch ($Duration.Seconds) {
        #0 { $null; break }
        1 { "{0} second" -f $Duration.Seconds; break }
        Default { "{0} seconds" -f $Duration.Seconds }
    }
    write-host ("[INFO] Copy took " + $Day + $Hour + $Minute + $Second + ".") -Foregroundcolor cyan
    return ($Day + $Hour + $Minute + $Second)
}

function Show-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    # ShowNormalNoActivate = 4
    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    # Hide = 0
    if ($global:hide_console -eq $True) {
        [Console.Window]::ShowWindow($consolePtr, 0)
    }
}


### .NET methods ###


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Method to hide console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

# Method for "DPI Aware" form + enable "Visual Styles"
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@
$null = [ProcessDPI]::SetProcessDPIAware()
[System.Windows.Forms.Application]::EnableVisualStyles()


### GUI elements ###


$form = New-Object System.Windows.Forms.Form
$form.SuspendLayout()
$form.AutoScaleDimensions =  New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode  = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.ClientSize = New-Object System.Drawing.Size(960,1200)
$form.Text = "Copy64 (support paths over 260 characters)"
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size
$form.MaximizeBox = $False
$form.MinimizeBox = $False
$form.Topmost = $True
$form.FormBorderStyle = 'Fixed3D'

debugMsg "Main form element loaded."

$labelSOURCE = New-Object Windows.Forms.Label
$labelSOURCE.Location = New-Object System.Drawing.Point(5,10)
$labelSOURCE.Size = New-Object System.Drawing.Size(175,20)
$labelSOURCE.AutoSize = $False
$labelSOURCE.Text = "Source(s):"

$listBoxSOURCE = New-Object Windows.Forms.ListBox
$listBoxSOURCE.Location = New-Object System.Drawing.Point(50,35)
$listBoxSOURCE.Size = New-Object System.Drawing.Size(850,140)
$listBoxSOURCE.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$listBoxSOURCE.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$listBoxSOURCE.HorizontalScrollbar = $True
$listBoxSOURCE.AllowDrop = $True

$labelDEST = New-Object Windows.Forms.Label
$labelDEST.Location = New-Object System.Drawing.Point(5,200)
$labelDEST.Size = New-Object System.Drawing.Size(175,20)
$labelDEST.Text = "Destination(s):"

$listBoxDEST = New-Object Windows.Forms.ListBox
$listBoxDEST.Location = New-Object System.Drawing.Point(50,225)
$listBoxDEST.Size = New-Object System.Drawing.Size(850,140)
$listBoxDEST.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$listBoxDEST.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$listBoxDEST.HorizontalScrollbar = $True
$listBoxDEST.AllowDrop = $True

debugMsg "Src + Dest elements loaded."

$buttonCOPY = New-Object System.Windows.Forms.Button
$buttonCOPY.Location = New-Object System.Drawing.Point(50,380)
$buttonCOPY.Size = New-Object System.Drawing.Size(120,50)
$buttonCOPY.Text = "&Copy"

$buttonRESET = New-Object System.Windows.Forms.Button
$buttonRESET.Location = New-Object System.Drawing.Point(630,380)
$buttonRESET.Size = New-Object System.Drawing.Size(120,50)
$buttonRESET.Text = "&Reset"

$buttonCLOSE = New-Object System.Windows.Forms.Button
$buttonCLOSE.Location = New-Object System.Drawing.Point(780,380)
$buttonCLOSE.Size = New-Object System.Drawing.Size(120,50)
$buttonCLOSE.Text = "Clo&se"
 
debugMsg "Button elements loaded."

$checkboxCLEAR = New-Object Windows.Forms.Checkbox
$checkboxCLEAR.Location = New-Object System.Drawing.Point(50,435)
$checkboxCLEAR.Size = New-Object System.Drawing.Size(275,23)
$checkboxCLEAR.Checked = $False
$checkboxCLEAR.Text = "Clear lists when finished"

$checkboxLOG = New-Object Windows.Forms.Checkbox
$checkboxLOG.Location = New-Object System.Drawing.Point(810,435)
$checkboxLOG.Size = New-Object System.Drawing.Size(275,23)
$checkboxLOG.Checked = $True
$checkboxLOG.Text = "Log to file"

debugMsg "Checkboxes loaded."

$objLabelProgressBytes = New-Object System.Windows.Forms.Label
$objLabelProgressBytes.Location = New-Object System.Drawing.Size(50,485) 
$objLabelProgressBytes.Size = New-Object System.Drawing.Size(850,20) 
$objLabelProgressBytes.Text = "No copy in progress."

$my_ProgressBar_Bytes = New-Object System.Windows.Forms.ProgressBar
$my_ProgressBar_Bytes.Minimum = 0
$my_ProgressBar_Bytes.Maximum = $my_bytes_total
$my_ProgressBar_Bytes.Location = new-object System.Drawing.Size(50,505)
$my_ProgressBar_Bytes.size = new-object System.Drawing.Size(850,30)

$objLabelProgressNote = New-Object System.Windows.Forms.Label
$objLabelProgressNote.Location = New-Object System.Drawing.Size(50,545) 
$objLabelProgressNote.Size = New-Object System.Drawing.Size(850,20) 
$objLabelProgressNote.Text = "(copying engine idle)"

debugMsg "Progress bar loaded."

$labelTrace = New-Object Windows.Forms.Label
$labelTrace.Location = New-Object System.Drawing.Point(5,775)
$labelTrace.Size = New-Object System.Drawing.Size(175,20)
$labelTrace.Text = "Trace log:"

$listBoxTRACE = New-Object Windows.Forms.ListBox
$listBoxTRACE.Location = New-Object System.Drawing.Point(50,795)
$listBoxTRACE.Size = New-Object System.Drawing.Size(850,380)
$listBoxTRACE.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$listBoxTRACE.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$listBoxTRACE.AllowDrop = $True
$listBoxTRACE.HorizontalScrollbar = $True
logMsg ("Application launched.  Drag and drop folders to get started.")

debugMsg "Trace log elements loaded."

$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$statusBar.Text = "Awaiting folders (drag and drop)"

debugMsg "Form elements loaded."

$form.AcceptButton = $buttonCOPY
$form.CancelButton = $buttonCLOSE
$form.Controls.Add($labelSOURCE)
$form.Controls.Add($labelDEST)
$form.Controls.Add($listBoxSOURCE)
$form.Controls.Add($listBoxDEST)
$form.Controls.Add($buttonCOPY)
$form.Controls.Add($buttonRESET)
$form.Controls.Add($buttonCLOSE)
$form.Controls.Add($checkboxCLEAR)
$form.Controls.Add($checkboxLOG)
$form.Controls.Add($objLabelProgressBytes)
$form.Controls.Add($my_ProgressBar_Bytes)
$form.Controls.Add($objLabelProgressNote)
$form.Controls.Add($labelTrace)
$form.Controls.Add($listBoxTRACE)
$form.Controls.Add($statusBar)
$form.ResumeLayout()
debugMsg "Elements added to form."

### Event handlers ###

$buttonCOPY_Click = {
    logMsg "Copy button clicked."
    $global:copy_count = 0

    # Now that the GUI is responsive during a copy, we need to disable controls to prevent incorrect usage
    $listBoxSOURCE.Enabled = $False
    $listBoxDEST.Enabled = $False
    $buttonCOPY.Enabled = $False
    $buttonRESET.Enabled = $False
    $buttonCLOSE.Enabled = $False
    
    debugMsg "-- Copying files"
    debugMsg "-- Sources:      $($global:source_total)"
    debugMsg "-- Destinations: $($global:dest_total)"
        
    if ($global:source_total -lt 1 ) {
        $statusBar.Text = ("No source files specified.")
        logMsg ("Cannot copy - No source files specified.")
        return 1
    }
    if ($global:dest_total -eq 0 ) {
        $statusBar.Text = ("No destination specified.")
        logMsg ("Cannot copy - No destination specified.")
        return 1
    }

    foreach ($destination in $listBoxDEST.Items) {
        $dest = Get-Item -LiteralPath $destination
        if($dest -is [System.IO.DirectoryInfo]) {
            # destination confirmed as a folder.  Loop through source entries and copy them across...
            debugMsg ("---- Destination: `t" + $dest + " [Directory]")
            foreach ($item in $listBoxSOURCE.Items) {
                $global:copy_count = $global:copy_count + 1
                $statusBar.Text = ("Copying source entry $global:copy_count of $global:source_total...")
                $i = Get-Item -LiteralPath $item
                if($i -is [System.IO.DirectoryInfo]) {
                    #----DIRECTORY copy (Using Copy-Item, because CopyFileEx doesn't handle a source that is a folder)
                    logMsg ("  Copying directory: `t" + $i.Name + " [Directory]")
                    debugMsg ("  Copy-Item -LiteralPath """ + $i + """ -Destination """ + $dest + """ -Recurse")
                    $stopwatch_start = (Get-Date)
                    $objLabelProgressBytes.Text = "Copying directory: " + $i
                    $objLabelProgressNote.Text = "(interface may become unresponsive until completion, for large folders)"
                    [System.Windows.Forms.Application]::DoEvents()

                    try {
                        # CopyFileEx doesn't support directories, so still using Copy-Item
                        Copy-Item -LiteralPath $i -Destination $dest -Recurse
                        $stopwatch_end = (Get-Date)
                        logMsg ("  Directory copy complete (duration: " + $(format_duration (New-TimeSpan -Start $stopwatch_start -End $stopwatch_end) ) + ",  source entry $global:copy_count of $global:source_total)")
                    }
                    catch {
                        $stopwatch_end = (Get-Date)
                        logMsg ("  Directory copy failed (duration: " + $(format_duration (New-TimeSpan -Start $stopwatch_start -End $stopwatch_end) ) + ",  source entry $global:copy_count of $global:source_total)")
                    }
                    $objLabelProgressBytes.Text = " "
                    $objLabelProgressNote.Text = " "
                }
                else {
                    #----FILE copy (Using CopyFileEx with callback)
                    logMsg ("  Copying file: `t" + $i.Name + " [" + [math]::round($i.Length/1MB, 2) + " MB]")
                    debugMsg ("  Copy-File """ + $i + """ """ + $dest + """ ")
                    $stopwatch_start = (Get-Date)
                    [System.Windows.Forms.Application]::DoEvents()

                    $objLabelProgressBytes.Text = "Copying file: " + $i
                    Copy-File "$($i)" "$($dest)"
                    try {
                        $stopwatch_end = (Get-Date)
                        logMsg ("  File copy complete (duration: " + $(format_duration (New-TimeSpan -Start $stopwatch_start -End $stopwatch_end) ) + ",  source entry $global:copy_count of $global:source_total)")
                    }
                    catch {
                        $stopwatch_end = (Get-Date)
                        logMsg ("  File copy failed (duration: " + $(format_duration (New-TimeSpan -Start $stopwatch_start -End $stopwatch_end) ) + ",  source entry $global:copy_count of $global:source_total)")
                    }
                }
            }
        }
        else {
            # should not get here
            logMsg ("Destination skipped - `t" + $dest.Name + " [destination is not a directory]")
        }
    }
 
    if($checkboxCLEAR.Checked -eq $True) {
        $listBoxSOURCE.Items.Clear()
        $listBoxDEST.Items.Clear()
        logMsg ("File listings cleared.")
    }
 
    logMsg ("All tasks completed.")
    $statusBar.Text = ("All tasks completed.")
    $listBoxTRACE.Items.Add(" ")
    LogfileMsg ("  ")

    $my_ProgressBar_Bytes.Maximum = 1
    $my_ProgressBar_Bytes.Value = 1
    $objLabelProgressBytes.Text = " "
    $objLabelProgressNote.Text = "All tasks completed."

    $listBoxSOURCE.Enabled = $True
    $listBoxDEST.Enabled = $True
    $buttonCOPY.Enabled = $True
    $buttonRESET.Enabled = $True
    $buttonCLOSE.Enabled = $True
}

 
$listBoxSOURCE_DragOver = [System.Windows.Forms.DragEventHandler] {     # COPY
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = 'Copy'
    }
    else {
        $_.Effect = 'None'
    }
}

$listBoxDEST_DragOver = [System.Windows.Forms.DragEventHandler]{
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = 'Copy'
    }
    else {
        $_.Effect = 'None'
    }
}

$listBoxSOURCE_DragDrop = [System.Windows.Forms.DragEventHandler] {
    logMsg ("Drag-and-drop detected (source).")
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        $listBoxSOURCE.Items.Add($filename)
        logMsg ("Added to source list: $filename")
    }
    $global:source_total = $listBoxSOURCE.Items.Count
    $statusBar.Text = ("Source list contains $($global:source_total) items")
    debugMsg "Sources: $($global:source_total)"
}

$listBoxDEST_DragDrop = [System.Windows.Forms.DragEventHandler] {
    logMsg ("Drag-and-drop detected (destination).")
    $listBoxDEST.Items.Clear()
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        $listBoxDEST.Items.Add($filename)
    }
    foreach ($item in $listBoxDEST.Items) {
        $i = Get-Item -LiteralPath $item
        if($i -is [System.IO.DirectoryInfo]) {
            debugMsg ("Destination added: `t" + $i.Name + " [Directory]")
        }
        else {
            logMsg ("Destination rejected: `t" + $i.Name + " [not a directory]")
            $listBoxDEST.Items.Clear()
        }
    }
    $global:dest_total = $listBoxDEST.Items.Count
    $statusBar.Text = ("Destination list contains $($global:dest_total) items")
    debugMsg "Destinations: $($global:dest_total)"
}
 
$form_FormClosed = {
    debugMsg "FormClose function running."
    try {
        $buttonCOPY.remove_Click($buttonCOPY_Click)
        $buttonRESET.remove_Click($buttonRESET_Click)
        $buttonCLOSE.remove_Click($buttonCLOSE_Click)
        $listBoxSOURCE.remove_DragOver($listBoxSOURCE_DragOver)
        $listBoxDEST.remove_DragOver($listBoxDEST_DragOver)
        $listBoxSOURCE.remove_DragDrop($listBoxSOURCE_DragDrop)
        $listBoxDEST.remove_DragDrop($listBoxDEST_DragDrop)
        $form.remove_FormClosed($Form_Cleanup_FormClosed)
    }
    catch [Exception]
        { }
}

$buttonRESET_Click = {
    $statusBar.Text = ("RESET button clicked.")
    logMsg ("Source and destination lists reset.")
    $listBoxSOURCE.Items.Clear()
    $listBoxDEST.Items.Clear()
    }
 

 $buttonCLOSE_Click = {
    debugMsg "CLOSE button clicked"
    $statusBar.Text = ("CLOSE button clicked.")
    logMsg ("Application closed.")
    $form.Close()
    }

debugMsg "Event handler functions added."
 
### Wire up event handlers ###

$form.Add_Shown( {
    Hide-Console }
)
$buttonCOPY.Add_Click($buttonCOPY_Click)
$buttonRESET.Add_Click($buttonRESET_Click)
$buttonCLOSE.Add_Click($buttonCLOSE_Click)
$listBoxSOURCE.Add_DragOver($listBoxSOURCE_DragOver)
$listBoxSOURCE.Add_DragDrop($listBoxSOURCE_DragDrop)
$listBoxDEST.Add_DragOver($listBoxDEST_DragOver)
$listBoxDEST.Add_DragDrop($listBoxDEST_DragDrop)
$form.Add_FormClosed($form_FormClosed)
debugMsg "Event handlers linked to controls."

### Pre-flight checks ###

## Check longfilenames is enabled
$my_error = $False
if (Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem') {
    if (Test-RegistryValue 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' 'LongPathsEnabled') {
        $my_registry_value = Get-ItemPropertyValue -Path: 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' -Name LongPathsEnabled
        if ($my_registry_value -eq 1) {
            debugMsg "LongPathsEnabled confirmed as set."
        }
        else {
            LogMsg "Registry setting 'LongPathsEnabled' has an incorrect value.  Functionality may be limited."
            $my_error = $True
        }
    }
    else {
        LogMsg "Registry setting 'LongPathsEnabled' doesn't exist.  Functionality may be limited."
        $my_error = $True
    }
}
else {
    LogMsg "Registry setting 'LongPathsEnabled' path doesn't exist or is not accessible.  Functionality may be limited."
    $my_error = $True
}
if ($my_error) {
    LogMsg "For further information see:"
    LogMsg "https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry"
    LogMsg " "
}
debugMsg "LongPathsEnabled check complete."

## Check if running as Administrator
$user = [Security.Principal.WindowsIdentity]::GetCurrent();
if ((New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))  {
    LogMsg "Running as Administrator - drag and drop functionality may be impacted."
} else {
    debugMsg "NOT Running as Administrator"
}
debugMsg "Administrator check complete."

## Check if Powershell is 32-bit, if not relaunch in 64-bit mode.
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    LogMsg "32bit Powershell session detected.  Relaunching in 64bit."
    debugMsg "32bit Powershell session detected.  Relaunching in 64bit."
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 1
    if ($myInvocation.Line) {
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    }else{
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
    }
    pause
    exit $lastexitcode
}
debugMsg "32bit check complete."

debugMsg "Launching form..."
[void] $form.ShowDialog()
