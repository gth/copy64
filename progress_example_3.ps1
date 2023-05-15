# [void][system.reflection.assembly]::LoadWithPartialName("System.Drawing")
# [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition '
public class DPIAware {
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'
[System.Windows.Forms.Application]::EnableVisualStyles()
[void] [DPIAware]::SetProcessDPIAware()

Function Test-RegistryValue ($regkey, $name) {
    if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
        $true
    } else {
        $false
    }
}

### ------------ MAIN COPY FUNCTION ------------ ###
function Copy-File {
## Credit to https://github.com/FranciscoNabas/PowerShellPublic/blob/main/Copy-File.ps1
## for the key parts of the function below, making CopyFileEx() useable in PowerShell.
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
    [Func[long, long, long, long, System.UInt32, Utilities.FileSystem+CopyProgressCallbackReason, System.IntPtr, System.IntPtr, System.IntPtr, Utilities.FileSystem+CopyProgressResult]]$copyProgressDelegate = {
        param($total, $transfered, $streamSize, $streamByteTrans, $dwStreamNumber, $reason, $hSourceFile, $hDestinationFile, $lpData)
        Write-Progress -Activity "Copying file" -Status "$Path ~> $Destination. $([Math]::Round(($transfered/1KB), 2))KB/$([Math]::Round(($total/1KB), 2))KB." -PercentComplete (($transfered / $total) * 100)
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


### Check longfilenames is enabled -
    # Reference:    https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry
    # Registry key: Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem
    #               LongPathsEnabled  REG_DWORD  0x01
    # TO ADD:   
    #               PS Command New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
$my_error = $False
# debugMsg "Testing registry path..."
if (Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem') {
    # debugMsg "Registry path exists.  Checking for field name..."
    if (Test-RegistryValue 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' 'LongPathsEnabled') {
        # debugMsg "Registry name exists.  Checking if it has the correct value."
        $my_registry_value = Get-ItemPropertyValue -Path: 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' -Name LongPathsEnabled
        if ($my_registry_value -eq 1) {
            # debugMsg "Registry field has the correct value."
            Write-host ("[INF] Windows long path support detected.") -ForegroundColor DarkCyan
        }
        else {
            write-warning "LongPathsEnabled registry field has the incorrect value.  Functionality may be limited."
            $my_error = $True
        }
    }
    else {
        write-warning "LongPathsEnabled registry field doesn't exist.  Functionality may be limited."
        $my_error = $True
    }
}
else {
    write-warning "LongPathsEnabled registry path doesn't exist or is not accessible.  Functionality may be limited."
    $my_error = $True
}
if ($my_error) {
    write-warning "Review the details shown at the link below for further information:"
    write-warning "https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry"
    write-warning " "
}


$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Testing"
$objForm.Size = New-Object System.Drawing.Size(800,900)  # (width, height)
$objForm.FormBorderStyle = 'Fixed3D'
$objForm.MaximizeBox = $false
$objForm.MinimizeBox = $false
$objForm.StartPosition = "CenterScreen"
$objForm.KeyPreview = $True

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(0,20)  # (x,y)
$objLabel.Size = New-Object System.Drawing.Size(280,20)     # (width, height)
$objLabel.Text = "Source(s):"

$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,40) 
$objListBox.Size = New-Object System.Drawing.Size(700,20) 
$objListBox.Height = 100

$objLabelDest = New-Object System.Windows.Forms.Label
$objLabelDest.Location = New-Object System.Drawing.Size(10,160) 
$objLabelDest.Size = New-Object System.Drawing.Size(280,20) 
$objLabelDest.Text = "Destination:"

$objListBoxDest = New-Object System.Windows.Forms.ListBox 
$objListBoxDest.Location = New-Object System.Drawing.Size(10,180) 
$objListBoxDest.Size = New-Object System.Drawing.Size(700,20) 
$objListBoxDest.Height = 100

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(10,280)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "&Copy"
# $OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})

$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Location = New-Object System.Drawing.Size(550,280)
$ResetButton.Size = New-Object System.Drawing.Size(75,23)
$ResetButton.Text = "&Reset"

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(635,280)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "C&ancel"
$CancelButton.Add_Click({$objForm.Close()})

$my_barheight = 20
$my_barlength = 700
$my_currentbarposition = 380
$my_dest_total   = 1
$my_source_total = 1
$my_folder_total = 1
$my_bytes_total  = 1

## Add progress bars: Destinations -> Sources -> Files/Folders -> Bytes

# Destination progress bar
$objLabelProgressDest = New-Object System.Windows.Forms.Label
$objLabelProgressDest.Location = New-Object System.Drawing.Size(10,$my_currentbarposition) 
$objLabelProgressDest.Size = New-Object System.Drawing.Size($my_barlength,$my_barheight) 
$objLabelProgressDest.Text = "No destination(s) selected."
$my_currentbarposition = $my_currentbarposition + $my_barheight 

$my_ProgressBar_Dest = New-Object System.Windows.Forms.ProgressBar
$my_ProgressBar_Dest.Minimum = 0
$my_ProgressBar_Dest.Maximum = $my_dest_total
$my_ProgressBar_Dest.Location = new-object System.Drawing.Size(10,$my_currentbarposition)
$my_ProgressBar_Dest.size = new-object System.Drawing.Size($my_barlength,$my_barheight)
$my_currentbarposition = $my_currentbarposition + $my_barheight + $my_barheight


# Source progress bar
$objLabelProgressSource = New-Object System.Windows.Forms.Label
$objLabelProgressSource.Location = New-Object System.Drawing.Size(10,$my_currentbarposition) 
$objLabelProgressSource.Size = New-Object System.Drawing.Size($my_barlength,$my_barheight) 
$objLabelProgressSource.Text = "No soure(s) selected."
$my_currentbarposition = $my_currentbarposition + $my_barheight 

$my_ProgressBar_Source = New-Object System.Windows.Forms.ProgressBar
$my_ProgressBar_Source.Minimum = 0
$my_ProgressBar_Source.Maximum = $my_source_total
$my_ProgressBar_Source.Location = new-object System.Drawing.Size(10,$my_currentbarposition)
$my_ProgressBar_Source.size = new-object System.Drawing.Size($my_barlength,$my_barheight)
$my_currentbarposition = $my_currentbarposition + $my_barheight + $my_barheight

# File/folder progress bar
$objLabelProgressFile = New-Object System.Windows.Forms.Label
$objLabelProgressFile.Location = New-Object System.Drawing.Size(10,$my_currentbarposition) 
$objLabelProgressFile.Size = New-Object System.Drawing.Size($my_barlength,$my_barheight) 
$objLabelProgressFile.Text = "No source files or folders specified."
$my_currentbarposition = $my_currentbarposition + $my_barheight 

$my_ProgressBar_Folder = New-Object System.Windows.Forms.ProgressBar
$my_ProgressBar_Folder.Minimum = 0
$my_ProgressBar_Folder.Maximum = $my_folder_total
$my_ProgressBar_Folder.Location = new-object System.Drawing.Size(10,$my_currentbarposition)
$my_ProgressBar_Folder.size = new-object System.Drawing.Size($my_barlength,$my_barheight)
$my_currentbarposition = $my_currentbarposition + $my_barheight  + $my_barheight


# Add bytes progress bar (bytes progress is always displayed, even if it target has zero bytes)
$objLabelProgressBytes = New-Object System.Windows.Forms.Label
$objLabelProgressBytes.Location = New-Object System.Drawing.Size(10,$my_currentbarposition) 
$objLabelProgressBytes.Size = New-Object System.Drawing.Size($my_barlength,$my_barheight) 
$objLabelProgressBytes.Text = "No copying in progress."
$my_currentbarposition = $my_currentbarposition + $my_barheight 

$my_ProgressBar_Bytes = New-Object System.Windows.Forms.ProgressBar
$my_ProgressBar_Bytes.Minimum = 0
$my_ProgressBar_Bytes.Maximum = $my_bytes_total
$my_ProgressBar_Bytes.Location = new-object System.Drawing.Size(10,$my_currentbarposition)
$my_ProgressBar_Bytes.size = new-object System.Drawing.Size($my_barlength,$my_barheight)
$my_currentbarposition = $my_currentbarposition + $my_barheight  + $my_barheight

$objForm.Controls.Add($objLabel) 
$objForm.Controls.Add($objListBox) 
$objForm.Controls.Add($objLabelDest) 
$objForm.Controls.Add($objListBoxDest) 
$objForm.Controls.Add($OKButton)
$objForm.Controls.Add($ResetButton)
$objForm.Controls.Add($CancelButton)
$objForm.Controls.Add($objLabelProgressDest)
$objForm.Controls.Add($my_ProgressBar_Dest)
$objForm.Controls.Add($objLabelProgressSource)
$objForm.Controls.Add($my_ProgressBar_Source)
$objForm.Controls.Add($objLabelProgressFile)
$objForm.Controls.Add($my_ProgressBar_Folder)
$objForm.Controls.Add($objLabelProgressBytes)
$objForm.Controls.Add($my_ProgressBar_Bytes)

# $objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()


# function copy_button

# Progress - Destinations
for ($my_destination = 1 ; $my_destination -le $my_dest_total ; $my_destination++) {
    write-host "Destination: $my_destination out of $my_dest_total"
    $my_ProgressBar_Dest.Value = $my_destination

    # Progress - This Source / All Sources
    for ($my_source = 1 ; $my_source -le $my_source_total ; $my_source++) {
        write-host "Source: $my_source out of $my_source_total"
        $my_ProgressBar_Source.Value = $my_source

        # Progress - This Source's Files and Folders
        for ($my_file_count = 1 ; $my_file_count -le $my_folder_total ; $my_file_count++) {
            # write-host "File $my_file_count out of $my_folder_total"
            $my_ProgressBar_Folder.Value = $my_file_count

            # Progress - This Particular File
            for ($my_bytes = 1 ; $my_bytes -le $my_bytes_total ; $my_bytes++) {
                # Write-Host -NoNewLine "`r Copying $my_bytes out of $my_bytes_total..."
                # Start-Sleep -Seconds 1
                # if ($my_Checkbox_finish = $False) {
                    #Start-Sleep -Milliseconds 5
                # }
                $my_ProgressBar_Bytes.Value = $my_bytes
                $objForm.Add_Shown({$objForm.Activate()})
                $Form = $objForm.Show()
            }
            # Write-Host " "
        }
    }
}


# close up shop
$Form = $objForm.Close()
$objForm.Dispose()
