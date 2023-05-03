Function Test-RegistryValue ($regkey, $name) {
     if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
         $true
     } else {
         $false
     }
 }

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


### Check longfilenames is enabled -
    # Reference:    https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry
    # Registry key: Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem
    # Name:         LongPathsEnabled
    # Type:         REG_DWORD
    # Value:        0x01
    # PS Command:   New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
$my_error = $False
# debugMsg "Testing registry path..."
if (Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem') {
    # debugMsg "Registry path exists.  Checking for field name..."
    if (Test-RegistryValue 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' 'LongPathsEnabled') {
        # debugMsg "Registry name exists.  Checking if it has the correct value."
        $my_registry_value = Get-ItemPropertyValue -Path: 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\' -Name LongPathsEnabled
        if ($my_registry_value -eq 1) {
            # debugMsg "Registry field has the correct value."
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

# Form functions
function AddPrinter { 
  # ADDING PRINTER LOGIC GOES HERE
}

# Create a form, and the various elements that go on it
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$jobScript = {
    Start-Sleep -Seconds 5
}



function Extract() {
    $ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar.Location = New-Object System.Drawing.Point(10, 35)
    $ProgressBar.Size = New-Object System.Drawing.Size(460, 40)
    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 5

    $main_form.Controls.Add($ProgressBar);

    $Label.Font = $procFont
    $Label.ForeColor = 'red'
    $Label.Text = "Processing ..."
    $ProgressBar.visible

    $job = Start-Job -ScriptBlock $jobScript
    do { [System.Windows.Forms.Application]::DoEvents() } until ($job.State -eq "Completed")
    Remove-Job -Job $job -Force


    $Label.Text = "Process Complete"
    $ProgressBar.Hide()
    $StartButton.Hide()
    $EndButton.Visible
}



$LocalPrinterForm                    = New-Object system.Windows.Forms.Form
$LocalPrinterForm.ClientSize         = '500,300'
$LocalPrinterForm.text               = "LazyAdmin - PowerShell GUI Example"
$LocalPrinterForm.BackColor          = "#ffffff"

$my_Title                           = New-Object system.Windows.Forms.Label
$my_Title.text                      = "Adding new printer"
$my_Title.AutoSize                  = $true
$my_Title.width                     = 25
$my_Title.height                    = 10
$my_Title.location                  = New-Object System.Drawing.Point(20,20)
$my_Title.Font                      = 'Microsoft Sans Serif,13'

$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "Add a new construction site printer to your computer. Make sure you are connected to the network of the construction site."
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 50
$Description.location            = New-Object System.Drawing.Point(20,50)
$Description.Font                = 'Microsoft Sans Serif,10'

$PrinterStatus                   = New-Object system.Windows.Forms.Label
$PrinterStatus.text              = "Status:"
$PrinterStatus.AutoSize          = $true
$PrinterStatus.location          = New-Object System.Drawing.Point(20,115)
$PrinterStatus.Font              = 'Microsoft Sans Serif,10,style=Bold'

$PrinterFound                    = New-Object system.Windows.Forms.Label
$PrinterFound.text               = "Searching for printer..."
$PrinterFound.AutoSize           = $true
$PrinterFound.location           = New-Object System.Drawing.Point(75,115)
$PrinterFound.Font               = 'Microsoft Sans Serif,10'

# Add dropdown list, populate it and highlight the first entry
$PrinterType                     = New-Object system.Windows.Forms.ComboBox
$PrinterType.text                = ""
$PrinterType.width               = 170
$printerType.autosize            = $true
@('Canon','Hp') | ForEach-Object {[void] $PrinterType.Items.Add($_)}
$PrinterType.SelectedIndex       = 0
$PrinterType.location            = New-Object System.Drawing.Point(20,210)
$PrinterType.Font                = 'Microsoft Sans Serif,10'

$AddPrinterBtn                   = New-Object system.Windows.Forms.Button
$AddPrinterBtn.BackColor         = "#a4ba67"
$AddPrinterBtn.text              = "Add Printer"
$AddPrinterBtn.width             = 90
$AddPrinterBtn.height            = 30
$AddPrinterBtn.location          = New-Object System.Drawing.Point(370,250)
$AddPrinterBtn.Font              = 'Microsoft Sans Serif,10'
$AddPrinterBtn.ForeColor         = "#ffffff"
$AddPrinterBtn.Add_Click({ AddPrinter })

$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Cancel"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(260,250)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$LocalPrinterForm.CancelButton   = $cancelBtn

# FORM ELEMENTS ABOVE THIS LINE

### Load the elements onto the form and then display the form
$LocalPrinterForm.controls.AddRange(@($my_Title,$Description,$PrinterStatus,$PrinterFound,$PrinterType,$AddPrinterBtn,$cancelBtn))
$result = $LocalPrinterForm.ShowDialog()

if ($result –eq [System.Windows.Forms.DialogResult]::Cancel) {
    write-output 'User pressed cancel'
}

# write-host "Copying file..."
# Copy-File ".\Clockwork.avi" ".\TARGET\"
# write-host "Copy completed."
