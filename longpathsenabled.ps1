Function Test-RegistryValue ($regkey, $name) {
     if (Get-ItemProperty -Path $regkey -Name $name -ErrorAction Ignore) {
         $true
     } else {
         $false
     }
 }

### Check if the LongPathsEnabled registry setting has been enabled -
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
