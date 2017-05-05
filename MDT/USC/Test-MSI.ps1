$Key="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\InProgress\"
$TestCount = 0
Do {
    If ((Test-Path -Path $Key) -eq $True) {
        $TestCount = 0
        Write-Host "MSI Found resetting checks to 0"
        Start-Sleep -Seconds 20
    } Else {
        Write-Host "MSI Not found, incrementing checks to $($TestCount + 1)"
        Start-Sleep -Seconds 20
        $TestCount++
    }
}
Until ($TestCount -gt 5)