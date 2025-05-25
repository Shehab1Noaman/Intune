try {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
    $name = "DisableDomainCreds"
    $value = 1

    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    Set-ItemProperty -Path $regPath -Name $name -Value $value -Type DWord

    Write-Output "Remediation applied successfully."
    exit 0
}
catch {
    Write-Error "Remediation failed: $_"
    exit 1
}
