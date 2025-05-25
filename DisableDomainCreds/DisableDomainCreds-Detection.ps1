$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
$name = "DisableDomainCreds"
$expectedValue = 1

try {
    $actualValue = Get-ItemPropertyValue -Path $regPath -Name $name -ErrorAction Stop
    if ($actualValue -eq $expectedValue) {
        Write-Output "Compliant"
        exit 0
    }
    else {
        Write-Output "Not compliant"
        exit 1
    }
}
catch {
    Write-Output "Not compliant - registry key missing or error: $_"
    exit 1
}
