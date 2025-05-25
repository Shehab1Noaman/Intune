Try {
    # Get a list of stored credentials
    $creds = cmdkey /list | Where-Object { $_ -match "Target:" } | ForEach-Object {
        ($_ -split "=")[1].Trim()
    }

    # Delete each credential
    foreach ($target in $creds) {
        Write-Host "Deleting credential: $target"
        cmdkey /delete:$target
    }

    Write-Output "Script executed successfully"
    Exit 0
}
Catch {
    Write-Error "Script failed: $_"
    Exit 1
}
