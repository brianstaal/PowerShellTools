<#
.SYNOPSIS
    Copies a public SSH key to a remote Linux host, similar to ssh-copy-id.

.DESCRIPTION
    This script finds your local public SSH key (typically id_rsa.pub or id_ed25519.pub)
    and appends it to the authorized_keys file on the specified remote host.

.PARAMETER Target
    The remote target in the format user@hostname or just hostname.

.PARAMETER p
    The port number to connect to. Default is 22.

.EXAMPLE
    .\ssh-copy-key.ps1 -p 2222 user@192.168.1.5
    Connects to 192.168.1.5 on port 2222 and installs the public key.
#>

param(
    [Parameter(Position=0, Mandatory=$true, HelpMessage="Remote target (e.g. user@hostname)")]
    [string]$Target,

    [Alias('Port')]
    [int]$p = 22
)

# Set up paths
$sshDir = "$env:USERPROFILE\.ssh"
$pubKeyFiles = @("id_ed25519.pub", "id_rsa.pub", "id_dsa.pub", "id_ecdsa.pub")
$foundKey = $null

# Check for existing public keys
if (Test-Path $sshDir) {
    foreach ($file in $pubKeyFiles) {
        $fullPath = Join-Path $sshDir $file
        if (Test-Path $fullPath) {
            $foundKey = $fullPath
            break
        }
    }
} else {
    Write-Error "SSH directory not found at $sshDir"
    exit 1
}

if (-not $foundKey) {
    Write-Error "No public key found in $sshDir. Please generate one using 'ssh-keygen'."
    exit 1
}

Write-Host "Found public key: $foundKey" -ForegroundColor Cyan
$keyContent = Get-Content -Path $foundKey -Raw

if (-not $keyContent) {
    Write-Error "Public key file is empty."
    exit 1
}

# Clean up the key content (trim whitespace)
$keyContent = $keyContent.Trim()

Write-Host "Attempting to copy key to $Target on port $p..." -ForegroundColor Cyan

# The command to run on the remote server
# 1. Create .ssh dir if not exists with 700 permissions
# 2. Append key to authorized_keys
# 3. Ensure authorized_keys has 600 permissions
$remoteCommand = "mkdir -p ~/.ssh && chmod 700 ~/.ssh && echo '$keyContent' >> ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys"

# Execute SSH command
try {
    # Using specific echo logic to avoid encoding issues across pipes
    ssh -p $p $Target $remoteCommand
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Successfully added key to '$Target'." -ForegroundColor Green
        Write-Host "Try logging in with: ssh -p $p $Target" -ForegroundColor Gray
    } else {
        Write-Error "Failed to copy key. SSH exited with code $LASTEXITCODE."
    }
} catch {
    Write-Error "An error occurred while executing ssh: $_"
}
