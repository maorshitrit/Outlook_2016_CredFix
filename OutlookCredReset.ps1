# TEST SCRIPT - Outlook Credential Reset Testing
# This simulates password changes without needing Outlook

$LogFile = "$env:TEMP\OutlookCredReset_TEST.log"
$PasswordHashFile = "$env:APPDATA\OutlookPasswordHash_TEST.txt"

function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Out-File -FilePath $LogFile -Append
    Write-Host "$Timestamp - $Message" -ForegroundColor Cyan
}

function Get-CurrentPasswordHash {
    try {
        $ADUser = ([ADSISEARCHER]"samaccountname=$env:USERNAME").FindOne()
        if ($ADUser) {
            $PwdLastSet = $ADUser.Properties.pwdlastset[0]
            Write-Log "Retrieved password hash from AD: $PwdLastSet"
            return $PwdLastSet
        }
    }
    catch {
        Write-Log "Error getting password hash: $_"
    }
    return $null
}

function Test-CredentialCleanup {
    Write-Log "=== TESTING CREDENTIAL CLEANUP ==="
    
    # Check for Outlook process (will be none in test env)
    $OutlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($OutlookProcess) {
        Write-Log "Outlook is running - would close it"
    } else {
        Write-Log "Outlook not running (expected in test environment)"
    }
    
    # List credentials that would be deleted
    Write-Log "Checking Credential Manager entries..."
    $CredList = cmdkey /list
    $OutlookCreds = $CredList | Select-String "(MicrosoftOffice|outlook|office365|autodiscover|exchange)"
    
    if ($OutlookCreds) {
        Write-Log "Found Outlook-related credentials:"
        foreach ($cred in $OutlookCreds) {
            if ($cred.Line -match "Target:\s*(.+)") {
            $TargetName = $matches[1].Trim()
            Write-Log "Deleting Outlook credential: $TargetName"
            cmdkey /delete:"$TargetName"
    }
}

    } else {
        Write-Log "No Outlook-related credentials found (expected in test environment)"
    }
    
    # Check registry path
    $OutlookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook"
    if (Test-Path $OutlookProfilePath) {
        Write-Log "Outlook profile registry path exists"
    } else {
        Write-Log "Outlook profile registry path NOT found (expected in test environment)"
    }
    
    Write-Log "=== CLEANUP TEST COMPLETED ==="
    return $true
}

# ===== MAIN TEST SCRIPT =====
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "OUTLOOK CREDENTIAL RESET - TEST MODE" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow

Write-Log "=== Test Started for User: $env:USERNAME ==="

# Get current password hash
$CurrentHash = Get-CurrentPasswordHash

if ($CurrentHash) {
    Write-Host "`n[1] Current Password Hash: $CurrentHash" -ForegroundColor Green
    
    # Check if hash file exists
    if (Test-Path $PasswordHashFile) {
        $StoredHash = Get-Content $PasswordHashFile
        Write-Host "[2] Stored Hash File Found: $StoredHash" -ForegroundColor Green
        
        if ($StoredHash -ne $CurrentHash) {
            Write-Host "`n*** PASSWORD CHANGE DETECTED! ***" -ForegroundColor Red
            Write-Host "Old Hash: $StoredHash" -ForegroundColor Yellow
            Write-Host "New Hash: $CurrentHash" -ForegroundColor Yellow
            
            Write-Log "Password change detected!"
            
            # Run cleanup test
            Test-CredentialCleanup
            
            # Update hash file
            $CurrentHash | Out-File -FilePath $PasswordHashFile -Force
            Write-Host "`n[3] Hash file UPDATED with new value" -ForegroundColor Green
            
        } else {
            Write-Host "`n[2] NO PASSWORD CHANGE - Hashes match" -ForegroundColor Green
            Write-Log "No password change detected"
        }
    } else {
        Write-Host "[2] First Run - Creating hash file" -ForegroundColor Yellow
        Write-Log "First run - creating password hash file"
        $CurrentHash | Out-File -FilePath $PasswordHashFile -Force
        Write-Host "[3] Hash file CREATED: $PasswordHashFile" -ForegroundColor Green
    }
} else {
    Write-Host "[ERROR] Could not retrieve password hash from AD" -ForegroundColor Red
    Write-Log "Could not retrieve password hash"
}

# Show file locations
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "Test Files Location:" -ForegroundColor Yellow
Write-Host "Log File: $LogFile" -ForegroundColor Cyan
Write-Host "Hash File: $PasswordHashFile" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Yellow

Write-Log "=== Test Completed ==="

# Open log file
Write-Host "Opening log file..." -ForegroundColor Green
#Start-Sleep -Seconds 2
#notepad $LogFile