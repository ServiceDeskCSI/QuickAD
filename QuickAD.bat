@echo off
setlocal enabledelayedexpansion

REM ========================================================
REM Check for administrative privileges
NET SESSION >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as an administrator.
    pause
    exit /b
)

echo ================================================
echo          Active Directory Unlock Tool
echo ================================================
echo.

REM ========================================================
REM Get count of locked accounts from the last 3 months using PowerShell
for /f "delims=" %%C in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "$threeMonthsAgo = (Get-Date).AddMonths(-3); $locked = Search-ADAccount -LockedOut | Get-ADUser -Properties lockoutTime | Select-Object GivenName,Surname,@{Name='sAMAccountName';Expression={$_.sAMAccountName.ToUpper()}},@{Name='LockoutTime';Expression={if ($_.lockoutTime -gt 0) {[datetime]::FromFileTime($_.lockoutTime).ToLocalTime()} else {$null}}} | Where-Object { $_.LockoutTime -and $_.LockoutTime -ge $threeMonthsAgo } | Sort-Object LockoutTime -Descending; Write-Output $locked.Count"') do set lockedCount=%%C

if "%lockedCount%"=="0" (
    echo No locked out accounts found in the last 3 months.
    echo.
    echo Hit any key to close the program.
    pause >nul
    exit /b
)

REM ========================================================
REM List locked accounts (Last 3 Months, Most Recent at Top)
echo Locked Accounts (Last 3 Months, Most Recent at Top):
powershell -NoProfile -ExecutionPolicy Bypass -Command "$threeMonthsAgo = (Get-Date).AddMonths(-3); $locked = Search-ADAccount -LockedOut | Get-ADUser -Properties lockoutTime | Select-Object GivenName,Surname,@{Name='sAMAccountName';Expression={$_.sAMAccountName.ToUpper()}},@{Name='LockoutTime';Expression={if ($_.lockoutTime -gt 0) {[datetime]::FromFileTime($_.lockoutTime).ToLocalTime()} else {$null}}} | Where-Object { $_.LockoutTime -and $_.LockoutTime -ge $threeMonthsAgo } | Sort-Object LockoutTime -Descending; if ($locked) { $i=1; foreach ($user in $locked) { Write-Output (\"$i. $($user.sAMAccountName) ($($user.LockoutTime.ToString('g')))\"); $i++ } }"
echo.

REM ========================================================
REM Prompt for the account number to unlock
set /p choice=Enter the number of the account to unlock: 
if "%choice%"=="" (
    echo ERROR: No selection made.
    pause
    exit /b
)

REM Retrieve the selected account's sAMAccountName using PowerShell
for /f "delims=" %%A in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "$threeMonthsAgo = (Get-Date).AddMonths(-3); $locked = Search-ADAccount -LockedOut | Get-ADUser -Properties lockoutTime | Select-Object GivenName,Surname,@{Name='sAMAccountName';Expression={$_.sAMAccountName.ToUpper()}},@{Name='LockoutTime';Expression={if ($_.lockoutTime -gt 0) {[datetime]::FromFileTime($_.lockoutTime).ToLocalTime()} else {$null}}} | Where-Object { $_.LockoutTime -and $_.LockoutTime -ge $threeMonthsAgo } | Sort-Object LockoutTime -Descending; if ($locked.Count -ge %choice%) { $user = $locked[%choice%-1]; Write-Output $user.sAMAccountName } else { exit 1 }"') do set selectedUser=%%A

if "%selectedUser%"=="" (
    echo ERROR: Invalid selection or no user found.
    pause
    exit /b
)

echo.
echo Selected User: %selectedUser%
echo.

REM ========================================================
REM Show account details for the selected user
echo Retrieving account details for "%selectedUser%"...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ADUser -Identity '%selectedUser%' -Properties AccountLockoutTime, BadLogonCount | Select-Object Name, AccountLockoutTime, BadLogonCount | Format-Table -AutoSize"
echo.

REM ========================================================
REM Unlock the account using PowerShell's Unlock-ADAccount cmdlet
echo Running unlock command for "%selectedUser%"...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module ActiveDirectory; try { Unlock-ADAccount -Identity '%selectedUser%' -ErrorAction Stop; } catch { Write-Error 'Unlock-ADAccount failed for %selectedUser%'; exit 1 }"
if %errorlevel% neq 0 (
    echo ERROR: Failed to unlock "%selectedUser%". Please verify the username and try again.
    pause
    exit /b
)
echo.

REM ========================================================
REM Verify the unlock status and display confirmation using PowerShell
echo Verifying account status for "%selectedUser%":
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module ActiveDirectory; $acct = Get-ADUser -Identity '%selectedUser%' -Properties LockedOut; Write-Host ''; Write-Host '========== ACCOUNT UNLOCK STATUS ==========' -ForegroundColor Yellow; if ($acct.LockedOut -eq $false) { Write-Host ('SUCCESS: The account %selectedUser% is now UNLOCKED.') -ForegroundColor Green; } else { Write-Host ('WARNING: The account %selectedUser% is STILL LOCKED.') -ForegroundColor Red; }; Write-Host '==========================================' -ForegroundColor Yellow; Write-Host ''"
echo.

REM ========================================================
REM Ask if the administrator wants to change the password
:askPassword
set /p changePass=Change user's password? [Y/N]: 
if /I "%changePass%"=="Y" goto ChangePass
if /I "%changePass%"=="N" goto EndScript
echo Please enter Y or N.
goto askPassword

:ChangePass
echo.
set /p newPass=Enter new password: 
if "%newPass%"=="" (
    echo ERROR: No password entered.
    pause
    exit /b
)
echo.
REM Ask if force password change at next logon is required
:askForce
set /p forceChange=Force password change at next logon? [Y/N]: 
if /I "%forceChange%"=="Y" (
    set forceFlag=Yes
) else if /I "%forceChange%"=="N" (
    set forceFlag=No
) else (
    echo Please enter Y or N.
    goto askForce
)

echo Resetting password for "%selectedUser%"...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module ActiveDirectory; Set-ADAccountPassword -Identity '%selectedUser%' -Reset -NewPassword (ConvertTo-SecureString '%newPass%' -AsPlainText -Force)"
if %errorlevel% neq 0 (
    echo ERROR: Failed to reset the password for "%selectedUser%".
    pause
    exit /b
)
if /I "%forceFlag%"=="Yes" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Import-Module ActiveDirectory; Set-ADUser -Identity '%selectedUser%' -ChangePasswordAtLogon $true"
    echo Password reset and force change on next login is enabled.
) else (
    echo Password reset without forcing change on next login.
)
goto EndScript

:EndScript
echo.
echo Script Done!
pause
endlocal
