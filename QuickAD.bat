@echo off
setlocal EnableDelayedExpansion

:: --- Retrieve locked accounts from the last 3 months using PowerShell ---
echo Searching for locked accounts (from the last 3 months)...
set count=0
for /f "delims=" %%A in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "$threeMonthsAgo = (Get-Date).AddMonths(-3); $locked = Search-ADAccount -LockedOut | Get-ADUser -Properties lockoutTime | Select-Object GivenName,Surname,@{Name='sAMAccountName';Expression={$_.sAMAccountName.ToUpper()}},@{Name='LockoutTime';Expression={if ($_.lockoutTime -gt 0) {[datetime]::FromFileTime($_.lockoutTime).ToLocalTime()} else {$null}}} | Where-Object { $_.LockoutTime -and $_.LockoutTime -ge $threeMonthsAgo } | Sort-Object LockoutTime -Descending; $locked | Select-Object -ExpandProperty sAMAccountName"') do (
    set /a count+=1
    set "account[!count!]=%%A"
)

:: --- Check if locked accounts were found ---
if %count%==0 goto noLockedOption

:: --- List the locked accounts if any were found ---
echo.
echo Locked Accounts:
for /L %%i in (1,1,%count%) do (
    echo  %%i. !account[%%i]!
)
echo  0. Manually enter a username.

:menuPrompt
echo.
set /p choice="Enter the number to the account, or 0 to enter a username manually: "
if "%choice%"=="0" (
    set /p username="Enter the username: "
    goto processAccount
)
:: Validate numeric input
for /f "delims=0123456789" %%x in ("%choice%") do (
    if not "%%x"=="" (
        echo Invalid selection. Please enter a valid number.
        goto menuPrompt
    )
)
if %choice% GTR %count% (
    echo Invalid selection. Please try again.
    goto menuPrompt
)
set "username=!account[%choice%]!"
goto processAccount

:noLockedOption
echo.
echo No locked accounts found.
echo.
echo 1. Manually enter a username to process.
echo 0. Exit.
set /p noLockChoice="Please choose an option [1 or 0]: "
if "%noLockChoice%"=="1" (
    set /p username="Enter the username: "
    goto processAccount
) else if "%noLockChoice%"=="0" (
    goto end
) else (
    echo Invalid selection. Please try again.
    goto noLockedOption
)

:processAccount
echo.
echo Processing account: %username%

:: --- Unlock the account using PowerShell ---
echo Unlocking account...
powershell -NoProfile -Command "Unlock-ADAccount -Identity '%username%'"
if %errorlevel% EQU 0 (
    echo Account unlocked successfully.
) else (
    echo Failed to unlock account.
)

:: --- Reset password section and then force password change only if password reset succeeds ---
echo.
echo Would you like to reset the password for %username%? (Y/N)
set /p resetChoice="Enter Y to reset password, any other key to skip: "
if /I "%resetChoice%"=="Y" (
    set /p newPass="Enter new password: "
    if "!newPass!"=="" (
        echo No password entered. Skipping password reset.
        goto skipReset
    )
    echo Resetting password...
    powershell -NoProfile -Command "$securePass = ConvertTo-SecureString '!newPass!' -AsPlainText -Force; Set-ADAccountPassword -Identity '%username%' -NewPassword $securePass -Reset"
    if %errorlevel% EQU 0 (
        echo Password reset successfully.
        echo.
        echo Would you like to force a password change on next logon for %username%? (Y/N)
        set /p forceChoice="Enter Y to force password change, any other key to skip: "
        if /I "%forceChoice%"=="Y" (
            echo Setting password change requirement to true...
            powershell -NoProfile -Command "Set-ADUser -Identity '%username%' -ChangePasswordAtLogon $true"
            if %errorlevel% EQU 0 (
                echo User will be forced to change password at next logon.
            ) else (
                echo Failed to reset password.
            )
        ) else (
            if /I "%forceChoice%"=="y" (
            echo Setting password change requirement to true...
            powershell -NoProfile -Command "Set-ADUser -Identity '%username%' -ChangePasswordAtLogon $true"
            if %errorlevel% EQU 0 (
                echo User will be forced to change password at next logon.
            ) else (
                echo Failed to reset password.
            )
			) else (
				echo Setting password change requirement to false...
				powershell -NoProfile -Command "Set-ADUser -Identity '%username%' -ChangePasswordAtLogon $false"
			)
        )
    ) else (
        echo Skipping password change!
    )
)
:skipReset

:end
echo.
echo Script complete.
endlocal
