
setlocal EnableDelayedExpansion

pushd c:\
set bcdedit=bcdedit.exe

    %bcdedit% /create {ramdiskoptions}
    %bcdedit% /set {ramdiskoptions} ramdisksdidevice boot
    %bcdedit% /set {ramdiskoptions} ramdisksdipath \Boot\boot.sdi


for /f "tokens=1-3 delims= " %%a in ('%bcdedit% ^/create ^/d ^"FMI Recovery^" ^/application osloader') do (
        set GUID=%%c
    )
REM %bcdedit%("/Create " & %GUID% & " -d """ & sDescription & """ /application OSLOADER")
	%bcdedit% /Set %GUID% systemroot \windows
	%bcdedit% /Set %GUID% detecthal yes
	%bcdedit% /Set %GUID% winpe yes
	%bcdedit% /set %GUID% osdevice ramdisk=[D:]\Deploy\Boot\LiteTouchPE_x86.wim,{ramdiskoptions}
	%bcdedit% /set %GUID% device ramdisk=[D:]\Deploy\Boot\LiteTouchPE_x86.wim,{ramdiskoptions}
	%bcdedit% /displayorder %GUID% /addlast	

bootsect /nt60 ALL

pause