setlocal EnableDelayedExpansion

pushd D:\
set bcdedit=bcdedit.exe
%bcdedit% /createstore bcd
%bcdedit% /import bcd
del bcd
%bcdedit% /create {bootmgr}
%bcdedit% /set {bootmgr} device boot
%bcdedit% /set {bootmgr} inherit {globalsettings}
%bcdedit% /timeout 30

REM Recuperation du GUID pour W8
for /f "tokens=1-3 delims= " %%a in ('%bcdedit% ^/create ^/d ^"Windows Server 2008 R2^" ^/application osloader') do (
        set GUID=%%c
    )
%bcdedit% /default %GUID%
%bcdedit% /set {default} device partition=C:
%bcdedit% /set {default} path \windows\system32\winload.exe
%bcdedit% /set {default} locale fr-FR
%bcdedit% /set {default} inherit {bootloadersettings}
%bcdedit% /set {default} nx OptIn
%bcdedit% /set {default} osdevice partition=C:
%bcdedit% /set {default} systemroot \windows
%bcdedit% /set {default} detecthal Yes
%bcdedit% /displayorder {default} /addlast

bootsect /nt60 ALL /force