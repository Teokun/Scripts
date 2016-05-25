setlocal EnableDelayedExpansion

pushd C:\
set bcdedit=bcdedit.exe
%bcdedit% /createstore bcd
%bcdedit% /import bcd
del bcd
%bcdedit% /create {bootmgr}
%bcdedit% /set {bootmgr} device boot
%bcdedit% /timeout 30

%bcdedit% /create {ntldr} /d "Windows XP Professionnel"
%bcdedit% /set {ntldr} device partition=D:
%bcdedit% /set {ntldr} path \ntldr
%bcdedit% /displayorder {ntldr} /addlast

bootsect /nt60 ALL