setlocal EnableDelayedExpansion

pushd c:\
set bcdedit=bcdedit.exe
%bcdedit% /createstore bcd
%bcdedit% /import bcd
del bcd
%bcdedit% /create {bootmgr}
%bcdedit% /set {bootmgr} device boot
%bcdedit% /set {bootmgr} inherit {globalsettings}
%bcdedit% /timeout 30

%bcdedit% /create {ntldr} /d "Windows XP Professionnel"
%bcdedit% /default {ntldr}
%bcdedit% /set {ntldr} device partition=C:
%bcdedit% /set {ntldr} path \ntldr
%bcdedit% /displayorder {ntldr} /addlast

bootsect /nt60 ALL /force