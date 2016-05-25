@echo off


set PROG=ibm_utl_uxspi_9.51_winsrvr_32-64.exe
set UXSP_path=M:
net use M: \\fmi-data\DATA\IBMUPX /user:dep\dep dep

for /f %%a in ('dir %UXSP_path%\ibm_utl_uxspi*.exe /B /ON') do set LatestUXSPI=%%a

if "%LatestUXSPI%" == "" (
   set RC=%errorlevel%
   set RMSG=Could not find a UXSP installer, exiting script.
   echo Could not find a UXSP installer, exiting script.  
   echo Could not find a UXSP installer, exiting script.
echo %RC%
   goto UXSP_Finish
)

echo Found latest UXSP Installer: %LatestUXSPI%
echo.
echo.
%UXSP_path%\%LatestUXSPI% update -u -l %UXSP_path% -n -L
rem -s all
net use M: /delete

:UXSP_Finish
color 2F
echo ---------   Termine