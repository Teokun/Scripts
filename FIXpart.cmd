
set file=FIXpar1.SCP

REM Modification OSXP=H:, OS7=C:
echo sel DISK system > %file%
echo sel PAR 1 >> %file%
echo assign letter=H >> %file%
echo sel PAR 3 >> %file%
echo assign letter=C >> %file%
DISKPART /S %file%

