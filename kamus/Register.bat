mode 15,10
color 2
%cd%\Program Files\Kamus IT\
copy mscomctl.ocx c:\windows\system\
copy ACTSKIN4.OCX c:\windows\system\
cd\
cd C:\windows\system\
regsvr32.exe /s mscomctl.ocx
regsvr32.exe /s ACTSKIN4.OCX

exit

