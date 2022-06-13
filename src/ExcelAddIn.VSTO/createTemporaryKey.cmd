set PATH=C:\Program Files (x86)\Windows Kits\10\bin\10.0.20348.0\x64;C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64;C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x64

makecert.exe ^
-n "CN=Allors Test Certificate" ^
-r ^
-pe ^
-a sha512 ^
-len 4096 ^
-cy authority ^
-sv ExcelAddIn.VSTO_TemporaryKey.pvk ^
ExcelAddIn.VSTO_TemporaryKey.cer
 
pvk2pfx.exe ^
-pvk ExcelAddIn.VSTO_TemporaryKey.pvk ^
-spc ExcelAddIn.VSTO_TemporaryKey.cer ^
-pfx ExcelAddIn.VSTO_TemporaryKey.pfx

del ExcelAddIn.VSTO_TemporaryKey.pvk
del ExcelAddIn.VSTO_TemporaryKey.cer