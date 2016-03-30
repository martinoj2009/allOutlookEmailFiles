
set arg1=%1

reg load HKU\Subkey "C:\Users\ITSUS\NTUSER.DAT"
reg query HKCU\SOFTWARE\Microsoft\Office /s