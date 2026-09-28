del /q D:\REPOS\EXCEL\ExcelAgent\_bat_nuevo.txt
del /q D:\REPOS\EXCEL\ExcelAgent\_ml2.txt
del /q D:\REPOS\EXCEL\ExcelAgent\_ml3.txt
rmdir /s /q "%TEMP%\ea_build"
for /d %d in ("%TEMP%\contactos_*") do rmdir /s /q "%d"