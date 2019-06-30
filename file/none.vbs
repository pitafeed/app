on error resume next 
dim wsh,ye 
set wsh = createobject("wscript.shell") 
WshShell.AppActivate"" 
for i=1 to 100 
wscript.sleep 300 
wsh.sendkeys "^v" 
wsh.sendkeys i
wsh.sendkeys "{ENTER}"
next

