msgbox"确定结束所有VBS程序？"
dim WSHshell 
set WSHshell = wscript.createobject("wscript.shell") 
WSHshell.run "taskkill /im wscript.exe /f ",0 ,true 
