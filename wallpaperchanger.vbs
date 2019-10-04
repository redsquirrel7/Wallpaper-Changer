' Change Desktop Wallpaper
' Change sWallPaper variable to location of wallpaper file

dim wShell
dim theUser

Set wShell = WScript.CreateObject("WScript.Shell")
theUser = wShell.ExpandEnvironmentStrings("%USERNAME%")

Set oShell = CreateObject("WScript.Shell")
Set oFile = CreateObject("Scripting.FileSystemObject")

sDir = oFile.GetSpecialFolder(0)
sWallPaper = "C:\Users\Trevor\wallpaper.jpg"

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True