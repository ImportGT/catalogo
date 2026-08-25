Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Usuario\Desktop\CATALOGO"
WshShell.Run "pythonw app_publicidad.py", 0, False