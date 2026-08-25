Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Usuario\Desktop\CATALOGO"
WshShell.Run "cmd.exe /c python app_pedidos.py", 0, False