Dim init
set Fso = CreateObject("Scripting.FileSystemObject")
set InputFile = fso.OpenTextFile("C:\Users\antonio\Desktop\WhatsappBOT\contatos.txt")

Do While Not (InputFile.atEndOfStream)
     contatos = InputFile.ReadLine
     set init = CreateObject("WScript.Shell")
'init.Run("https://web.whatsapp.com/")'
      wscript.sleep 5000 
     init.Sendkeys "{TAB}"
     wscript.sleep 600 
init.SendKeys"" & contatos
WScript.sleep 3000
init.SendKeys "{ENTER}"
WScript.sleep 2000
init.SendKeys "^(V)"
WScript.sleep 6000

     init.SendKeys "Mensagem"

WScript.sleep 9000
       init.SendKeys "{ENTER}"
Loop

Wscript.Quit
