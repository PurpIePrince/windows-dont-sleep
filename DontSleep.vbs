Set ws = CreateObject("WScript.shell")

Do

Wscript.Sleep 600000
'단위: 밀리 초

ws.SendKeys "{SCROLLLOCK}{SCROLLLOCK}"

Loop