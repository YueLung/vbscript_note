Option Explicit

Dim objPPT
Dim objPres
Dim fileName

fileName = Wscript.Arguments(0)

Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open(fileName)

Set objPPT = Nothing