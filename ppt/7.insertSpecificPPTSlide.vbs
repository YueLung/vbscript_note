Option Explicit

Dim objPPT
Dim objPres
Dim newPres
Dim dstName


Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open("D:\Vbscript\ppt\tmp.pptx")

objPres.Slides.InsertFromFile "D:\Vbscript\ppt\test.pptx",objPres.Slides.count

objPres.save
objPres.close

objPPT.Quit
Set objPPT = Nothing

'Wscript.echo("OK")