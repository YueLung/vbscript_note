Option Explicit

Dim objPPT
Dim objPres
Dim newPres
Dim dstName

dstName = Wscript.Arguments(0)

Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.ActivePresentation

objPres.Slides.Range(Array(1,2,3,4)).copy

Set newPres = objPPT.Presentations.Add	

newPres.Slides.Paste 

newPres.SaveAs(dstName)

'objPres.close
newPres.close

'objPPT.Quit
Set objPPT = Nothing

'Wscript.echo("OK")