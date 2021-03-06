Option Explicit 

Dim objPPT 
Dim objPres 
Dim curSlide
Dim curShape

Dim slideCount : slideCount = 0
Dim shapeCount : shapeCount = 0

Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open("D:\Vbscript\ppt\test.pptx")

For Each curSlide in objPres.Slides
	slideCount = slideCount + 1
	For Each curShape in curSlide.Shapes
		shapeCount = shapeCount + 1
		
		'Wscript.echo(curShape.Type )
		
		If curShape.Type = 17 Then
			Dim ShpTxt
			'Dim TmpTxt
			
			Set ShpTxt = curShape.TextFrame.TextRange
			
			Wscript.echo(ShpTxt.Font.Size)
			'ShpTxt.Replace "replace","NOTLIKE",false
			'Set TmpTxt = ShpTxt.Replace("replace","NOTLIKE",false)
		End IF
		
	Next
Next


objPres.Save
objPres.Close

objPPT.Quit
Set objPPT = Nothing

'Wscript.echo("slideCount : " & slideCount & vbCrlf & "shapeCount : " & shapeCount)