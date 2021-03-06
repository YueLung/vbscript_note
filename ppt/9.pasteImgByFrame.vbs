Dim objPPT
Dim objPres


Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open("D:\Vbscript\ppt\pasteImgByFrame\test.pptx")

For each curShp in objPres.Slides(1).Shapes
	If curShp.Type = 17 Then

		If InStr(curShp.TextFrame.TextRange,"1") Then
			objPres.Slides(1).Shapes.AddPicture "D:\p\2.png", False, True, curShp.Left, curShp.top, curShp.Width, curShp.Height
		End If
		
	End IF
Next

objPres.Save
objPres.Close

objPPT.Quit

Set objPPT = Nothing