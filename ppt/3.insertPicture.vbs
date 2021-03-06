Dim oPicture 
Dim objPPT 
Dim objPres 
Dim curSlide

Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open("D:\Vbscript\ppt\test.pptx")

Set curSlide = objPres.Slides(1)

Set oPicture = curSlide.Shapes.AddPicture("D:\p\2.png", False, True, 10, 10	, -1, -1)
	oPicture.ScaleHeight 0.8, msoTrue
	oPicture.ScaleWidth 0.8, msoTrue

objPres.Save
objPres.Close

objPPT.Quit
Set objPPT = Nothing