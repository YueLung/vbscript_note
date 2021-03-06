'delete slide by contain '@' and '['

Set objPPT = CreateObject("PowerPoint.Application")
Set objPres = objPPT.Presentations.Open("D:\Vbscript\ppt\deleteSlideByTxt.pptx")

Set myDict = CreateObject("Scripting.Dictionary")
Set myDict2 = CreateObject("Scripting.Dictionary")
count = 0

For Each curSlide in objPres.Slides
	isFound = 0
	count = count +1
	For Each curShape in curSlide.Shapes
		If curShape.Type = 17 Then
		    if Instr(curShape.TextFrame.TextRange,"@") and Instr(curShape.TextFrame.TextRange,"[")then
				isFound = 1
				myDict2.add count,curSlide
				Exit For
			End if
		End IF
	Next
	
	If isFound = 0 then
		myDict.add count,curSlide
	End If
Next

For Each item in myDict.Items
	item.delete
Next 

For Each item in myDict2.Items
	objPres.Slides.AddSlide item.SlideIndex+1 ,item.CustomLayout
Next 

objPres.Save
objPres.Close

objPPT.Quit
Set objPPT = Nothing
