

If Wscript.arguments.Count = 0 Then
	Wscript.echo "Missing Parameters"
Else
	For i = 0 to Wscript.arguments.Count - 1
		Wscript.echo Wscript.arguments(i)
	Next
End If

