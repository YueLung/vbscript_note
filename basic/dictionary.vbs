Set myDict = CreateObject("Scripting.Dictionary")

myDict.Add "james", "98"
myDict.Add "kobe",  "96"
myDict.Add "kd",    "95"
myDict.Add "wade",  "94"


For Each item in myDict.Keys
	Wscript.echo item & " -> " & myDict.item(item)
Next 