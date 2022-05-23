Dim message, sapi
message = "DEFAULT"
Set sapi=Createobject("sapi.spvoice")
Do While message <> "EXIT"
	message = InputBox("What do you want the computer to say? (Type EXIT to quit)", "Speech options")
	If message <> "EXIT" Then
		sapi.Speak message
	Else
	End If
Loop
