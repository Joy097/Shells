Dim message, sapi
message=InputBox("what do you want me to say?","Speak to me")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message