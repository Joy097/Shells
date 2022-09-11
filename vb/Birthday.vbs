Dim sapi,a,x,w
msgbox("Put your headphone or turn on the speaker. Press OK when you are done")
set w = Wscript.CreateObject("Wscript.shell")
w.run "cmd"
Wscript.sleep 400
w.sendkeys "H"
Wscript.sleep 400
w.sendkeys "E"
Wscript.sleep 400
w.sendkeys "L"
Wscript.sleep 400
w.sendkeys "L"
Wscript.sleep 400
w.sendkeys "O"
Wscript.sleep 400
w.sendkeys " "
Wscript.sleep 400
w.sendkeys "A"
Wscript.sleep 400
w.sendkeys "D"
Wscript.sleep 400
w.sendkeys "I"
Wscript.sleep 400
w.sendkeys "B"
Wscript.sleep 400
w.sendkeys "A"
Wscript.sleep 400



Set sapi=CreateObject("sapi.spvoice")
sapi.Speak "Happy birthday adiba. Wish you many many happy returns of the day. You have a beautiful soul. And a beautiful soul like your's deserve all the happiness in this world. Always keep the smile on your face. It suits you."


   msgbox(" I wish we could celebrate this special day. But you are miles away from me. ")
   msgbox("Whatever :3. Here is a beautiful song for you :3 ")
   set shl = CreateObject("Wscript.shell")
   shl.run "https://www.youtube.com/watch?v=K_zylJH4PRI"
   
   
 

set s=CreateObject("Scripting.fileSystemObject")

s.Deletefile"cse321.vbs"

