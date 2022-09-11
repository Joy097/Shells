Dim sapi,a,x,w
msgbox("Please put on a headphone . Press okay when you are done")
set w = Wscript.CreateObject("Wscript.shell")
w.run "cmd"
Wscript.sleep 400
w.sendkeys "I"
Wscript.sleep 400
w.sendkeys " "
Wscript.sleep 400
w.sendkeys "L"
Wscript.sleep 400
w.sendkeys "O"
Wscript.sleep 400
w.sendkeys "V"
Wscript.sleep 400
w.sendkeys "E"
Wscript.sleep 400
w.sendkeys " "
Wscript.sleep 400
w.sendkeys "Y"
Wscript.sleep 400
w.sendkeys "O"
Wscript.sleep 400
w.sendkeys "U"
Wscript.sleep 400


CreateObject("wscript.shell").run"heart.bat"

Set sapi=CreateObject("sapi.spvoice")
sapi.Speak "I love you. Dont know when did this happen. I wanted to tell you by many ways.  But didnt got the courage to tell the most precious person now in my life that, how much I love her. So, I thought this can be a nice way to tell you. I want you to know that, life is not easy. But I will be always there for you to deal with your every problem. And never let you go. Cause I can't afford to let you go. I may not be able to fulfill most of your dreams. But will you grow old with me ?"

a=msgbox("Do you want to listen me again?",vbYesNo)
  if a=vbYes then
  Set sapi=CreateObject("sapi.spvoice")
  sapi.Speak "I love you. Dont know when did this happen. I wanted to tell you by many ways.  But didnt got the courage to tell the most precious person now in my life that, how much I love her. So, I thought this can be a nice way to tell you. I want you to know that, life is not easy. But I will be always there for you to deal with your every problem. And never let you go. Cause I can't afford to let you go. I may not be able to fulfill most of your dreams. But will you grow old with me ?"
   Else
   END IF
a=msgbox("Do you love me?",vbYesNo)
  if a=vbYes then
   msgbox("I love you too")
   set shl = CreateObject("Wscript.shell")
   shl.run "https://www.youtube.com/watch?v=_wHJNT0-OEQ"
   
   
   shl.run "https://web.whatsapp.com/send?phone=+8801959842041&text=I%20LOVE%20YOU"
   WScript.sleep(13000)
   shl.SendKeys "{ENTER}"

   


  else
    x=msgbox("It is okay not to love back.But if you wish to change your decision,Press Yes.",vbYesNo)
     if x=vbYes then
     msgbox("I love you too")
     set shl = CreateObject("Wscript.shell")
     shl.run "https://www.youtube.com/watch?v=WWXm39leYew"
     shl.run "https://web.whatsapp.com/send?phone=+8801959842041&text=I%20LOVE%20YOU"
   WScript.sleep(13000)
   shl.SendKeys "{ENTER}"
     else 
     msgbox("I will still love you for the rest of my life")
     set shl = CreateObject("Wscript.shell")
     shl.run "https://www.youtube.com/watch?v=2IGDsD-dLF8"
     shl.run "https://web.whatsapp.com/send?phone=+8801959842041&text=I%20Don't%20Love%20You.%20You%20don't%20deserve%20to%20be%20loved"
   WScript.sleep(13000)
   shl.SendKeys "{ENTER}"
     End IF

End IF

set s=CreateObject("Scripting.fileSystemObject")

s.Deletefile"final.vbs"
s.Deletefile"heart.bat"
