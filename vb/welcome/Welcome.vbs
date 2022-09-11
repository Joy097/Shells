
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak "Hello Joy Sir" 

CreateObject("wscript.shell").run"C:\Users\User\Desktop\Notepad.txt"

a=msgbox("Do you have any bux pending ?",4+32,"Warning")
  
if a=vbYes then

    set shl = WScript.CreateObject("WScript.Shell")
   
    shl.run "https://bux.bracu.ac.bd/dashboard"
else 

     set shl = WScript.CreateObject("WScript.Shell")
   
     shl.run "https://www.youtube.com/watch?v=e8W2ByH9BCk"
     shl.run "https://www.facebook.com/"

     END IF