StartProgram "D:\secu\bin\freadv2.exe", "", hmiShowNormal, hmiYes
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim WSHSmessage_codeerror, Msgcodeerror
Dim WSHSmessage_code, Msgcode
Dim WSHSmessage_null, Msgnull
Dim WSHSmessage_send, Msgsend
Dim WSHSmessage_abs, Msgabs
Dim fsobis, fb
Dim fsoraz, fraz
Dim txt1r
Dim fsoread, fr
Dim fsocreat,fc
Dim fsoadmi,fa
Dim fsowrit, fw
Dim fso, f
Dim myident
Dim ident4
Dim defaut 
Dim readbegin
Dim forloop
Dim admin_mtp
Dim admin_sup
Dim admin_sup2
Dim cpt
Dim sec1
Dim sec2 
Dim secconv
Dim sec3
Dim sec4
Dim secconv2
myident = ""
num4 = ""
f = ""
fso = ""
readbegin = ""
forloop = 0
admin_mtp = ""
defaut = ""
sec1 = Second(Now)
 Do While secconv < 3
 sec2 = Second(Now)
 secconv = sec2 - sec1 
 Loop
 
Set fsoadmi = CreateObject("Scripting.FileSystemObject")
Set fa = fsoadmi.OpenTextFile("d:\secu\admin.txt", ForReading)
admin_mtp = fa.ReadLine
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("d:\secu\testfile.txt", ForReading)

If bp2 = True Then
  bp2 = False
    Select Case name
 
      Case "long"
        Do Until f.AtEndOfStream
 
          If num4 = "long" Then
            ghost = "ghost1"

            SendEMail"longuevergn@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
             Set WSHSmessage_send = CreateObject("Wscript.Shell")
             Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, 
            "Récupération Login - Mtp", 64)
         End If

         myident = f.ReadLine
        num4 = myident
      Loop
      
    Case "chat"
      Do Until f.AtEndOfStream
         If num4 = "chat" Then
           SendEMail"chatelet@ipno.in2p3.fr", "rappel de votre identifiant 
          - securite 106", myident, "chapoul@ipno.in2p3.fr"
           Set WSHSmessage_send = CreateObject("Wscript.Shell")
           Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
         End If
         
        myident = f.ReadLine
       num4 = myident
      Loop
 
    Case "gand"
      Do Until f.AtEndOfStream
         If num4 = "gand" Then
           SendEMail"gandolfo@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
           Set WSHSmessage_send = CreateObject("Wscript.Shell")
           Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
         End If
         
       myident = f.ReadLine
       num4 = myident
      Loop
      
  Case "chap"
    Do Until f.AtEndOfStream
       If num4 = "chap" Then
         ghost = "5"
         SendEMail"chapoul@ipno.in2p3.fr", "rappel de votre identifiant -securite 106", myident, "chapoul@ipno.in2p3.fr"
         Set WSHSmessage_send = CreateObject("Wscript.Shell")
         Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
       End If

     myident = f.ReadLine
     num4 = myident
    Loop
    
    Case "lesr"
      Do Until f.AtEndOfStream
         If num4 = "lesr" Then
           SendEMail"lesrel@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
           Set WSHSmessage_send = CreateObject("Wscript.Shell")
           Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
       End If
       
       myident = f.ReadLine
       num4 = myident
      Loop
 
   Case "gesl"
     Do Until f.AtEndOfStream
       If num4 = "gesl" Then
         SendEMail"geslin@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
         Set WSHSmessage_send = CreateObject("Wscript.Shell")
         Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
       End If
       
     myident = f.ReadLine
     num4 = myident
     Loop
 
  Case "joly"
     Do Until f.AtEndOfStream
       If num4 = "joly" Then
         SendEMail"joly@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
         Set WSHSmessage_send = CreateObject("Wscript.Shell")
         Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
       End If
       
     myident = f.ReadLine
     num4 = myident
    Loop
  
  Case "olry"
     Do Until f.AtEndOfStream
       If num4 = "olry" Then
         SendEMail"olry@ipno.in2p3.fr", "rappel de votre identifiant - securite 106", myident, "chapoul@ipno.in2p3.fr"
         Set WSHSmessage_send = CreateObject("Wscript.Shell")
         Msgsend= WSHSmessage_send.Popup("Votre identifiant et vôtre mot de passe vous ont étés envoyés, veuillez consulter votre boite mail.", 5, "Récupération Login - Mtp", 64)
       End If

     myident = f.ReadLine
     num4 = myident
    Loop

  Case Else
     Set WSHSmessage_abs = CreateObject("Wscript.Shell")
     Msgabs= WSHSmessage_abs.Popup("Login inéxistant.", 5, "Erreur 
    d'éxécution", 48)
  End Select
End If 
  
If mtp_change = admin_mtp And bp = True And name <> Empty Then
   Set fsoread = CreateObject("Scripting.FileSystemObject")
   Set fr = fsoread.OpenTextFile("d:\secu\testfile.txt", ForReading)
   
      Do Until fr.AtEndOfStream
       readbegin = fr.ReadLine 
       num4 = readbegin 
       Set fsowrit = CreateObject("Scripting.FileSystemO bject")
       Set fw = fsowrit.OpenTextFile("d:\secu\testfiles. txt",ForAppending)

       If num4 <> name Then 
         fw.WriteLine readbegin 
         fw.Close 
       End If

       If num4 = name Then
         fw.WriteLine mtp_new
         fw.Close 
         admin_sup = mtp_new 
         admin_sup2 = name 
         SendEMail"chapoul@ipno.in2p3.fr", "Nouveau mot d e passe 
        utilisateur rentre - mtp", admin_sup , "chapoul@ipno.in2p3.fr"
         SendEMail"chapoul@ipno.in2p3.fr", "Nouveau mot d e passe 
        utilisateur rentré - login", admin_sup2 , "chapoul@ipno.in2p3.fr"
       End If 
       
     Loop

   Set fsowrit = CreateObject("Scripting.FileSystemObject")
   Set fw = fsowrit.OpenTextFile("d:\secu\testfile.txt", ForWriting)
   fw.Write ""
   fw.Close 
   Set fsoread = CreateObject("Scripting.FileSystemObject")
   Set fr = fsoread.OpenTextFile("d:\secu\testfiles.txt", ForReading)

   Do Until fr.AtEndOfStream
     readbegin = fr.ReadLine 
     num4 = readbegin 
     Set fsowrit = CreateObject("Scripting.FileSystem Object")
     Set fw = fsowrit.OpenTextFile("d:\secu\testfile.t xt", ForAppending)
     fw.WriteLine readbegin 
     fw.Close 
  Loop

  Set fsoraz = CreateObject("Scripting.FileSystemObject")
  Set fraz = fsoraz.OpenTextFile("d:\secu\testfiles .txt", ForWriting)
  fraz.WriteLine "" 
  fraz.Close 
  bp = False
  SmartTags("bp") = False

  Set WSHSmessage_code = CreateObject("Wscript.Shell")
  Msgcode = WSHSmessage_code.Popup("Nouveau mot de passe validé." & vbCrLf
  & "En attente validation du service C&C.",5, "Information", 64)
End If
  
If name = Empty And bp = True Then
    Set WSHSmessage_codeerror = CreateObject("Wscript.Shell")
    Msgcodeerror = WSHSmessage_codeerror.Popup("Mauvais Login 
    renseigné.", 5, "Erreur d'éxécution", 48)
 End If
  
If mtp_change <> admin_mtp And bp = True Then
   Set WSHSmessage_codeerror = CreateObject("Wscript.Shell")
   Msgcodeerror = WSHSmessage_codeerror.Popup("Mauvais mtp 
  renseigné.", 5, "Erreur d'éxécution", 48)
End If

mtp_change = ""
admin_mtp = ""
mtp_new = ""
name = ""
cpt = 0
bp = False
StartProgram "D:\secu\bin\fwritev2.exe", "", hmiShowNormal, hmiNo
