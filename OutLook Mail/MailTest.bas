Attribute VB_Name = "Module1"
Option Explicit

Sub MailSimple()
Dim outlook As New OutLookMail
On Error GoTo Err
    outlook.SendTo = "email"
    outlook.Body = "Texte du mail"
    outlook.SubJect = "Sujet"
    outlook.Send
    Exit Sub
Err:
Debug.Print Err.Description

End Sub

Sub MailSimpleAvecFichier()
Dim outlook As New OutLookMail
On Error GoTo Err
    outlook.SendTo = "email"
    outlook.Body = "Texte du mail"
    outlook.SubJect = "Sujet"
    outlook.AddFile = "c:\temp\Accents.txt"
    outlook.Send
    Exit Sub
Err:
Debug.Print Err.Description

End Sub

Sub MailHTMLAvecFichier()
Dim outlook As New OutLookMail
Dim strHTML As String
Dim strText As String
On Error GoTo Err

    outlook.SendTo = "email"
    outlook.SubJect = "Sujet"
    strText = "Texte du mail"
    outlook.AddFile = "c:\temp\Accents.txt"
    strHTML = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & _
          "<HTML><HEAD>" & _
          "<META http-equiv=Content-Type content=""text/html; charset=iso-8859-1"">" & _
          "<META content=""MSHTML 6.00.2800.1516"" name=GENERATOR></HEAD>" & _
          "<BODY><DIV STYLE=""font-size: 25px; font-face: Tahoma;"">"
 
    outlook.HTMLBody = strHTML & Replace(strText, vbCrLf, "<br>") & "</DIV></BODY></HTML>"
    outlook.Send
    Exit Sub
Err:
Debug.Print Err.Description

End Sub

