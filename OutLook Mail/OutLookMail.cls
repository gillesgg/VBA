VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutLookMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOutLookApp As Object
Private mOutMail As Object

Private mBody As String
Private mHTMLBody As String
Private mSubJect As String
Private mSendTo As String
Private mSendToCC As String
Private mSendToBCC As String
Private mAFile As Collection

Private Sub Class_Initialize()
    On Error GoTo Err:
        Set mOutLookApp = CreateObject("Outlook.Application")
        Set mOutMail = mOutLookApp.CreateItem(0)
        Set mAFile = New Collection
        Exit Sub
Err:
    Debug.Print Err.Description
End Sub
Private Sub Class_Terminate()
    Set mOutMail = Nothing
    Set mOutLookApp = Nothing
End Sub
Public Property Let SendToCC(value As String)
    mSendToCC = value
End Property
Public Property Get SendToCC() As String
    SendToCC = mSendToBCC
End Property

Public Property Get SendToBCC() As String
    SendToBCC = mSendToBCC
End Property
Public Property Let SendToBCC(value As String)
    mSendToBCC = value
End Property

Public Property Get HTMLBody() As String
    HTMLBody = mHTMLBody
End Property
Public Property Let HTMLBody(value As String)
    mHTMLBody = value
End Property
Public Property Get Body() As String
    Body = mBody
End Property
Public Property Let Body(value As String)
    mBody = value
End Property
Public Property Get SubJect() As String
    SubJect = mSubJect
End Property
Public Property Let SubJect(value As String)
    mSubJect = value
End Property
Public Property Get SendTo() As String
    SendTo = mSendTo
End Property
Public Property Let SendTo(value As String)
    mSendTo = value
End Property

Public Property Let AddFile(value As String)

        If (fileExists(value)) Then
            If (mAFile.Count > 0) Then
                If (CollectionContains(mAFile, value)) Then
                    Exit Property
                End If
            End If
            
            mAFile.Add value, value
        Else
            Err.Raise 1024, "VBA", "The file : " & value & " do not exist"
        End If
    
End Property
Private Sub sendToOutLook()
    Dim item As Variant
    On Error GoTo Err:
    
    If AsProfile = False Then
            Err.Raise 1024, "VBA", "unable to send the email, unable to detect the profile"
    End If

    If mSubJect = "" Then
        Err.Raise 1024, "VBA", "unable to send the email, the subject is empty"
    End If

    If mBody = "" And mHTMLBody = "" Then
        Err.Raise 1024, "VBA", "unable to send the email, the body is empty"
    End If

    If mSendTo = "" Then
        Err.Raise 1024, "VBA", "the email subject is empty"
    End If

    If Not mOutLookApp Is Nothing And Not mOutMail Is Nothing Then
        With mOutMail
            .To = SendTo
            .CC = SendToCC
            .BCC = SendToBCC
            .SubJect = SubJect
            If (HTMLBody <> "") Then
                .HTMLBody = HTMLBody
            ElseIf Body <> "" Then
                .Body = Body
            End If
            
            For Each item In mAFile
            .Attachments.Add item
            Next
            .Send
        End With
    Else
        Err.Raise 1024, "VBA", "The outlook object is not valid"
    End If
    Exit Sub
Err:
     Err.Raise Err.Number, "VBA", Err.Description
End Sub
Private Sub sendToNotes()
    Debug.Print "sendToNotes"
End Sub

Public Sub Send()
    If AsProfile() = True Then
        sendToOutLook
    Else
        sendToNotes
    End If
End Sub

Private Function fileExists(fileName As String) As Boolean
    Dim obj_fso As Object
        Set obj_fso = CreateObject("Scripting.FileSystemObject")
        fileExists = obj_fso.fileExists(fileName)
End Function

Function CollectionContains(myCol As Collection, checkVal As Variant) As Boolean
    On Error Resume Next
    CollectionContains = False
    Dim it As Variant
    For Each it In myCol
        If it = checkVal Then
            CollectionContains = True
            Exit Function
        End If
    Next
End Function
'HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles --> Office 2010
'Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook --> Office 2016
Public Function AsProfile() As Boolean
Dim bExist As Boolean
    bExist = RegKeyExists("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\DefaultProfile")
    If bExist = True Then
        AsProfile = True
        Exit Function
    End If
     bExist = RegKeyExists("HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\DefaultProfile")
    If bExist = True Then
        AsProfile = True
        Exit Function
    End If
AsProfile = False
End Function

Private Function RegKeyExists(i_RegKey As String) As Boolean

Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  'key was not found
  RegKeyExists = False
End Function


Public Function IsWindows10() As Boolean
Dim numver As Integer

    numver = CInt(Application.System.Version)

If (numver >= 10) Then
    IsWindows10 = True
Else
   IsWindows10 = False
End If
End Function


