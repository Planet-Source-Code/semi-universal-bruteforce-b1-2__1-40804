Attribute VB_Name = "modWordBrute"

Dim Wrd As Object
Dim Docs As Object



Public Sub WBInit()
With frmSettings.cdl
    .Filter = "Word Files (*.doc)|*.doc"
    .ShowOpen
    FileName = .FileName
End With

If FileName = "" Then
MsgBox "No Filename specified.", vbCritical, "Error"
Quit = True
Exit Sub
End If

Set Wrd = CreateObject("Word.Application")
Set Docs = Wrd.Documents


On Error Resume Next
Dim doc As Object

Set doc = Docs.Open(FileName, , , , " ")

If Not (doc Is Nothing) Then
    MsgBox "File is not password-protected!", vbCritical, "Error"
    Quit = True
    Exit Sub
End If


End Sub



Public Sub WBBrute(pass As String)
On Error Resume Next
Dim doc As Object
Set doc = Docs.Open(FileName, , , , pass)


If Not (doc Is Nothing) Then
    MsgBox "Success! " & vbCrLf & "Password is: " & pass, vbInformation, "Word Password Recoverer"
    Quit = True
End If

End Sub
