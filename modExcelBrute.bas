Attribute VB_Name = "modExcelBrute"

Dim WBs As Object
Dim Ex As Object

Public FileName As String


Public Sub EBInit()

With frmSettings.cdl
    .Filter = "Excel Files (*.xls)|*.xls"
    .ShowOpen
    FileName = .FileName
End With

If FileName = "" Then
MsgBox "No Filename specified.", vbCritical, "Error"
Quit = True
Exit Sub
End If


Set Ex = CreateObject("Excel.Application")

Set WBs = Ex.Workbooks


On Error Resume Next
Dim wb As Object

Set wb = WBs.Open(FileName, , , , "")


If Not (wb Is Nothing) Then
    MsgBox "File is not password-protected!", vbCritical, "Error"
    Quit = True
    Exit Sub
End If

End Sub


Public Sub EBBrute(pass As String)

On Error Resume Next
Dim wb As Object
Set wb = WBs.Open(FileName, , , , pass)

If Not (wb Is Nothing) Then
    MsgBox "Success! " & vbCrLf & "Password is: " & pass, vbInformation, "Excel Password Recoverer"
    Quit = True
End If
End Sub
