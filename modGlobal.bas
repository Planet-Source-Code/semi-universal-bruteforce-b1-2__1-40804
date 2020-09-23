Attribute VB_Name = "modGlobal"
Public Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetDlgItemInt Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
Public Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_SPACE = &H20
Public Const WM_NOTIFY = &H4E
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_ACTIVATE = &H6

Public Const WM_MENUSELECT = &H11F
Public Const WM_SETTEXT = &HC




Public Type POINTAPI
    X As Long
    Y As Long

End Type

Public ReturnValue As Long
Public ReturnOK As Boolean


Public Quit As Boolean
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Const VK_END = &H23


Public Function File2Str(file As String) As String
Dim s As String

Open file For Binary Access Read As 3
    s = String(LOF(3), Chr(0))
    Get #3, , s
Close 3

File2Str = s
End Function

Public Sub Str2File(file As String, strng As String)
Kill file
Open file For Binary Access Write As 2
    Put #2, , strng
Close 2

End Sub

Public Sub WordListInit()
With frmSettings.cdl
    .Filter = "Text files (*.txt)|*.txt"
    .ShowSave


    If .FileName = "" Then
        MsgBox "No Filename specified!", vbCritical, "Error"
        Quit = True
    End If
    
    Open .FileName For Output As 1

End With

End Sub
