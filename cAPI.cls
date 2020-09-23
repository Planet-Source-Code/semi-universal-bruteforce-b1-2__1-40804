VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Public Key As String

Public Sub aDoevents()
DoEvents
End Sub

Public Sub aSleep(time)
    Sleep time
End Sub

Public Function aFindWindow(WindowName) As Long
    aFindWindow = FindWindow(vbNullString, WindowName)
End Function

Public Function aSetFocus(hwnd) As Long
    aSetFocus = SetFocusAPI(hwnd)
End Function

Public Sub aSetWindowText(iHWnd, txt, Optional isDlgItem = False)
    If isDlgItem Then
        prnt = GetParent(iHWnd)
        ctlid = GetDlgCtrlID(iHWnd)
        
        'r = SetDlgItemText(prnt, ctlid, "Hallo")

        Dim t As String
        t = txt
        
        
        r = SendMessage(iHWnd, WM_SETTEXT, 0, ByVal t)

        
        Else
        r = SetWindowText(iHWnd, txt)
    End If
    

End Sub



Public Function aGetWindowText(hwnd) As String
    Dim bfr As String
    
    
    
    bfr = String(GetWindowTextLength(hwnd), Chr(0))
    
    GetWindowText hwnd, bfr, 255
    
    aGetWindowText = bfr
End Function

Public Sub aClick(hwnd)
    modMClick.CenterMouseOn hwnd
    modMClick.MouseFullClick btcLeft

End Sub


Public Function aGetDlgItem(hDlg, nIDDlgItem)
    aGetDlgItem = GetDlgItem(hDlg, nIDDlgItem)
End Function

