Attribute VB_Name = "modBrowser"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
   
   Public Sub HyperJump(ByVal URL As String)
      Call ShellExecute(0&, vbNullString, URL, vbNullString, _
                        vbNullString, vbNormalFocus)
   End Sub


