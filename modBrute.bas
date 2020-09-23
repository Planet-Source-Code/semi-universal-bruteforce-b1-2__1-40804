Attribute VB_Name = "modBrute"

Public Alphabet As String
Public MaxLength As Long
Public MinLength As Long

Public InitScript As String
Public BruteScript As String

Public BruteMode As Long

Public MaxKeys As Long
Public CurrentKeys As Long

Public Sub GetSubKeys(SKeysFor As String)
    If GetKeyState(VK_END) < -50 Then Quit = True
    
    If Len(SKeysFor) >= MaxLength Then Exit Sub

   

    For i = 1 To Len(Alphabet)
    
        If Quit Then End
                 
        If Len(SKeysFor) + 1 >= MinLength Then KeyFound SKeysFor & Mid(Alphabet, i, 1)
            
        GetSubKeys (SKeysFor & Mid(Alphabet, i, 1))
    Next i
    

End Sub



Public Sub KeyFound(Key As String)
    CurrentKeys = CurrentKeys + 1
    frmProgress.Show
    frmProgress.Progress ((CurrentKeys / MaxKeys) * 100)

    frmProgress.picWordNum.Cls
    frmProgress.picWordNum.Print CStr(CurrentKeys)
    frmProgress.picCurWord.Cls
    frmProgress.picCurWord.Print Key
    frmProgress.picPercent.Cls
    frmProgress.picPercent.Print Int((CurrentKeys / MaxKeys) * 100) & "%"
    
    DoEvents


    Select Case BruteMode
        Case 0 ' Just output
        
        Case 1 ' BruteScript
        
            frmSettings.Api.Key = Key
            frmSettings.vbs.AddCode vbCrLf & BruteScript & vbCrLf
            frmSettings.vbs.Run "Brute"
        Case 2 ' BruteFile
            Print #1, Key
            
            
        Case 3 'Excel
            modExcelBrute.EBBrute Key
        Case 4 'Word
            modWordBrute.WBBrute Key
    End Select

End Sub


Public Function GetKeyCount() As Long

If MaxLength = MinLength Then
 GetKeyCount = Len(Alphabet) ^ MaxLength
Else

    Dim kc As Long
    For i = MinLength To MaxLength
        kc = kc + Len(Alphabet) ^ i
    Next i
    GetKeyCount = kc


End If

End Function
