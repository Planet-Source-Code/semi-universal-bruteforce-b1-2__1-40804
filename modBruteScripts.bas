Attribute VB_Name = "modBruteScripts"
Public Scripts As New Collection
Public CurrentScript As Long

Public Sub LoadAllScripts()
Dim f As String

f = Dir(App.Path & "\scripts\*.bs")


Set Scripts = New Collection

Do Until f = ""
Scripts.Add f

If LCase(f) = "default.bs" Then CurrentScript = Scripts.Count


f = Dir()
Loop


LoadScript CurrentScript
End Sub

Public Sub LoadScript(scriptnum As Long)
Dim s As String
Dim initScr As String
Dim mainScr As String

s = File2Str(App.Path & "\scripts\" & Scripts(scriptnum))

initScr = Left(s, InStr(1, LCase(s), "sub brute") - 1)

mainScr = Right(s, Len(s) - Len(initScr))

modBrute.BruteScript = mainScr
modBrute.InitScript = initScr

End Sub

