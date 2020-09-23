VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "universal bruteforce v1.2"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   3915
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   2640
      ScaleHeight     =   735
      ScaleWidth      =   1095
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
      Begin VB.TextBox txtMinLength 
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Text            =   "3"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtMaxLen 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Text            =   "3"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   2175
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
      Begin VB.OptionButton optWordList 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Wordlist file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1845
         Width           =   1695
      End
      Begin VB.OptionButton optGenWords 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Just display"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   1485
         Width           =   1215
      End
      Begin VB.OptionButton optScript 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BruteScript™"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   1305
      End
      Begin VB.OptionButton optExcel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Excel file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   525
         Width           =   945
      End
      Begin VB.OptionButton optWord 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Word file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   885
         Width           =   945
      End
      Begin VB.Image cmdEditScript 
         Height          =   345
         Left            =   1320
         Picture         =   "frmSettings.frx":1272
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.ComboBox txtKeyspace 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmSettings.frx":171E
      Left            =   120
      List            =   "frmSettings.frx":1737
      TabIndex        =   2
      Text            =   "abcdefghijklmnopqrstuvwxyz"
      Top             =   1440
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Unten ausrichten
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   5490
      Width           =   3915
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting for setup..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1365
      End
   End
   Begin MSScriptControlCtl.ScriptControl vbs 
      Left            =   2640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image cmdQuit 
      Height          =   630
      Left            =   2880
      Picture         =   "frmSettings.frx":1866
      Top             =   4560
      Width           =   945
   End
   Begin VB.Image cmdAbout 
      Height          =   630
      Left            =   1200
      Picture         =   "frmSettings.frx":2045
      Top             =   4560
      Width           =   1590
   End
   Begin VB.Image cmdOK 
      Height          =   660
      Left            =   120
      Picture         =   "frmSettings.frx":2A83
      Top             =   4560
      Width           =   870
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   2520
      Picture         =   "frmSettings.frx":3204
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   120
      Picture         =   "frmSettings.frx":387E
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   120
      Picture         =   "frmSettings.frx":3F80
      Top             =   240
      Width           =   2445
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyspace:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   780
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Api As New cAPI

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdEditScript_Click()
frmScript.Show vbModal
End Sub

Private Sub cmdOK_Click()

BruteMode = -1

If optGenWords.Value Then BruteMode = 0
If optScript.Value Then BruteMode = 1
If optWordList.Value Then BruteMode = 2
If optExcel.Value Then BruteMode = 3
If optWord.Value Then BruteMode = 4


If BruteMode = -1 Then
MsgBox "No BruteMode selected!", vbCritical, "Error"
Exit Sub
End If


cmdOK.Enabled = False

If Not (txtKeyspace.Text = "(ALL ASCII)") Then

lblStatus.Caption = "Using Normal Keyspace..."
DoEvents

modBrute.Alphabet = txtKeyspace.Text
Else

lblStatus.Caption = "Generating ASCII Keyspace..."
DoEvents

modBrute.Alphabet = ""

For i = 0 To 255
modBrute.Alphabet = modBrute.Alphabet & Chr(i)
Next i
End If


modBrute.MaxLength = txtMaxLen.Text
modBrute.MinLength = txtMinLength.Text


lblStatus.Caption = "Calculating amount of keys..."
DoEvents

MaxKeys = GetKeyCount
frmProgress.lblMax = MaxKeys & " (100%)"
frmProgress.lblMid = MaxKeys / 2 & " (50%)"

frmProgress.Show
Me.Hide



lblStatus.Caption = "Bruteforceing..."
DoEvents

If BruteMode = 1 Then
vbs.AddCode InitScript & vbCrLf
vbs.Run "Init"
End If

If BruteMode = 2 Then
WordListInit
End If

If BruteMode = 3 Then
modExcelBrute.EBInit
End If

If BruteMode = 4 Then
modWordBrute.WBInit
End If



modBrute.GetSubKeys ""

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub txtHwnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    frmSelectWin.Show vbModal
    txthWnd.Text = ReturnValue
End If

End Sub


Private Sub txtOKhWnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then


    frmSelectWin.Show vbModal
    txtOKhWnd.Text = ReturnValue
End If
End Sub

Private Sub Form_Load()
vbs.AddObject "BF", Api

LoadAllScripts

'BruteScript = File2Str(App.Path & "\brute.txt")
'InitScript = File2Str(App.Path & "\init.txt")

End Sub



