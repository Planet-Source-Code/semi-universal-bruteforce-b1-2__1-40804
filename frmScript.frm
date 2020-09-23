VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmScript 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Universal BruteForce - <BruteScript>"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   8310
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8310
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture2 
      Align           =   1  'Oben ausrichten
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'Kein
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   8310
      TabIndex        =   6
      Top             =   510
      Width           =   8310
   End
   Begin VB.ComboBox cboScripts 
      Height          =   315
      ItemData        =   "frmScript.frx":0442
      Left            =   120
      List            =   "frmScript.frx":0444
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   330
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Oben ausrichten
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   8310
      TabIndex        =   3
      Top             =   0
      Width           =   8310
   End
   Begin TabDlg.SSTab sst 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Initialization"
      TabPicture(0)   =   "frmScript.frx":0446
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtInitScript"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bruteforce Step"
      TabPicture(1)   =   "frmScript.frx":0462
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBruteStep"
      Tab(1).ControlCount=   1
      Begin CodeSenseCtl.CodeSense txtBruteStep 
         Height          =   4815
         Left            =   -74880
         OleObjectBlob   =   "frmScript.frx":047E
         TabIndex        =   1
         Top             =   360
         Width           =   7815
      End
      Begin CodeSenseCtl.CodeSense txtInitScript 
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "frmScript.frx":05E4
         TabIndex        =   2
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuhWnd 
         Caption         =   "Get hWnd / DlgCtrlId"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhWnd_Click()

End Sub

Private Sub cmdOK_Click()
BruteScript = txtBruteStep.Text
InitScript = txtInitScript.Text
Unload Me
End Sub




Private Sub cboScripts_Click()
CurrentScript = cboScripts.ListIndex + 1
LoadScript CurrentScript

txtBruteStep.Text = BruteScript
txtInitScript.Text = InitScript

End Sub

Private Sub cmdSave_Click()
Str2File App.Path & "\scripts\" & Scripts(CurrentScript), txtInitScript.Text & txtBruteStep.Text


End Sub

Private Sub Form_Load()
txtBruteStep.Text = BruteScript
txtInitScript.Text = InitScript


txtBruteStep.Language = "Basic"
txtBruteStep.DisplayLeftMargin = False
txtBruteStep.DisplayWhitespace = False

txtInitScript.Language = "Basic"
txtInitScript.DisplayLeftMargin = False
txtInitScript.DisplayWhitespace = False


For Each s In Scripts
cboScripts.AddItem s
Next

cboScripts.ListIndex = CurrentScript - 1

End Sub

Private Sub Form_Resize()
On Error Resume Next
sst.Width = ScaleWidth - 2 * sst.Left
sst.Height = ScaleHeight - sst.Top - sst.Left

txtInitScript.Width = sst.Width - txtInitScript.Left * 2
txtInitScript.Height = sst.Height - txtInitScript.Top - txtInitScript.Left
txtBruteStep.Width = txtInitScript.Width
txtBruteStep.Height = txtInitScript.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
BruteScript = txtBruteStep.Text
InitScript = txtInitScript.Text
Unload Me
End Sub

Private Sub mnuClose_Click()
BruteScript = txtBruteStep.Text
InitScript = txtInitScript.Text
Unload Me
End Sub

Private Sub mnuhWnd_Click()
frmSelectWin.Show vbModal
Clipboard.Clear
Clipboard.SetText ReturnValue
MsgBox "You hWnd is: " & ReturnValue & "." & vbCr & "It has been copied to clipboard.", vbInformation

End Sub

