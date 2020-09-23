VERSION 5.00
Begin VB.Form frmCrackme 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Universal Bruteforce crackme"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4455
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdHint 
      Caption         =   "Hint"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Give Up"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Try"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmCrackme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHint_Click()
MsgBox "The Code is a long integer without dividers. Now fire up Universal BruteForce and have fun!"
End Sub

Private Sub cmdOK_Click()
Dim c As Long

c = 12
For i = 1 To Len(txtName.Text)
c = c + Asc(txtName.Text) + i
Next i


If Int("0" & txtCode.Text) = c Then
    MsgBox "Code OK!", vbInformation, "Congratulations!"
Else
    MsgBox "Invalid Code!", vbCritical, "error"
End If


End Sub

Private Sub cmdQuit_Click()

End
End Sub

Private Sub Command1_Click()

End Sub
