VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "about..."
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Image imgSC 
      Height          =   465
      Left            =   180
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.2"
      Height          =   195
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   495
      Left            =   3960
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgSC_Click()
HyperJump "http://www.semicolonsoftware.de/"
End Sub
