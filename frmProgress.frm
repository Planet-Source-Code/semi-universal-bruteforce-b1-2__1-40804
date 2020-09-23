VERSION 5.00
Begin VB.Form frmProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Universal Bruteforce - <progress window>"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picPgs 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   4455
      TabIndex        =   6
      Top             =   480
      Width           =   4455
   End
   Begin VB.PictureBox picPercent 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   990
      ScaleHeight     =   255
      ScaleWidth      =   420
      TabIndex        =   5
      Top             =   960
      Width           =   420
   End
   Begin VB.PictureBox picWordNum 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picCurWord 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      ScaleHeight     =   255
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   2355
      X2              =   2355
      Y1              =   360
      Y2              =   540
   End
   Begin VB.Label lblMid 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblMin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 (0%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblMax 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   4290
      TabIndex        =   1
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Quit = True
End
End Sub

Sub Progress(Percent As Integer)

picPgs.Line (0, 0)-((picPgs.ScaleWidth / 100) * Percent, picPgs.ScaleHeight), &HCCFF, BF


End Sub
