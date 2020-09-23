VERSION 5.00
Begin VB.Form frmSelectWin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Universal Bruteforce - <Select Window>"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Trg 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "frmSelectWin.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtTitle 
         Appearance      =   0  '2D
         BackColor       =   &H00404040&
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   30
         TabIndex        =   2
         Text            =   "(window title appears here)"
         Top             =   255
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drag the crosshair over the window you want to use."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   60
         TabIndex        =   3
         Top             =   15
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   4335
      TabIndex        =   4
      Top             =   600
      Width           =   4335
      Begin VB.TextBox txtDlgItemID 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DlgItemID:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSelectWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_bDragging As Boolean

Private Sub Trg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And Not m_bDragging Then

        m_bDragging = True

    End If
End Sub

Private Sub Trg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    If Button = vbLeftButton And m_bDragging Then
        Dim tPA As POINTAPI
        Dim lhWnd As Long
        GetCursorPos tPA
        
        lhWnd = WindowFromPoint(tPA.X, tPA.Y)
        SetFocusAPI lhWnd
        modGlobal.ReturnValue = lhWnd
        
        Dim wt As String
        
        wt = String(GetWindowTextLength(lhWnd), Chr(0))
        
        GetWindowText lhWnd, wt, 255
        
        txtTitle.Text = wt
        txtDlgItemID.Text = GetDlgCtrlID(lhWnd)
    End If


End Sub


Private Sub Trg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And m_bDragging Then
        m_bDragging = False
        Me.MousePointer = vbNormal
        
        Unload Me
        
    End If
End Sub
