Attribute VB_Name = "modMClick"
Option Explicit

' ----------------------------------------------
' *        MouseEvent Related Declares         *
' ----------------------------------------------
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, _
    ByVal dwExtraInfo As Long)

' ----------------------------------------------
' *     GetSystemMetrics Related Declares      *
' ----------------------------------------------
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const TWIPS_PER_INCH = 1440
Private Const POINTS_PER_INCH = 72
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex _
    As Long) As Long

' ----------------------------------------------
' *       GetWindowRect Related Declares       *
' ----------------------------------------------
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long


' ----------------------------------------------
' *       Internal Constants and Types         *
' ----------------------------------------------

Private Const MOUSE_MICKEYS = 65535

Public Enum enReportStyle
    rsPixels
    rsTwips
    rsInches
    rsPoints
End Enum

Public Enum enButtonToClick
    btcLeft
    btcRight
    btcMiddle
End Enum
'--------------------------------------------------------------------------------

' Returns the screen size in pixels or, optionally,
' in others scalemode styles
Public Sub GetScreenRes(ByRef X As Long, ByRef Y As Long, Optional ByVal _
    ReportStyle As enReportStyle)

    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    If Not IsMissing(ReportStyle) Then
         If ReportStyle <> rsPixels Then
            X = X * Screen.TwipsPerPixelX
            Y = Y * Screen.TwipsPerPixelY
            If ReportStyle = rsInches Or ReportStyle = rsPoints Then
                X = X \ TWIPS_PER_INCH
                Y = Y \ TWIPS_PER_INCH
                If ReportStyle = rsPoints Then
                    X = X * POINTS_PER_INCH
                    Y = Y * POINTS_PER_INCH
                End If
            End If
        End If
    End If
End Sub
'--------------------------------------------------------------------------------

' Convert's the mouses coordinate system to
' a pixel position.
Public Function MickeyXToPixel(ByVal mouseX As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tX As Single
    Dim tmouseX As Single
    Dim tMickeys As Single
    
    GetScreenRes X, Y
    tX = X
    tMickeys = MOUSE_MICKEYS
    tmouseX = mouseX
    
    MickeyXToPixel = CLng(tmouseX / (tMickeys / tX))
    
End Function
'--------------------------------------------------------------------------------

' Converts mouse Y coordinates to pixels
Public Function MickeyYToPixel(ByVal mouseY As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tY As Single
    Dim tmouseY As Single
    Dim tMickeys As Single
    
    GetScreenRes X, Y
    tY = Y
    tMickeys = MOUSE_MICKEYS
    tmouseY = mouseY
    
    MickeyYToPixel = CLng(tmouseY / (tMickeys / tY))
    
End Function
'--------------------------------------------------------------------------------

' Converts pixel X coordinates to mickeys
Public Function PixelXToMickey(ByVal pixX As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tX As Single
    Dim tpixX As Single
    Dim tMickeys As Single
    
    GetScreenRes X, Y
    tMickeys = MOUSE_MICKEYS
    tX = X
    tpixX = pixX
    
    PixelXToMickey = CLng((tMickeys / tX) * tpixX)

End Function
'--------------------------------------------------------------------------------

' Converts pixel Y coordinates to mickeys
Public Function PixelYToMickey(ByVal pixY As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tY As Single
    Dim tpixY As Single
    Dim tMickeys As Single
    
    GetScreenRes X, Y
    tMickeys = MOUSE_MICKEYS
    tY = Y
    tpixY = pixY
    
    PixelYToMickey = CLng((tMickeys / tY) * tpixY)

End Function
'--------------------------------------------------------------------------------

' The function will center the mouse on a window
' or control with an hWnd property.  No checking
' is done to ensure that the window is not obscured
' or not minimized, however it does make sure that
' the target is within the boundaries of the
' screen.
Public Function CenterMouseOn(ByVal hwnd As Long) As Boolean
    Dim X As Long
    Dim Y As Long
    Dim maxX As Long
    Dim maxY As Long
    Dim crect As RECT
    Dim rc As Long

    GetScreenRes maxX, maxY
    rc = GetWindowRect(hwnd, crect)
    
    If rc Then
        X = crect.Left + ((crect.Right - crect.Left) / 2)
        Y = crect.Top + ((crect.Bottom - crect.Top) / 2)
        If (X >= 0 And X <= maxX) And (Y >= 0 And Y <= maxY) Then
            MouseMove X, Y
            CenterMouseOn = True
        Else
            CenterMouseOn = False
        End If
    Else
        CenterMouseOn = False
    End If
End Function
'--------------------------------------------------------------------------------

' Simulates a mouse click
Public Function MouseFullClick(ByVal MBClick As enButtonToClick) As Boolean
    Dim cbuttons As Long
    Dim dwExtraInfo As Long
    Dim mevent As Long
    
    Select Case MBClick
        Case btcLeft
            mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP
        Case btcRight
            mevent = MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP
        Case btcMiddle
            mevent = MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP
        Case Else
            MouseFullClick = False
            Exit Function
    End Select
    mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo
    MouseFullClick = True
    
End Function
'--------------------------------------------------------------------------------

Public Sub MouseMove(ByRef xPixel As Long, ByRef yPixel As Long)
    Dim cbuttons As Long
    Dim dwExtraInfo As Long
    
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, _
        PixelXToMickey(xPixel), PixelYToMickey(yPixel), cbuttons, dwExtraInfo

End Sub

