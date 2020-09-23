Attribute VB_Name = "Main"
Option Explicit

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Sub formdrag(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hWnd, &HA1, 2, 0&)
End Sub

Public Function RoundedForm()
    Dim hrgn As Long
    hrgn = CreateRoundRectRgn(0, 0, frmSample.ScaleX(frmSample.Width, vbTwips, vbPixels), frmSample.ScaleY(frmSample.Height, vbTwips, vbPixels), 50, 50)
    SetWindowRgn frmSample.hWnd, hrgn, True
    DeleteObject hrgn
End Function

Public Function Convert(cString As String) As String
    Dim Ccode As Integer
    For Ccode = 1 To Len(cString)
        Convert = Convert + Chr(255 - Asc(Mid(cString, CInt(Ccode), 1)))
    Next Ccode
End Function
