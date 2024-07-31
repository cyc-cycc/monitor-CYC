Attribute VB_Name = "º¯ÊýÉùÃ÷"
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long)
    Dim w As Long, h As Long
    w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels) - 2
    h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels) - 2
    outrgn = CreateRoundRectRgn(3, 25, w, h, fw, fh)
    Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub
