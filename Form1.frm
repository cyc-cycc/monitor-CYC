VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   3900
   ClientTop       =   1890
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "��س���"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   4935
   ScaleWidth      =   4455
   Begin VB.CommandButton end7 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton refresh7 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox ss 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton end6 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton refresh6 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox cxjk 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Text            =   "20"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton end5 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton refresh5 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox gbcs 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   855
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ӧ��"
      Height          =   855
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Text            =   "210"
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton refresh3 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton refresh4 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton refresh2 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton refresh1 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2640
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton end4 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox spcz 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton enda 
      Caption         =   "ȫ������"
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton exit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   22
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton refresh 
      Caption         =   "����ˢ��"
      Height          =   375
      Left            =   1200
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton end3 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton end2 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton end1 
      Caption         =   "��������"
      Height          =   375
      Left            =   3360
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton gy 
      Caption         =   "����"
      Height          =   375
      Left            =   2280
      MousePointer    =   1  'Arrow
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox xye 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox spgb 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   1200
      Top             =   600
   End
   Begin VB.TextBox spjq 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "������"
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1200
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   5000
      Left            =   1200
      Top             =   1080
   End
   Begin VB.Timer Timer6 
      Interval        =   5000
      Left            =   1200
      Top             =   3000
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   1200
      Top             =   2520
   End
   Begin VB.Timer Timer7 
      Interval        =   5000
      Left            =   1200
      Top             =   1560
   End
   Begin VB.Label Label9 
      Caption         =   "���ٲ�������"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "�����أ�"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   34
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Բ�Ǵ�С��"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   33
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "�����ԣ�"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   31
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "��͸����(0~225)��"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "ˢ������_���ƣ�"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   29
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "���˶죺"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   27
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ˢ������_ؤ�棺"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   25
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ˢ��������"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Function CheckApplicationIsRun(ByVal szExeFileName As String) As Boolean
    On Error GoTo Err
    Dim WMI
    Dim Obj
    Dim Objs
    CheckApplicationIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
        If InStr(UCase(szExeFileName), UCase(Obj.Description)) <> 0 Then
            CheckApplicationIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
        End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function
Err:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function
Private Sub Command1_Click()
    aero = Me.Text1.Text
    rou = Me.Text2.Text
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    Call rgnform(Me, rou, rou)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "settings", "left", Me.Left
    SaveSetting App.Title, "settings", "top", Me.Top
    SaveSetting App.Title, "settings", "aero", Me.Text1.Text
    SaveSetting App.Title, "settings", "round", Me.Text2.Text
    DeleteObject outrgn
    End
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long)
    Dim w As Long, h As Long
    w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels) - 1
    h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels) - 1
    outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
    Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub
Private Sub Command2_Click()
    Me.Text1.Text = 210
    Me.Text2.Text = 20
    aero = Me.Text1.Text
    rou = Me.Text2.Text
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    Call rgnform(Me, rou, rou)
End Sub
Private Sub end1_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ˢ������.exe"
End Sub
Private Sub end2_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ˢ������_ؤ��.exe"
End Sub
Private Sub end3_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ���˶�.exe"
End Sub
Private Sub end4_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ˢ������_����.exe"
End Sub
Private Sub end5_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ������.exe"
End Sub
Private Sub end6_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ��س���.exe"
End Sub
Private Sub end7_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ���ٲ�����.exe"
End Sub
Private Sub enda_Click()
    On Error GoTo 0
    Shell "cmd /c taskkill /f /im ˢ������.exe"
    Shell "cmd /c taskkill /f /im ˢ������_ؤ��.exe"
    Shell "cmd /c taskkill /f /im ���˶�.exe"
    Shell "cmd /c taskkill /f /im ˢ������_����.exe"
    Shell "cmd /c taskkill /f /im ������.exe"
    Shell "cmd /c taskkill /f /im ���ٲ�����.exe"
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    aero = GetSetting(App.Title, "settings", "aero", 210)
    rou = GetSetting(App.Title, "settings", "round", 20)
    Me.Text1.Text = aero
    Me.Text2.Text = rou
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    Call rgnform(Me, rou, rou)
    Me.Left = GetSetting(App.Title, "settings", "left", 0)
    Me.Top = GetSetting(App.Title, "settings", "top", 0)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sx = X
    sy = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        main.Left = main.Left + (X - sx)
        main.Top = main.Top + (Y - sy)
    End If
End Sub
Private Sub gy_Click()
    frmAbout.Show
End Sub
Private Sub refresh_Click()
    Me.spjq.Text = "������"
    Me.spgb.Text = "������"
    Me.spcz.Text = "������"
    Me.xye.Text = "������"
    Me.gbcs.Text = "������"
    Me.ss.Text = "������"
    If CheckApplicationIsRun("ˢ������.exe") = True Then
        Me.spjq.Text = "������"
    Else
        Me.spjq.Text = "���˳�"
    End If
    If CheckApplicationIsRun("ˢ������_ؤ��.exe") = True Then
        Me.spgb.Text = "������"
    Else
        Me.spgb.Text = "���˳�"
    End If
    If CheckApplicationIsRun("ˢ������_����.exe") = True Then
        Me.spcz.Text = "������"
    Else
        Me.spcz.Text = "���˳�"
    End If
    If CheckApplicationIsRun("���ٲ�����.exe") = True Then
        Me.ss.Text = "������"
    Else
        Me.ss.Text = "���˳�"
    End If
    If CheckApplicationIsRun("���˶�.exe") = True Then
        Me.xye.Text = "������"
    Else
        Me.xye.Text = "���˳�"
    End If
    If CheckApplicationIsRun("������.exe") = True Then
        Me.gbcs.Text = "������"
    Else
        Me.gbcs.Text = "���˳�"
    End If
    If CheckApplicationIsRun("��س���.exe") = True Then
        Me.cxjk.Text = "������"
    Else
        Me.cxjk.Text = "���˳�"
    End If
End Sub
Private Sub refresh1_Click()
    If CheckApplicationIsRun("ˢ������.exe") = True Then
        Me.spjq.Text = "������"
    Else
        Me.spjq.Text = "���˳�"
    End If
End Sub
Private Sub refresh2_Click()
    If CheckApplicationIsRun("ˢ������_ؤ��.exe") = True Then
        Me.spgb.Text = "������"
    Else
        Me.spgb.Text = "���˳�"
    End If
End Sub
Private Sub refresh3_Click()
    If CheckApplicationIsRun("���˶�.exe") = True Then
        Me.xye.Text = "������"
    Else
        Me.xye.Text = "���˳�"
    End If
End Sub
Private Sub refresh4_Click()
    If CheckApplicationIsRun("ˢ������_����.exe") = True Then
        Me.spcz.Text = "������"
    Else
        Me.spcz.Text = "���˳�"
    End If
End Sub
Private Sub refresh5_Click()
    If CheckApplicationIsRun("������.exe") = True Then
        Me.gbcs.Text = "������"
    Else
        Me.gbcs.Text = "���˳�"
    End If
End Sub
Private Sub refresh6_Click()
    If CheckApplicationIsRun("��س���.exe") = True Then
        Me.cxjk.Text = "������"
    Else
        Me.cxjk.Text = "���˳�"
    End If
End Sub
Private Sub refresh7_Click()
    If CheckApplicationIsRun("���ٲ�����.exe") = True Then
        Me.ss.Text = "������"
    Else
        Me.ss.Text = "���˳�"
    End If
End Sub
Private Sub Timer1_Timer()
    If CheckApplicationIsRun("ˢ������.exe") = True Then
        Me.spjq.Text = "������"
    Else
        Me.spjq.Text = "���˳�"
    End If
End Sub
Private Sub Timer2_Timer()
    If CheckApplicationIsRun("ˢ������_ؤ��.exe") = True Then
        Me.spgb.Text = "������"
    Else
        Me.spgb.Text = "���˳�"
    End If
End Sub
Private Sub Timer3_Timer()
    If CheckApplicationIsRun("���˶�.exe") = True Then
        Me.xye.Text = "������"
    Else
        Me.xye.Text = "���˳�"
    End If
End Sub
Private Sub Timer4_Timer()
    If CheckApplicationIsRun("ˢ������_����.exe") = True Then
        Me.spcz.Text = "������"
    Else
        Me.spcz.Text = "���˳�"
    End If
End Sub
Private Sub Timer5_Timer()
    If CheckApplicationIsRun("������.exe") = True Then
        Me.gbcs.Text = "������"
    Else
        Me.gbcs.Text = "���˳�"
    End If
End Sub
Private Sub Timer6_Timer()
    If CheckApplicationIsRun("��س���.exe") = True Then
        Me.cxjk.Text = "������"
    Else
        Me.cxjk.Text = "���˳�"
    End If
End Sub
Private Sub Timer7_Timer()
    If CheckApplicationIsRun("���ٲ�����.exe") = True Then
        Me.ss.Text = "������"
    Else
        Me.ss.Text = "���˳�"
    End If
End Sub
