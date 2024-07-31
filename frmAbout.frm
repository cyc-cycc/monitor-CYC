VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于我的应用程序"
   ClientHeight    =   2325
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3285
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   1604.756
   ScaleMode       =   0  'User
   ScaleWidth      =   3084.785
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "支持系统:WindowsXPsp3 ~ Windows11"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "made by CYC"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblDescription 
      Caption         =   "监控应用程序"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblTitle 
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   600
      Width           =   2325
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject outrgn
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    Call rgnform(Me, rou, rou)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sx = X
    sy = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.Left = Me.Left + (X - sx)
        Me.Top = Me.Top + (Y - sy)
    End If
End Sub
