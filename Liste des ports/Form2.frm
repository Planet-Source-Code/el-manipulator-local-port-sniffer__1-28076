VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4395
      Top             =   390
   End
   Begin VB.Label Label16 
      Caption         =   "10"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2730
      Width           =   1605
   End
   Begin VB.Label Label15 
      Caption         =   "[Remote IP]"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "[Local Port]"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "[State]"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Don't care"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3615
      TabIndex        =   11
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Don't care"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3630
      TabIndex        =   10
      Top             =   1245
      Width           =   1185
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Learn [BAD]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2250
      TabIndex        =   9
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Learn [BAD]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2265
      TabIndex        =   8
      Top             =   1245
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Learn [OK]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   885
      TabIndex        =   7
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Learn [OK]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   900
      TabIndex        =   6
      Top             =   1245
      Width           =   1185
   End
   Begin VB.Shape Shape8 
      Height          =   300
      Left            =   3600
      Top             =   1185
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Shape Shape7 
      Height          =   300
      Left            =   2235
      Top             =   1185
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Shape Shape6 
      Height          =   300
      Left            =   870
      Top             =   1185
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   4815
      X2              =   4815
      Y1              =   1200
      Y2              =   1470
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   3615
      X2              =   4830
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00755433&
      X1              =   3450
      X2              =   3450
      Y1              =   1200
      Y2              =   1470
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00755433&
      X1              =   2250
      X2              =   3465
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   2085
      X2              =   2085
      Y1              =   1200
      Y2              =   1470
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   885
      X2              =   2100
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   885
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   2250
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   3615
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "#2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   900
      TabIndex        =   5
      Top             =   690
      Width           =   3945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "#2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   915
      TabIndex        =   4
      Top             =   705
      Width           =   3945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   885
      TabIndex        =   3
      Top             =   480
      Width           =   3945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   900
      TabIndex        =   2
      Top             =   495
      Width           =   3945
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   150
      Picture         =   "Form2.frx":0000
      Top             =   405
      Width           =   510
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   4920
      X2              =   4920
      Y1              =   15
      Y2              =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4935
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   0
      Top             =   255
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Notification!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   15
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Notification!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   -930
      Picture         =   "Form2.frx":090A
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 1
 X_Initial = X
 Y_Initial = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + X - X_Initial
  Me.Top = Me.Top + Y - Y_Initial
 Else
  Call remap
 End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 0
 Dist_Am = 100
 
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape4.BorderColor
 Line7.BorderColor = Shape4.BorderColor
 Shape4.BorderColor = Value
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape7.Visible = False Then
  Call remap
  Shape7.Visible = True
 End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape4.BorderColor
 Line7.BorderColor = Shape4.BorderColor
 Shape4.BorderColor = Value
 
 Value = Line4.BorderColor
 Line4.BorderColor = Shape5.BorderColor
 Line5.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
 
 If Label13.Caption = "LISTEN" Then
  TrackedMode(Trackeds) = -1
  TrackedLocalPort(Trackeds) = Label14.Caption
  TrackedRemoteIP$(Trackeds) = "*"
  Trackeds = Trackeds + 1
 End If
 If Label13.Caption = "ESTAB" Then
  TrackedMode(Trackeds) = -1
  TrackedLocalPort(Trackeds) = Label14.Caption
  TrackedRemoteIP$(Trackeds) = Label15.Caption
  Trackeds = Trackeds + 1
 End If
 
 ReadyToShow = True
 Unload Me
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape3.BorderColor
 Line9.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape8.Visible = False Then
  Call remap
  Shape8.Visible = True
 End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape3.BorderColor
 Line9.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value

 ReadyToShow = True
 Unload Me
End Sub

Sub remap()
 If Shape6.Visible = True Then Shape6.Visible = False
 If Shape7.Visible = True Then Shape7.Visible = False
 If Shape8.Visible = True Then Shape8.Visible = False
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape5.BorderColor
 Line5.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape5.BorderColor
 Line5.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
 
 If Label13.Caption = "LISTEN" Then
  TrackedMode(Trackeds) = 1
  TrackedLocalPort(Trackeds) = Label14.Caption
  TrackedRemoteIP$(Trackeds) = "*"
  Trackeds = Trackeds + 1
 End If
 If Label13.Caption = "ESTAB" Then
  TrackedMode(Trackeds) = 1
  TrackedLocalPort(Trackeds) = Label14.Caption
  TrackedRemoteIP$(Trackeds) = Label15.Caption
  Trackeds = Trackeds + 1
 End If
 
 ReadyToShow = True
 Unload Me
End Sub

Private Sub Timer1_Timer()
 If Val(Label16.Caption) <> 0 Then
  SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
  Label16.Caption = Val(Label16.Caption - 1)
 End If
End Sub
