VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form4"
   ScaleHeight     =   1575
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4395
      Top             =   360
   End
   Begin VB.Label Label8 
      Caption         =   "10"
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label7 
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
      Left            =   810
      TabIndex        =   7
      Top             =   675
      Width           =   3945
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   825
      TabIndex        =   6
      Top             =   690
      Width           =   3945
   End
   Begin VB.Label Label5 
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
      Left            =   810
      TabIndex        =   5
      Top             =   450
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   825
      TabIndex        =   4
      Top             =   465
      Width           =   3945
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      Left            =   3465
      TabIndex        =   3
      Top             =   1215
      Width           =   1320
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1230
      Width           =   1320
   End
   Begin VB.Shape Shape4 
      Height          =   300
      Left            =   3450
      Top             =   1170
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00004080&
      X1              =   4785
      X2              =   4785
      Y1              =   1185
      Y2              =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004080&
      X1              =   3465
      X2              =   4800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000060E0&
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   3465
      Top             =   1185
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   165
      Picture         =   "Form4.frx":0000
      Top             =   405
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Alert!"
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
      TabIndex        =   0
      Top             =   30
      Width           =   4785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Alert!"
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
      TabIndex        =   1
      Top             =   45
      Width           =   4815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      X1              =   0
      X2              =   4935
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      X1              =   4920
      X2              =   4920
      Y1              =   0
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      X1              =   0
      X2              =   4935
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000060E0&
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   0
      Top             =   255
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000060E0&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   -930
      Picture         =   "Form4.frx":0882
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Private Sub Form_Load()
 Label8.Caption = "10"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

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

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape3.BorderColor
 Line5.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Sub remap()
 If Shape4.Visible = True Then Shape4.Visible = False
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape3.BorderColor
 Line5.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value

 ReadyToShowAlert = True
 Unload Me
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Timer1_Timer()
 If Val(Label8.Caption) <> 0 Then
  SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H20 Or &H40
  Label8.Caption = Val(Label8.Caption - 1)
 End If
End Sub
