VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{570928AD-1209-11D3-967B-B4129805661E}#5.0#0"; "CSTRAY.OCX"
Begin VB.Form form1 
   BorderStyle     =   0  'None
   Caption         =   "Get TCP Table using IP Helper API"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   Icon            =   "frmTcpTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5235
      Top             =   1530
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5235
      Top             =   1050
   End
   Begin csTrayOCX.csTray csTray1 
      Left            =   5205
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "frmTcpTable.frx":08CA
      ToolTip         =   "Manipulator - Local Port Sniffer"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   4815
      TabIndex        =   13
      Text            =   "1"
      Top             =   3300
      Width           =   960
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local IP"
         Object.Width           =   2002
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local Port"
         Object.Width           =   1958
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote IP"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Port"
         Object.Width           =   1976
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "State"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   4800
      Top             =   3285
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rules"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2670
      TabIndex        =   17
      Top             =   3795
      Width           =   1440
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rules"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2685
      TabIndex        =   16
      Top             =   3810
      Width           =   1440
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00755433&
      X1              =   4110
      X2              =   4110
      Y1              =   3765
      Y2              =   4035
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00755433&
      X1              =   2670
      X2              =   4125
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Shape Shape11 
      Height          =   300
      Left            =   2655
      Top             =   3750
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   2670
      Top             =   3765
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4305
      TabIndex        =   15
      Top             =   3795
      Width           =   1440
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      Height          =   270
      Left            =   4320
      TabIndex        =   14
      Top             =   3810
      Width           =   1440
   End
   Begin VB.Shape Shape9 
      Height          =   300
      Left            =   4290
      Top             =   3750
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00755433&
      X1              =   5745
      X2              =   5745
      Y1              =   3765
      Y2              =   4035
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00755433&
      X1              =   4305
      X2              =   5760
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   4305
      Top             =   3765
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Refresh after [ s ] :"
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
      Height          =   240
      Left            =   2685
      TabIndex        =   12
      Top             =   3330
      Width           =   2235
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Refresh after [ s ] :"
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
      Height          =   240
      Left            =   2700
      TabIndex        =   11
      Top             =   3345
      Width           =   2235
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   2475
      X2              =   2475
      Y1              =   3285
      Y2              =   4170
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   45
      X2              =   2490
      Y1              =   4155
      Y2              =   4155
   End
   Begin VB.Image Image9 
      Height          =   210
      Left            =   840
      Picture         =   "frmTcpTable.frx":11A4
      Top             =   4320
      Width           =   210
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic rules learning"
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
      Height          =   240
      Left            =   375
      TabIndex        =   10
      Top             =   3900
      Width           =   2160
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic rules learning"
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
      Height          =   240
      Left            =   390
      TabIndex        =   9
      Top             =   3915
      Width           =   2160
   End
   Begin VB.Image Image8 
      Height          =   210
      Left            =   90
      Picture         =   "frmTcpTable.frx":16C6
      Top             =   3900
      Width           =   210
   End
   Begin VB.Image Image7 
      Height          =   210
      Left            =   480
      Picture         =   "frmTcpTable.frx":1BE8
      Top             =   4320
      Width           =   210
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn if Broken rules"
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
      Height          =   240
      Left            =   375
      TabIndex        =   8
      Top             =   3630
      Width           =   1845
   End
   Begin VB.Image Image6 
      Height          =   210
      Left            =   90
      Picture         =   "frmTcpTable.frx":210A
      Top             =   3615
      Width           =   210
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn if Broken rules"
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
      Height          =   240
      Left            =   390
      TabIndex        =   7
      Top             =   3645
      Width           =   1860
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00755433&
      X1              =   15
      X2              =   5895
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00755433&
      X1              =   5865
      X2              =   5865
      Y1              =   0
      Y2              =   4215
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   120
      Picture         =   "frmTcpTable.frx":262C
      Top             =   4320
      Width           =   210
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "View local to local ports"
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
      Height          =   240
      Left            =   375
      TabIndex        =   6
      Top             =   3330
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "View local to local ports"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   390
      TabIndex        =   5
      Top             =   3345
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   90
      Picture         =   "frmTcpTable.frx":2B4E
      Top             =   3315
      Width           =   210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   15
      X2              =   5865
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   5400
      X2              =   5565
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   5550
      X2              =   5550
      Y1              =   45
      Y2              =   195
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   5655
      X2              =   5820
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   5805
      X2              =   5805
      Y1              =   45
      Y2              =   195
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5385
      TabIndex        =   2
      Top             =   15
      Width           =   210
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5640
      TabIndex        =   1
      Top             =   30
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   5670
      Picture         =   "frmTcpTable.frx":3070
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   5415
      Picture         =   "frmTcpTable.frx":31AE
      Top             =   60
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   5655
      Top             =   45
      Width           =   165
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   5400
      Top             =   45
      Width           =   165
   End
   Begin VB.Shape Shape6 
      Height          =   195
      Left            =   5385
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape4 
      Height          =   195
      Left            =   5640
      Top             =   30
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - Local Port Sniffer"
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
      TabIndex        =   3
      Top             =   15
      Width           =   5250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipulator - Local Port Sniffer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   4
      Top             =   15
      Width           =   5070
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   15
      Picture         =   "frmTcpTable.frx":32EC
      Top             =   15
      Width           =   5850
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   5880
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   45
      Top             =   3285
      Width           =   2445
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   3960
      Left            =   0
      Top             =   255
      Width           =   5880
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.

Private Sub csTray1_MouseUp(Button As Integer)
 ' On retire l'icone de la barre des taches et on restaure la fenêtre
 csTray1.Visible = False
 Me.WindowState = 0
 Me.Visible = True
End Sub

Private Sub Form_Load()
 LocalToLocal = True
 AutoWarn = True
 AutoLearn = True
 FirstTime = True
 QueuedMessages = 0
 QueuedAlertMessages = 0
 ReadyToShow = True
 ReadyToShowAlert = True
 
 Trackeds = 0
 CpIpnoBanRules = 0
 CpLocalPortnoBanRules = 0
 
 Call ListConnections(-1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub Image1_Click()
 Y = Image1.Top
 X = Image1.Left
 Image1.Visible = False
 Image1.Top = Image4.Top
 Image1.Left = Image4.Left
 Image1.Visible = True
 
 Image4.Visible = False
 Image4.Top = Y
 Image4.Left = X
 Image4.Visible = True
 
 LocalToLocal = False
 Call ListConnections(-1)
End Sub

Private Sub Image4_Click()
 Y = Image4.Top
 X = Image4.Left
 Image4.Visible = False
 Image4.Top = Image1.Top
 Image4.Left = Image1.Left
 Image4.Visible = True
 
 Image1.Visible = False
 Image1.Top = Y
 Image1.Left = X
 Image1.Visible = True
 
 LocalToLocal = True
 Call ListConnections(-1)
End Sub

Private Sub Image6_Click()
 Y = Image7.Top
 X = Image7.Left
 Image7.Visible = False
 Image7.Top = Image6.Top
 Image7.Left = Image6.Left
 Image7.Visible = True
 
 Image6.Visible = False
 Image6.Top = Y
 Image6.Left = X
 Image6.Visible = True
 
 AutoWarn = False
End Sub

Private Sub Image7_Click()
 Y = Image6.Top
 X = Image6.Left
 Image6.Visible = False
 Image6.Top = Image7.Top
 Image6.Left = Image7.Left
 Image6.Visible = True
 
 Image7.Visible = False
 Image7.Top = Y
 Image7.Left = X
 Image7.Visible = True
 
 AutoWarn = True
End Sub

Private Sub Image8_Click()
 Y = Image9.Top
 X = Image9.Left
 Image9.Visible = False
 Image9.Top = Image8.Top
 Image9.Left = Image8.Left
 Image9.Visible = True
 
 Image8.Visible = False
 Image8.Top = Y
 Image8.Left = X
 Image8.Visible = True
 
 AutoLearn = False
End Sub

Private Sub Image9_Click()
 Y = Image8.Top
 X = Image8.Left
 Image8.Visible = False
 Image8.Top = Image9.Top
 Image8.Left = Image9.Left
 Image8.Visible = True
 
 Image9.Visible = False
 Image9.Top = Y
 Image9.Left = X
 Image9.Visible = True
 
 AutoLearn = True
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line10.BorderColor
 Line10.BorderColor = Shape8.BorderColor
 Line11.BorderColor = Shape8.BorderColor
 Shape8.BorderColor = Value
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape9.Visible = False Then
  Call remap
  Shape9.Visible = True
 End If
End Sub

Private Sub Label14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line10.BorderColor
 Line10.BorderColor = Shape8.BorderColor
 Line11.BorderColor = Shape8.BorderColor
 Shape8.BorderColor = Value
 
 Call ListConnections(-1)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line12.BorderColor
 Line12.BorderColor = Shape10.BorderColor
 Line13.BorderColor = Shape10.BorderColor
 Shape10.BorderColor = Value
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape11.Visible = False Then
  Call remap
  Shape11.Visible = True
 End If
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line12.BorderColor
 Line12.BorderColor = Shape10.BorderColor
 Line13.BorderColor = Shape10.BorderColor
 Shape10.BorderColor = Value
 
 Form3.Show
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 1
 X_Initial = X
 Y_Initial = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Status = 1 Then
  Me.Left = Me.Left + X - X_Initial
  Me.Top = Me.Top + Y - Y_Initial
 Else
  Call remap
 End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Status = 0
 Dist_Am = 100
 
 If Me.Left < Dist_Am Then Me.Left = 0
 If Me.Top < Dist_Am Then Me.Top = 0
 If Me.Left + Me.Width > Screen.Width - Dist_Am Then Me.Left = Screen.Width - Me.Width
 If Me.Top + Me.Height > Screen.Height - Dist_Am Then Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape7.BorderColor
 Line5.BorderColor = Shape7.BorderColor
 Shape7.BorderColor = Value

 End
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape4.Visible = False Then
  Call remap
  Shape4.Visible = True
 End If
End Sub

Private Sub Label4_Click()
 If LocalToLocal = True Then
  Y = Image1.Top
  X = Image1.Left
  Image1.Visible = False
  Image1.Top = Image4.Top
  Image1.Left = Image4.Left
  Image1.Visible = True
  Image4.Visible = False
  Image4.Top = Y
  Image4.Left = X
  Image4.Visible = True
  LocalToLocal = False
 Else
  Y = Image4.Top
  X = Image4.Left
  Image4.Visible = False
  Image4.Top = Image1.Top
  Image4.Left = Image1.Left
  Image4.Visible = True
  Image1.Visible = False
  Image1.Top = Y
  Image1.Left = X
  Image1.Visible = True
  LocalToLocal = True
 End If
End Sub

Private Sub Label6_Click()
 If AutoWarn = True Then
  Y = Image7.Top
  X = Image7.Left
  Image7.Visible = False
  Image7.Top = Image6.Top
  Image7.Left = Image6.Left
  Image7.Visible = True
  Image6.Visible = False
  Image6.Top = Y
  Image6.Left = X
  Image6.Visible = True
  AutoWarn = False
 Else
  Y = Image6.Top
  X = Image6.Left
  Image6.Visible = False
  Image6.Top = Image7.Top
  Image6.Left = Image7.Left
  Image6.Visible = True
  Image7.Visible = False
  Image7.Top = Y
  Image7.Left = X
  Image7.Visible = True
  AutoWarn = True
 End If
End Sub

Private Sub Label8_Click()
 If AutoLearn = True Then
  Y = Image9.Top
  X = Image9.Left
  Image9.Visible = False
  Image9.Top = Image8.Top
  Image9.Left = Image8.Left
  Image9.Visible = True
  Image8.Visible = False
  Image8.Top = Y
  Image8.Left = X
  Image8.Visible = True
  AutoLearn = False
 Else
  Y = Image8.Top
  X = Image8.Left
  Image8.Visible = False
  Image8.Top = Image9.Top
  Image8.Left = Image9.Left
  Image8.Visible = True
  Image9.Visible = False
  Image9.Top = Y
  Image9.Left = X
  Image9.Visible = True
  AutoLearn = True
 End If
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line2.BorderColor
 Line2.BorderColor = Shape5.BorderColor
 Line3.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
 
 Me.WindowState = 1
 Me.Visible = False
 csTray1.Visible = True
End Sub

Sub remap()
 If Shape4.Visible = True Then Shape4.Visible = False
 If Shape6.Visible = True Then Shape6.Visible = False
 If Shape9.Visible = True Then Shape9.Visible = False
 If Shape11.Visible = True Then Shape11.Visible = False
 If Shape12.Visible = True Then Shape12.Visible = False
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  SelectedIp$ = ListView1.SelectedItem.ListSubItems(2)
  Dim Cursor As POINTAPI
  Call GetCursorPos(Cursor)
  Form5.Top = (Cursor.Y * 15) - 50
  Form5.Left = (Cursor.X * 15) - 200
  Form5.Show
 End If
End Sub

Private Sub Text1_Change()
 Timer1.Interval = Val(Text1.Text) * 1000
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape12.Visible = False Then
  Call remap
  Shape12.Visible = True
 End If
End Sub

Private Sub Timer1_Timer()
 Call ListConnections(0)
End Sub

Private Sub Timer2_Timer()
 If QueuedMessages > 0 And ReadyToShow = True Then
  Call ShowQueuedMessage
 End If
 If QueuedAlertMessages > 0 And ReadyToShowAlert = True Then
  ShowQueuedAlertMessage
 End If
End Sub
