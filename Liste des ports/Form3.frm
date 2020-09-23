VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Text            =   "255.255.255.255"
      Top             =   1935
      Width           =   1350
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1650
      TabIndex        =   5
      Text            =   "0"
      Top             =   1935
      Width           =   870
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1485
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   2619
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "St"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local Port"
         Object.Width           =   3497
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote IP"
         Object.Width           =   3497
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   2760
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
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
      Left            =   3120
      TabIndex        =   14
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
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
      Left            =   3135
      TabIndex        =   13
      Top             =   2370
      Width           =   1320
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove rule"
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
      Left            =   1620
      TabIndex        =   12
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove rule"
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
      Left            =   1635
      TabIndex        =   11
      Top             =   2370
      Width           =   1320
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add rule"
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
      Left            =   135
      TabIndex        =   10
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add rule"
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
      Left            =   150
      TabIndex        =   9
      Top             =   2370
      Width           =   1320
   End
   Begin VB.Shape Shape8 
      Height          =   300
      Left            =   3105
      Top             =   2310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Shape Shape7 
      Height          =   300
      Left            =   1620
      Top             =   2310
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Shape Shape6 
      Height          =   300
      Left            =   120
      Top             =   2310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00755433&
      X1              =   1455
      X2              =   1455
      Y1              =   2325
      Y2              =   2595
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00755433&
      X1              =   135
      X2              =   1470
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00755433&
      X1              =   2940
      X2              =   2940
      Y1              =   2325
      Y2              =   2595
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00755433&
      X1              =   1635
      X2              =   2955
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00755433&
      X1              =   4440
      X2              =   4440
      Y1              =   2325
      Y2              =   2595
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00755433&
      X1              =   3120
      X2              =   4455
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   135
      Top             =   2325
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   1635
      Top             =   2325
      Width           =   1320
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   270
      Left            =   3120
      Top             =   2325
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IP :"
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
      TabIndex        =   7
      Top             =   1965
      Width           =   315
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IP :"
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
      TabIndex        =   6
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn on port :"
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
      TabIndex        =   4
      Top             =   1965
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Warn on port :"
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
      TabIndex        =   3
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   105
      Picture         =   "Form3.frx":0522
      Top             =   1950
      Width           =   210
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00755433&
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   2730
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00755433&
      X1              =   0
      X2              =   4575
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00755433&
      X1              =   15
      X2              =   4575
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FDA04F&
      FillColor       =   &H008F6100&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   0
      Top             =   255
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Rules"
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
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   4470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Sniffer - Rules"
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
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00B99063&
      FillColor       =   &H009B6F43&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   -1290
      Picture         =   "Form3.frx":0A44
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Status, X_Initial, Y_Initial, Dist_Am  ' Pour le déplacement de la fenêtre.
Dim WarnOn As Boolean

Private Sub Form_Load()
 WarnOn = True
 Call RefreshList
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
 
 WarnOn = False
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
 
 WarnOn = True
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

Sub remap()
 If Shape6.Visible = True Then Shape6.Visible = False
 If Shape7.Visible = True Then Shape7.Visible = False
 If Shape8.Visible = True Then Shape8.Visible = False
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape3.BorderColor
 Line5.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape8.Visible = False Then
  Call remap
  Shape8.Visible = True
 End If
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line4.BorderColor
 Line4.BorderColor = Shape3.BorderColor
 Line5.BorderColor = Shape3.BorderColor
 Shape3.BorderColor = Value
 
 Unload Me
End Sub

Private Sub Label3_Click()
 If WarnOn = True Then
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
  WarnOn = False
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
  WarnOn = True
 End If
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape5.BorderColor
 Line9.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape6.Visible = False Then
  Call remap
  Shape6.Visible = True
 End If
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line8.BorderColor
 Line8.BorderColor = Shape5.BorderColor
 Line9.BorderColor = Shape5.BorderColor
 Shape5.BorderColor = Value

 If WarnOn = True Then
  TrackedMode(Trackeds) = -1
 Else
  TrackedMode(Trackeds) = 1
 End If
 If Text1.Text = "*" Then
  TrackedLocalPort(Trackeds) = -1
 Else
  TrackedLocalPort(Trackeds) = Val(Text1.Text)
 End If
 TrackedRemoteIP$(Trackeds) = Text2.Text
 Trackeds = Trackeds + 1

 Call RefreshList
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape4.BorderColor
 Line7.BorderColor = Shape4.BorderColor
 Shape4.BorderColor = Value
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Shape7.Visible = False Then
  Call remap
  Shape7.Visible = True
 End If
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Value = Line6.BorderColor
 Line6.BorderColor = Shape4.BorderColor
 Line7.BorderColor = Shape4.BorderColor
 Shape4.BorderColor = Value
 'Trackeds
 
 Value = -1
 For i = ListView1.ListItems.Count To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   Value = i - 1
   If Value = Trackeds - 1 Then
    Trackeds = Trackeds - 1
   Else
    TrackedMode(Value) = TrackedMode(Trackeds - 1)
    TrackedLocalPort(Value) = TrackedLocalPort(Trackeds - 1)
    TrackedRemoteIP$(Value) = TrackedRemoteIP$(Trackeds - 1)
    Trackeds = Trackeds - 1
   End If
  End If
 Next i
 
 Call RefreshList
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call remap
End Sub

Sub RefreshList()
 ListView1.ListItems.Clear
 For i = 0 To Trackeds - 1
  If TrackedMode(i) = 1 Then
   LeMode$ = "+"
  Else
   LeMode$ = "-"
  End If
  Set lvItem = ListView1.ListItems.Add(, , LeMode$)
  If TrackedLocalPort(i) = -1 Then
   lvItem.SubItems(1) = "*"
  Else
   lvItem.SubItems(1) = TrackedLocalPort(i)
  End If
  If TrackedRemoteIP$(i) = "*" Then
   lvItem.SubItems(2) = "*"
  Else
   lvItem.SubItems(2) = TrackedRemoteIP$(i)
  End If
 Next i
End Sub
