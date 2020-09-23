Attribute VB_Name = "Module1"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetTcpTable Lib "IPhlpAPI" (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Type MIB_TCPROW
 dwState As Long
 dwLocalAddr As String * 4
 dwLocalPort As String * 4
 dwRemoteAddr As String * 4
 dwRemotePort As String * 4
End Type

Type MIB_TCPTABLE
 dwNumEntries As Long
 table(100) As MIB_TCPROW
End Type

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Public Const ERROR_BUFFER_OVERFLOW = 111&
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_NO_DATA = 232&
Public Const ERROR_NOT_SUPPORTED = 50&
Public Const ERROR_SUCCESS = 0&

Public Const MIB_TCP_STATE_CLOSED = 0
Public Const MIB_TCP_STATE_LISTEN = 1
Public Const MIB_TCP_STATE_SYN_SENT = 2
Public Const MIB_TCP_STATE_SYN_RCVD = 3
Public Const MIB_TCP_STATE_ESTAB = 4
Public Const MIB_TCP_STATE_FIN_WAIT1 = 5
Public Const MIB_TCP_STATE_FIN_WAIT2 = 6
Public Const MIB_TCP_STATE_CLOSE_WAIT = 7
Public Const MIB_TCP_STATE_CLOSING = 8
Public Const MIB_TCP_STATE_LAST_ACK = 9
Public Const MIB_TCP_STATE_TIME_WAIT = 10
Public Const MIB_TCP_STATE_DELETE_TCB = 11

Public LocalToLocal As Boolean
Public AutoWarn As Boolean
Public AutoLearn As Boolean
Public FirstTime As Boolean
Public ReadyToShow As Boolean
Public ReadyToShowAlert As Boolean
Public SelectedIp$

' Liste des connexions actives
Public Connections As Integer
Public ConnectionsPort(3000) As Long
Public ConnectionIp$(3000)
Public ConnectionState$(3000)

' Messages mis en attente d'affichage [QUERY]
Public QueuedMessages
Public QueuedPort(100) As Long
Public QueuedIP$(100)
Public QueuedState$(100)

' Messages mis en attente d'affichage [WARNING]
Public QueuedAlertMessages
Public QueuedAlertLine1$(100)
Public QueuedAlertLine2$(100)

' Permet de gérer la liste de surveillance
Public TrackedMode(1000) As Integer
Public TrackedLocalPort(1000) As Long
Public TrackedRemoteIP$(1000)
Public Trackeds As Integer

Sub ListConnections(mode As Integer)
 Dim pTcpTable As MIB_TCPTABLE
 Dim pdwSize As Long
 Dim bOrder As Long
 Dim nRet As Long
 Dim i As Integer, s As String
 
 If FirstTime = True Then
  FirstTime = False
  Call MapTables
 End If
 
 txtOutput = ""
 nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
 nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
 
 If pTcpTable.dwNumEntries <> Connections Or mode = -1 Then
  form1.ListView1.ListItems.Clear
 Else
  Exit Sub
 End If
 
 For i = 0 To pTcpTable.dwNumEntries - 1
  If pTcpTable.table(i).dwState - 1 <> MIB_TCP_STATE_LISTEN Then
   If LocalToLocal = False And (c_ip(pTcpTable.table(i).dwRemoteAddr) = "0.0.0.0" Or c_ip(pTcpTable.table(i).dwRemoteAddr) = "127.0.0.1") Then
   Else
    Value = form1.ListView1.ListItems.Count
    Set lvItem = form1.ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    lvItem.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    lvItem.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    lvItem.SubItems(3) = c_port(pTcpTable.table(i).dwRemotePort)
    lvItem.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    form1.ListView1.ListItems(Value + 1).ToolTipText = PortName$(c_port(pTcpTable.table(i).dwLocalPort), c_port(pTcpTable.table(i).dwRemotePort))
    
    Found = False
    For t = 0 To Connections - 1
     If (ConnectionsPort(t) = c_port(pTcpTable.table(i).dwLocalPort)) And (ConnectionIp$(t) = c_ip(pTcpTable.table(i).dwRemoteAddr)) And (ConnectionState$(t) = c_state(pTcpTable.table(i).dwState - 1)) Then Found = True
    Next t
    If Found = False Then Call Notify(c_port(pTcpTable.table(i).dwLocalPort), c_state(pTcpTable.table(i).dwState - 1), c_ip(pTcpTable.table(i).dwRemoteAddr))
       
   End If
  Else
   If LocalToLocal = False And (c_ip(pTcpTable.table(i).dwLocalAddr) = "0.0.0.0" Or c_ip(pTcpTable.table(i).dwLocalAddr) = "127.0.0.1") Then
   Else
    Value = form1.ListView1.ListItems.Count
    Set lvItem = form1.ListView1.ListItems.Add(, , c_ip(pTcpTable.table(i).dwLocalAddr))
    lvItem.SubItems(1) = c_port(pTcpTable.table(i).dwLocalPort)
    lvItem.SubItems(2) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    lvItem.SubItems(3) = "-"
    lvItem.SubItems(4) = c_state(pTcpTable.table(i).dwState - 1)
    form1.ListView1.ListItems(Value + 1).ToolTipText = PortName$(c_port(pTcpTable.table(i).dwLocalPort), 0)
    
    Found = False
    For t = 0 To Connections - 1
     If (ConnectionsPort(t) = c_port(pTcpTable.table(i).dwLocalPort)) And (ConnectionIp$(t) = c_ip(pTcpTable.table(i).dwRemoteAddr)) And (ConnectionState$(t) = c_state(pTcpTable.table(i).dwState - 1)) Then Found = True
    Next t
    If Found = False Then Call Notify(c_port(pTcpTable.table(i).dwLocalPort), c_state(pTcpTable.table(i).dwState - 1), c_ip(pTcpTable.table(i).dwRemoteAddr))
       
   End If
  End If
 Next

 Call MapTables
End Sub

Sub MapTables()
 Dim pTcpTable As MIB_TCPTABLE
 Dim pdwSize As Long
 Dim bOrder As Long
 Dim nRet As Long
 Dim i As Integer, s As String
 
 txtOutput = ""
 Connections = 0
 nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
 nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
 For i = 0 To pTcpTable.dwNumEntries - 1
  If pTcpTable.table(i).dwState - 1 <> MIB_TCP_STATE_LISTEN Then
   If LocalToLocal = False And (c_ip(pTcpTable.table(i).dwRemoteAddr) = "0.0.0.0" Or c_ip(pTcpTable.table(i).dwRemoteAddr) = "127.0.0.1") Then
   Else
    ConnectionsPort(Connections) = c_port(pTcpTable.table(i).dwLocalPort)
    ConnectionIp$(Connections) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    ConnectionState$(Connections) = c_state(pTcpTable.table(i).dwState - 1)
    Connections = Connections + 1
   End If
  Else
   If LocalToLocal = False And (c_ip(pTcpTable.table(i).dwLocalAddr) = "0.0.0.0" Or c_ip(pTcpTable.table(i).dwLocalAddr) = "127.0.0.1") Then
   Else
    ConnectionsPort(Connections) = c_port(pTcpTable.table(i).dwLocalPort)
    ConnectionIp$(Connections) = c_ip(pTcpTable.table(i).dwRemoteAddr)
    ConnectionState$(Connections) = c_state(pTcpTable.table(i).dwState - 1)
    Connections = Connections + 1
   End If
  End If
 Next
End Sub

Function c_port(s) As Long
 c_port = Asc(Mid(s, 1, 1)) * 256 + Asc(Mid(s, 2, 1))
End Function

Function c_ip(s) As String
 c_ip = Asc(Mid(s, 1, 1)) & "." & Asc(Mid(s, 2, 1)) & "." & Asc(Mid(s, 3, 1)) & "." & Asc(Mid(s, 4, 1))
End Function

Function c_state(s) As String
 Select Case s
  Case MIB_TCP_STATE_CLOSED: c_state = "CLOSED"
  Case MIB_TCP_STATE_LISTEN: c_state = "LISTEN"
  Case MIB_TCP_STATE_SYN_SENT: c_state = "SYN_SENT"
  Case MIB_TCP_STATE_SYN_RCVD: c_state = "SYN_RCVD"
  Case MIB_TCP_STATE_ESTAB: c_state = "ESTAB"
  Case MIB_TCP_STATE_FIN_WAIT1: c_state = "FIN_WAIT1"
  Case MIB_TCP_STATE_FIN_WAIT2: c_state = "FIN_WAIT2"
  Case MIB_TCP_STATE_CLOSE_WAIT: c_state = "CLOSE_WAIT"
  Case MIB_TCP_STATE_CLOSING: c_state = "CLOSING"
  Case MIB_TCP_STATE_LAST_ACK: c_state = "LAST_ACK"
  Case MIB_TCP_STATE_TIME_WAIT: c_state = "TIME_WAIT"
  Case MIB_TCP_STATE_DELETE_TCB: c_state = "DELETE_TCB"
  Case Else: c_state = "UNDEFINED"
 End Select
End Function

Sub ShowQueuedAlertMessage()
 If QueuedAlertMessages = 0 Then Exit Sub
 If QueuedAlertMessages = 1 Then
  QueuedAlertMessages = 0
  Call ShowAlert(QueuedAlertLine1$(0), QueuedAlertLine2$(0))
 Else
  Call ShowAlert(QueuedAlertLine1$(0), QueuedAlertLine2$(0))
  For i = 0 To QueuedAlertMessages - 1
   QueuedAlertLine1$(i) = QueuedAlertLine1$(i + 1)
   QueuedAlertLine2$(i) = QueuedAlertLine2$(i + 1)
  Next i
  QueuedAlertMessages = QueuedAlertMessages - 1
 End If
End Sub

Sub ShowAlert(Line1$, Line2$)
 If ReadyToShowAlert = True Then
  ReadyToShowAlert = False
  Form4.Label4.Caption = Line1$
  Form4.Label5.Caption = Line1$
  Form4.Label6.Caption = Line2$
  Form4.Label7.Caption = Line2$
  Form4.Show
 Else
  If QueuedAlertMessages >= 100 Then Exit Sub
  QueuedAlertLine1$(QueuedAlertMessages) = Line1$
  QueuedAlertLine2$(QueuedAlertMessages) = Line2$
  QueuedAlertMessages = QueuedAlertMessages + 1
 End If
End Sub

Sub Notify(Port As Long, State$, IP$)
 If AutoWarn = False Then Exit Sub
 
 Value = IsTracked(Port, State$, IP$)
 If Value = -1 Then Exit Sub
 If Value = -2 Then
  Call ShowAlert("Rules break on port #" + Str$(Port) + " with the computer", "adress " + IP$)
  Exit Sub
 End If
 
 If AutoLearn = False Then Exit Sub
  
 If ReadyToShow = True Then
  ReadyToShow = False
  Call ShowNotification(Port, State$, IP$)
 Else
  If QueuedMessages >= 100 Then Exit Sub
  QueuedPort(QueuedMessages) = Port
  QueuedIP$(QueuedMessages) = IP$
  QueuedState$(QueuedMessages) = State$
  QueuedMessages = QueuedMessages + 1
 End If
End Sub

Sub ShowNotification(Port As Long, State$, IP$)
 Done = False
 If State$ = "LISTEN" Then
  Form2.Label14.Caption = Port
  Form2.Label15.Caption = IP$
  Form2.Label3.Caption = "The port #" + Str$(Port) + " has been opened on your"
  Form2.Label4.Caption = "The port #" + Str$(Port) + " has been opened on your"
  Form2.Label5.Caption = "computer. This could be a trojan."
  Form2.Label6.Caption = "computer. This could be a trojan."
  Form2.Label13.Caption = "LISTEN"
  Form2.Show
  Done = True
 End If
 If State$ = "ESTAB" Then
  Form2.Label14.Caption = Port
  Form2.Label15.Caption = IP$
  Form2.Label3.Caption = "The Computer " + IP$ + " has connected to your"
  Form2.Label4.Caption = "The Computer " + IP$ + " has connected to your"
  Form2.Label5.Caption = "system on port #" + Str$(Port)
  Form2.Label6.Caption = "system on port #" + Str$(Port)
  Form2.Label13.Caption = "ESTAB"
  Form2.Show
  Done = True
 End If
 If Done = False Then ReadyToShow = True
End Sub

Sub ShowQueuedMessage()
 If QueuedMessages = 1 Then
  ReadyToShow = False
  Call ShowNotification(QueuedPort(0), QueuedState$(0), QueuedIP$(0))
  QueuedMessages = 0
 Else
  ReadyToShow = False
  Call ShowNotification(QueuedPort(0), QueuedState$(0), QueuedIP$(0))
  For i = 0 To QueuedMessages - 1
   QueuedPort(i) = QueuedPort(i + 1)
   QueuedState$(i) = QueuedState$(i + 1)
   QueuedIP$(i) = QueuedIP$(i + 1)
  Next i
  QueuedMessages = QueuedMessages - 1
 End If
End Sub

Function IsTracked(Port As Long, State$, IP$)
 IsTracked = 0
 
 For i = 0 To Trackeds - 1
  If TrackedRemoteIP$(i) = "*" Then
   If TrackedLocalPort(i) = -1 Then
    If TrackedMode(i) = 1 Then IsTracked = -1: Exit Function
    If TrackedMode(i) = -1 Then IsTracked = -2: Exit Function
   Else
    If TrackedLocalPort(i) = Port Then
     If TrackedMode(i) = 1 Then IsTracked = -1: Exit Function
     If TrackedMode(i) = -1 Then IsTracked = -2: Exit Function
    End If
   End If
  Else
   If TrackedRemoteIP$(i) = IP$ Then
    If TrackedLocalPort(i) = -1 Then
     If TrackedMode(i) = 1 Then IsTracked = -1: Exit Function
     If TrackedMode(i) = -1 Then IsTracked = -2: Exit Function
    Else
     If TrackedLocalPort(i) = Port Then
      If TrackedMode(i) = 1 Then IsTracked = -1: Exit Function
      If TrackedMode(i) = -1 Then IsTracked = -2: Exit Function
     End If
    End If
   End If
  End If
 Next i
End Function

Function PortName$(Gin, Gout)
 If Dir$("ports.lst") <> "" Then
  Inport$ = "Unknown": OutPort$ = "Unknown"
  Open "ports.lst" For Input As #1
  Do
   Input #1, Port, Name$, Desc$
   Legende$ = Desc$
   If Desc$ = "" Then Legende$ = Name$
   If Port = Gin Then Inport$ = Legende$
   If Port = Gout Then OutPort$ = Legende$
  Loop Until EOF(1)
  Close #1
  PortName$ = Inport$ + " / " + OutPort
 Else
  PortName$ = "Unknow / Unknow"
 End If
End Function

