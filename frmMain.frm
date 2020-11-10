VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BetaScanner 1.0.2"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10140
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRanges 
      Appearance      =   0  'Flat
      Caption         =   " Range Selection  "
      Height          =   1575
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   4095
      Begin VB.CommandButton cmdRangeNext 
         Caption         =   ">>"
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton CmdRangeBack 
         Caption         =   "<<"
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstRangeSelection 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   " Method "
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   5400
      Width           =   4095
      Begin VB.OptionButton OptionMethod 
         Appearance      =   0  'Flat
         Caption         =   "Ranges.txt"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionMethod 
         Appearance      =   0  'Flat
         Caption         =   "Scanlist.txt"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionMethod 
         Appearance      =   0  'Flat
         Caption         =   "Manual"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSWinsockLib.Winsock winSock 
      Index           =   0
      Left            =   4320
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame FrameScanDetails 
      Appearance      =   0  'Flat
      Caption         =   " Scan Details "
      Height          =   1935
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   4095
      Begin VB.CheckBox CheckHTTP 
         Caption         =   "Scan as HTTP unless otherwise specified"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtScanProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtScanThreads 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtScanTimeout 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtRangePort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtRangeEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtRangeStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress && Information"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Threads"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FrameControl 
      Appearance      =   0  'Flat
      Caption         =   " Control "
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   4095
      Begin VB.CommandButton cmdTrayApplication 
         Caption         =   "Send to Tray"
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdScanStop 
         Caption         =   "Pause Scan"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdScanStart 
         Caption         =   "Start Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frameResults 
      Appearance      =   0  'Flat
      Caption         =   " Results "
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   4095
      Begin VB.Label lblConnectionsMade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Active Connections"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblDeniedRequest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Denied Request"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblGrantedUse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Granted Use"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblBnetVerified 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BattleNet Verified"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FrameVerification 
      Appearance      =   0  'Flat
      Caption         =   " Verification IP-Address && Port "
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   4095
      Begin VB.CommandButton cmdVerificationUpdate 
         Caption         =   "Save All Settings"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   255
         Width           =   1455
      End
      Begin VB.TextBox txtVerificationPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtVerificationIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image icoFinished 
      Height          =   480
      Left            =   6120
      Picture         =   "frmMain.frx":08CA
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblBanner 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Official Release -  "
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   -10
      Width           =   4095
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

Dim destServerHash As String, destPortHash As String
Dim destServerNormal As String, destPortNormal As Long
Dim lastMethodindex As Byte
Dim connectionsMade As Long
'
Dim verificationIPdata() As String
'for auto complete of verification ips



Private Sub CmdRangeBack_Click()
  
  If proxyRangesIndex > 0 Then
    proxyRangesIndex = proxyRangesIndex - 1
    set_NewRange
  End If

End Sub
'

Private Sub cmdRangeNext_Click()
  
  If proxyRangesIndex < (lstRangeSelection.ListCount - 1) Then
    proxyRangesIndex = proxyRangesIndex + 1
    set_NewRange
  End If

End Sub

Private Sub cmdScanStart_Click()

  cancelNext = False
  
  If Not cmdScanStop.Enabled Then
    set_winSocks
    '
    set_currentIP txtRangeStart.Text
    '
    ReDim proxyTicks(winSock.Count - 1)
    ReDim proxyHTTP(winSock.Count - 1)
    '
    proxyScanlistIndex = 0
    'reset scanlist/range index
    connectionsMade = connectionsMade = 0
    '
  End If
  
  With FrmMain
    .txtRangeStart.Enabled = False
    .txtRangeEnd.Enabled = False
    .txtRangePort.Enabled = False
    .txtScanThreads.Enabled = False
    .txtScanTimeout.Enabled = False
    .cmdScanStart.Enabled = False
    .cmdScanStop.Enabled = True
    .cmdScanStop.Caption = "Pause Scan"
  End With
  '
'  SetTimer Me.hwnd, 5, Val#(FrmMain.txtScanTimeout.Text), AddressOf checkChecked
  original_TimeoutTick = Val#(FrmMain.txtScanTimeout.Text)
  connect_ProxySet

End Sub

Private Sub set_winSocks()

  Do Until FrmMain.winSock.Count = 1
    Unload winSock(winSock.Count - 1)
  Loop

  Do Until winSock.Count = txtScanThreads.Text
    Load winSock(winSock.Count)
  Loop

End Sub
'spawns threads



Public Sub cmdScanStop_Click()
  
  cancelNext = True
  'stop or pause
  Select Case cmdScanStop.Caption
    Case "Pause Scan"
      With FrmMain
        .txtRangePort.Enabled = True
        .txtScanTimeout.Enabled = True
        .cmdScanStart.Enabled = True
        .cmdScanStop.Enabled = True
        .cmdScanStop.Caption = "Stop Scan"
      End With
    'stops and pauses
    Case "Stop Scan"
      With FrmMain
        .txtRangeStart.Enabled = True
        .txtRangeEnd.Enabled = True
        .txtRangePort.Enabled = True
        .txtScanThreads.Enabled = True
        .txtScanTimeout.Enabled = True
        .cmdScanStart.Enabled = True
        .cmdScanStop.Enabled = False
        If methodType = 2 And lstRangeSelection.ListCount <> 0 Then
          proxyRangesIndex = 0
          set_NewRange
        End If
      End With
    'stops and resets
  End Select
  
End Sub
'pauses or stops scanning

Private Sub cmdTrayApplication_Click()
  
  With notifID
    .cbSize = Len(notifID)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon 'can also use picture boxes
    If cancelNext Then
      .szTip = "BetaScanner - Idle..." & vbNullChar
    Else
      .szTip = "BetaScanner - Scanning..." & vbNullChar
    End If
  End With
  Shell_NotifyIcon NIM_ADD, notifID
  '
  Me.Visible = False
  
End Sub

Private Sub save_Settings(settingTypes As Byte)
Dim proxySplit() As String
Dim byteIndex As Byte
On Error GoTo Err

  With FrmMain
'############# SAVES VERIFICATION INFO
    If settingTypes = 0 Or settingTypes = 1 Then
      WriteINI "VerificationServer", .txtVerificationIP.Text
      proxySplit = Split(.txtVerificationIP.Text, ".")
      If UBound(proxySplit) = 3 Then
        destServerHash = vbNullString
        For byteIndex = 0 To UBound(proxySplit)
          destServerHash = destServerHash & Chr$(CInt(proxySplit(byteIndex)))
        Next byteIndex
      End If
      'save server-address hash
      WriteINI "VerificationPort", .txtVerificationPort.Text
      destPortHash = get_PortHash(.txtVerificationPort.Text)
      'save server-port hash
      destServerNormal = .txtVerificationIP.Text
      destPortNormal = .txtVerificationPort.Text
      'saves normal verification
    End If
'############# SAVES SCAN DETAILS
    WriteINI "HTTP", Val#(CheckHTTP.Value)
    If settingTypes = 0 Or settingTypes = 2 Then
      WriteINI "StartIP", .txtRangeStart.Text
      WriteINI "EndIP", .txtRangeEnd.Text
      WriteINI "ScanPORT", .txtRangePort.Text
      WriteINI "Timeout", .txtScanTimeout.Text
      WriteINI "Threads", .txtScanThreads.Text
    End If
'############# SAVES SCAN METHOD
    If settingTypes = 0 Or settingTypes = 3 Then
      If OptionMethod(0).Value Then
        WriteINI "Method", "0"
      ElseIf OptionMethod(1).Value Then
        WriteINI "Method", "1"
      ElseIf OptionMethod(2).Value Then
        WriteINI "Method", "2"
      End If
    End If
'//############# END
  End With



  If settingTypes = 0 Then txtScanProgress.Text = "Saved! " & Time$

Err:
  If Err.Number <> 0 Then txtScanProgress.Text = "Save Error! " & Time$
  
End Sub
'saves all settings to variables and INI

Private Sub cmdVerificationUpdate_Click()
  save_Settings 0
End Sub
'save settings button

Private Sub Form_Load()
On Error GoTo Err

  Me.Height = 8340
  Me.Width = 4425
  load_Settings 0
  cancelNext = True 'set so we know our scanning progress for tray
  MkDir App.Path & "\Proxies\"

Err:
  lblBanner.Caption = lblBanner.Caption & ChrW$(67) & ChrW$(111) & ChrW$(100) & ChrW$(101) & ChrW$(100) & ChrW$(32) & ChrW$(98) & ChrW$(121) & ChrW$(32) & ChrW$(70) & ChrW$(108) & ChrW$(101) & ChrW$(101) & ChrW$(116)
  'coded by fleet-
End Sub

Private Sub load_Settings(settingsType As Byte)
Dim fileData As String, newFile As Integer
On Error GoTo Err

  ReDim verificationIPdata(0)
  If LenB(Dir$(App.Path & "\VerificationIPs.txt")) <> 0 Then
    newFile = FreeFile
    Open App.Path & "\VerificationIPs.txt" For Input As #newFile
    Do Until EOF(newFile)
    Line Input #newFile, fileData
      If InStrB(2, fileData, ":") And InStrB(4, fileData, ChrW$(32)) Then
        verificationIPdata(UBound(verificationIPdata)) = fileData
        ReDim Preserve verificationIPdata(UBound(verificationIPdata) + 1)
      End If
    Loop
    Close #newFile
  End If
  'load autocompletes for verification ips


  With FrmMain
'%%%%%%%%%%%%%%%%%% LOADS VERIFICATION
    If settingsType = 0 Or settingsType = 1 Then
      .txtVerificationIP.Text = GetINI("VerificationServer")
      .txtVerificationPort.Text = GetINI("VerificationPort")
    End If
'%%%%%%%%%%%%%%%%%% LOADS SCAN DETAILS
    .CheckHTTP = Val#(GetINI("HTTP"))
    If settingsType = 0 Or settingsType = 2 Then
      .txtRangeStart.Text = GetINI("StartIP")
      .txtRangeEnd.Text = GetINI("EndIP")
      .txtRangePort.Text = GetINI("ScanPort")
      .txtScanTimeout.Text = GetINI("Timeout")
      .txtScanThreads.Text = GetINI("Threads")
    End If
'%%%%%%%%%%%%%%%%%% LOADS METHOD TYPE
    If settingsType = 0 Or settingsType = 3 Then
      lastMethodindex = CInt(GetINI("Method"))
      methodType = lastMethodindex
      .OptionMethod(methodType).Value = True
      If .OptionMethod(1).Value Then
        Load_Scanlist
      ElseIf .OptionMethod(2).Value Then
        Load_Ranges
      End If
    End If
'//%%%%%%%%%%%%%%%%%% END
  End With
  
  If settingsType = 0 Then
    cmdVerificationUpdate_Click
    txtScanProgress.Text = "Program Loaded!"
  End If

Err:
  If Err.Number <> 0 Then txtScanProgress.Text = "Load Error! " & Time$
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result, Action As Long
    
  If Me.ScaleMode = vbPixels Then
    Action = X
  Else
    Action = X / Screen.TwipsPerPixelX
  End If
  
  Select Case Action
    Case WM_LBUTTONDBLCLK 'Left Button Double Click
      Shell_NotifyIcon NIM_DELETE, notifID 'remove from tray
      Me.Show
  End Select

End Sub
'return from tray




Private Sub lstRangeSelection_Click()
  
  If lstRangeSelection.ListIndex <> -1 Then
    proxyRangesIndex = lstRangeSelection.ListIndex
    set_NewRange
  End If
  
End Sub
'set starting range index


Private Sub OptionMethod_Click(Index As Integer)

  If cmdScanStop.Enabled Or Index = lastMethodindex Or LenB(txtScanProgress.Text) = 0 Then
    If Index <> lastMethodindex Then OptionMethod(lastMethodindex).Value = True
    Exit Sub
  End If
  
  If Index = 0 Then
    ReDim proxyScanlist(0)
    lstRangeSelection.Clear
    txtScanProgress.Text = "Lists dropped! " & Time$
    load_Settings 2
  'manual
  ElseIf Index = 1 Then
    If MsgBox("Would you like to Load / Reload Scanlist.txt?", vbYesNo) = 6 Then 'if YES
      If LenB(Dir$(App.Path & "\Scanlist.txt")) = 0 Then
        MsgBox "Scanlist.txt is missing, cannot use this scanning type."
        OptionMethod(lastMethodindex).Value = True
        Exit Sub
      'failed load, file missing
      Else
        lstRangeSelection.Clear
        Load_Scanlist
      End If
    'yes load scanlist
    Else
      OptionMethod(lastMethodindex).Value = True
      Exit Sub
    End If
  'load scanlist
  ElseIf Index = 2 Then
    If MsgBox("Would you like to Load / Reload Ranges.txt?", vbYesNo) = 6 Then 'if YES
      If LenB(Dir$(App.Path & "\Ranges.txt")) = 0 Then
        MsgBox "Ranges.txt is missing, cannot use this scanning type."
        OptionMethod(lastMethodindex).Value = True
        Exit Sub
      'failed load, file missing
      Else
        ReDim proxyScanlist(0)
        Load_Ranges
      End If
    'yes load ranges
    Else
      OptionMethod(lastMethodindex).Value = True
      Exit Sub
    End If
  'load ranges
  End If

  overrideHTTP = False
  overrideS4 = False
  proxyRangesIndex = 0
  set_NewRange
  '
  save_Settings 3
  methodType = Index
  lastMethodindex = Index

End Sub
'prompt for reloading files on method change


Private Sub txtRangeEnd_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtRangePort_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtRangeStart_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtScanThreads_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtScanTimeout_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtVerificationIP_Change()
Dim arrayIndex As Integer
Dim VeriAddy As String, VeriString As String

  DoEvents
  Do Until arrayIndex = UBound(verificationIPdata)
    VeriString = Mid$(verificationIPdata(arrayIndex), InStr(verificationIPdata(arrayIndex), ChrW$(32)) + 1)
    If InStrB(LCase$(txtVerificationIP.Text), LCase$(VeriString)) And Len(txtVerificationIP.Text) > 2 Then
      VeriAddy = verificationIPdata(arrayIndex)
      VeriAddy = Left$(VeriAddy, InStr(VeriAddy, ChrW$(32)) - 1)
      '
      txtVerificationPort.Text = Mid$(VeriAddy, InStr(VeriAddy, ":") + 1)
      txtVerificationIP.Text = Left$(VeriAddy, InStr(VeriAddy, ":") - 1)
      Exit Sub
    End If
    'match found set values and exit sub
    arrayIndex = arrayIndex + 1
  Loop
  
End Sub
'quick server change by partial match

Private Sub txtVerificationIP_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub txtVerificationPort_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdVerificationUpdate_Click
End Sub

Private Sub winSock_Close(Index As Integer)
  modifyConnectedDisplay False
End Sub

Private Sub winSock_Connect(Index As Integer)
Dim proxySplit() As String
Dim byteIndex As Byte
On Error GoTo Err

  modifyConnectedDisplay True
  'display purposes
  proxyTicks(Index) = GTC + 6000
  'DoEvents
  If Not proxyHTTP(Index) Then
  'only check if not true
    If overrideHTTP Then
      proxyHTTP(Index) = True
    ElseIf overrideS4 Then
      proxyHTTP(Index) = False
    'check for overrides
    Else
    'if no overrides
      If CheckHTTP.Value = 1 Then
          proxyHTTP(Index) = True
      'if always use http then
      Else
      'if not checked, determine by ports
          Select Case winSock(Index).RemotePort
              Case 80, 8080, 3124, 3125, 3126, 3127, 3128: proxyHTTP(Index) = True
              Case Else: proxyHTTP(Index) = False
          End Select
          'if known http ports then
      End If
    End If
    'sets if http or not
  End If
  
  If proxyHTTP(Index) Then
    winSock(Index).SendData "CONNECT " & destServerNormal & ":" & destPortNormal & " HTTP/1.1" & vbCrLf & vbCrLf
    'send http
  Else
    winSock(Index).SendData Chr$(&H4) & Chr$(&H1) & destPortHash & destServerHash & Chr$(&H0)
    'send proxy hash for socks4
  End If
  
Err:
End Sub

Private Sub modifyConnectedDisplay(increaseValue As Boolean)
  
  If increaseValue Then
    connectionsMade = connectionsMade + 1
  Else
    If connectionsMade > 0 Then connectionsMade = connectionsMade - 1
    If connectionsMade < 0 Then connectionsMade = 0
  End If
  lblConnectionsMade.Caption = connectionsMade

End Sub


Private Sub winSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim socketData As String, resultCode As Integer
Static lastIndex As Integer, lastTick As Long, thisTick As Long
On Error GoTo Err

  winSock(Index).GetData socketData
  'buffer data
  
  If Asc(Left$(socketData, 1)) = &HFF Then
    Addproxy_Bnet winSock(Index).RemoteHostIP & ":" & winSock(Index).RemotePort
    DoEvents
    winSock(Index).Close
    '
    modifyConnectedDisplay False
    proxyTicks(Index) = 0
  'bnet verified!
  Else
    If proxyHTTP(Index) Then
    'is http
        If InStrB(socketData, "200 OK") Then
            proxyTicks(Index) = GTC + 6000
            
            Addproxy_Granted winSock(Index).RemoteHostIP & ":" & winSock(Index).RemotePort
            DoEvents
            
            winSock(Index).SendData Chr$(1)
            winSock(Index).SendData Chr$(&HFF) & Chr$(50) & Chr$(0) & Chr$(58) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & "68XINB2W" & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(&H4F) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & "USA" & Chr$(0) & "United States" & Chr$(0)
            'granted, so verify bnet            '
        Else
            Addproxy_Denied winSock(Index).RemoteHostIP & ":" & winSock(Index).RemotePort
            '
            proxyTicks(Index) = 0
            '
            DoEvents
            winSock(Index).Close
            modifyConnectedDisplay False
        'denied
        End If
    Else
    'not http
        resultCode = Asc(Mid$(socketData, 2, 1))
        If AscB(socketData) = 0 And AscB(resultCode) = 57 Then
          If Right$(resultCode, 1) = 0 Then
            proxyTicks(Index) = GTC + 6000
            '
            Addproxy_Granted winSock(Index).RemoteHostIP & ":" & winSock(Index).RemotePort
            DoEvents
            
            winSock(Index).SendData Chr$(1)
            winSock(Index).SendData Chr$(&HFF) & Chr$(50) & Chr$(0) & Chr$(58) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & "68XINB2W" & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(&H4F) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & "USA" & Chr$(0) & "United States" & Chr$(0)
            'granted, so verify bnet
          Else
            Addproxy_Denied winSock(Index).RemoteHostIP & ":" & winSock(Index).RemotePort
            '
            proxyTicks(Index) = 0
            '
            DoEvents
            winSock(Index).Close
            modifyConnectedDisplay False
            'denied
          End If
        End If
    'not http
    End If
  'verify socks
  End If
  
  

Err:
End Sub


Private Function get_PortHash(portString As String) As String
Dim spaceCount As Byte, tS As Byte
Dim biN As String, biN2 As String

  get_PortHash = portString
 
  Do Until spaceCount = 16
    
    spaceCount = spaceCount + 1

    If Val#(get_PortHash) > 0 Then
      tS = get_PortHash Mod 2
      get_PortHash = Int(get_PortHash / 2)
    Else
      tS = 0
    End If

    If spaceCount > 8 Then
      biN2 = tS & biN2       'Byte 2
    Else
      biN = tS & biN    'Byte 1
    End If
    
  Loop

  get_PortHash = Chr$(Hex("&H" & Bin2Dec(biN2))) & Chr$(Hex("&H" & Bin2Dec(biN)))
  
End Function
'conv port to bin

Function Bin2Dec(binString As String) As Byte
Dim i As Integer, a As String, p As Byte, t As Byte

  For i = 8 To 1 Step -1
    a = Mid(binString, i, 1)
    t = t + a * 2 ^ p
    p = p + 1
  Next i

  Bin2Dec = t

End Function
'conv bin to dec

Private Sub winSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  modifyConnectedDisplay False
End Sub
