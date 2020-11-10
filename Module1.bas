Attribute VB_Name = "Module1"
Option Explicit
Option Compare Binary

'################ TRAY CODE
Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Const NIM_ADD = &H0 'Add to Trayf
Public Const NIM_MODIFY = &H1 'Modify Detailsl
Public Const NIM_DELETE = &H2 'Remove From Traye
Public Const NIF_MESSAGE = &H1 'Messagee
Public Const NIF_ICON = &H2 'Iconbt
Public Const NIF_TIP = &H4 'TooTipTexty-
Public Const WM_MOUSEMOVE = &H200 'On Mousemove
'Public Const WM_LBUTTONDOWN = &H201 'Left Button Downf
'Public Const WM_LBUTTONUP = &H202 'Left Button Upl
Public Const WM_LBUTTONDBLCLK = &H203 'Left Double Clicke
'Public Const WM_RBUTTONDOWN = &H204 'Right Button Downe
'Public Const WM_RBUTTONUP = &H205 'Right Button Upt
'Public Const WM_RBUTTONDBLCLK = &H206 'Right Double Click-
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public notifID As NOTIFYICONDATA
'/################ TRAY CODE

'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'read/write from/to an INI file

Private Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'timer handling

Public proxyHTTP() As Boolean
Public proxyTicks() As Long
Public cancelNext As Boolean
Public proxyPort As Long
Public original_TimeoutTick As Long
Public proxyScanlist() As String, proxyScanlistIndex As Long
Public proxyRangesIndex As Long
Public methodType As Byte
'
Dim currentIP() As String
Dim socketQuarter As Integer, socketQuarterStart As Integer
'
'Public sss1 As Integer
Public overrideS4 As Boolean, overrideHTTP As Boolean
Public Sub Addproxy_Granted(proxyString As String)
  FrmMain.lblGrantedUse.Caption = CInt(FrmMain.lblGrantedUse.Caption) + 1
  printFile App.Path & "\Proxies\Proxies_Granted.txt", proxyString
End Sub
'adds onto granted proxies

Public Sub Addproxy_Denied(proxyString As String)
  FrmMain.lblDeniedRequest.Caption = CInt(FrmMain.lblDeniedRequest.Caption) + 1
  printFile App.Path & "\Proxies\Proxies_Denied.txt", proxyString
End Sub
'adds onto denied

Public Sub Addproxy_Bnet(proxyString As String)
  FrmMain.lblBnetVerified.Caption = CInt(FrmMain.lblBnetVerified.Caption) + 1
  FrmMain.lblGrantedUse.Caption = CInt(FrmMain.lblGrantedUse.Caption) - 1
  printFile App.Path & "\Proxies\Proxies_Verified.txt", proxyString
End Sub
'adds onto bnet verified


Public Function GTC() As Long
Static lExtraTick As Long

  GTC = GetTickCount
  
  If GTC < 0 Then If lExtraTick = 0 Then lExtraTick = Mid$(Val#(GTC), 2)

  GTC = GTC + lExtraTick

End Function


Public Sub set_currentIP(IPString As String)
  
  If CInt(FrmMain.txtScanThreads.Text) >= 3 Then
    socketQuarter = (CInt(FrmMain.txtScanThreads.Text) / 3)
    If InStrB(socketQuarter, ".") Then socketQuarter = Left$(socketQuarter, InStr(socketQuarter, ".") - 1)
    socketQuarter = socketQuarter + 1
  Else
    socketQuarter = 3
  End If
  socketQuarterStart = 0
  'sets scan check timer and how many ips to scan during these burst
  
  proxyPort = CLng(FrmMain.txtRangePort.Text)
  currentIP = Split(IPString, ".")

End Sub
'sets currentIP, starting ip currently

Public Sub Load_Ranges()
Dim newFile As Integer, fileData As String
Dim rangeSet() As String
Dim rangeType As String

  proxyRangesIndex = 0
  '
  DoEvents
  newFile = FreeFile
  Open App.Path & "\Ranges.txt" For Input As #newFile
  Do Until EOF(newFile)
  Line Input #newFile, fileData
    If InStrB(fileData, ChrW$(32)) And InStrB(fileData, ".") Then
      rangeType = vbNullString
      If InStrB(fileData, "@") Then
        rangeType = Mid$(fileData, InStr(fileData, "@"))
        fileData = Left$(fileData, InStr(fileData, "@") - 1)
      End If
      'sets if http or socks
      rangeSet = Split(fileData, ChrW$(32))
      If UBound(rangeSet) = 1 Then
        rangeSet(0) = Trim$(rangeSet(0))
        rangeSet(1) = Trim$(rangeSet(1))
        If InStrB(rangeSet(0), ":") Then rangeSet(0) = Left$(rangeSet(0), InStr(rangeSet(0), ":") - 1)
        'removes : from start of range if they included it
        If InStrB(rangeSet(1), ":") = 0 Then rangeSet(1) = rangeSet(1) & ":1080"
        'missing port assume 1080
        fileData = rangeSet(0) & "   " & rangeSet(1)
        'display purposes
        fileData = fileData & rangeType
        FrmMain.lstRangeSelection.AddItem fileData
      'if start and end supplied
      End If
    End If
  Loop
  Close #newFile
  'done loading
  If FrmMain.lstRangeSelection.ListCount <> 0 Then FrmMain.lstRangeSelection.ListIndex = 0
  FrmMain.txtScanProgress.Text = "Ranges Set! " & Time$
  
End Sub


Public Sub Load_Scanlist()
Dim newFile As Integer, fileData As String
  
  proxyScanlistIndex = 0
  ReDim proxyScanlist(0)
  '
  DoEvents
  newFile = FreeFile
  Open App.Path & "\Scanlist.txt" For Input As #newFile
  Do Until EOF(newFile)
  Line Input #newFile, fileData
    If InStrB(fileData, ".") Then
      fileData = LTrim$(fileData)
      If InStrB(fileData, ChrW$(32)) Then fileData = Left$(fileData, InStr(fileData, ChrW$(32)) - 1)
      'removes anything after a space
      If InStrB(fileData, ":") = 0 Then fileData = fileData & ":1080"
      'missing port assume 1080
      proxyScanlist(UBound(proxyScanlist)) = fileData
      ReDim Preserve proxyScanlist(UBound(proxyScanlist) + 1)
    End If
  Loop
  Close #newFile
  'done loading
  FrmMain.txtScanProgress.Text = "Scanlist Set! " & Time$
  
End Sub

Public Function set_NewRange() As Boolean
On Error GoTo Err
  
  overrideS4 = False
  overrideHTTP = False
  'if set true then always scan as that type
  If proxyRangesIndex = FrmMain.lstRangeSelection.ListCount Then
    set_NewRange = False
  Else
    Dim rangeSplit() As String, theRange As String
    theRange = FrmMain.lstRangeSelection.List(proxyRangesIndex)
    If InStrB(theRange, "@") Then
        If Right$(LCase$(theRange), 4) = "http" Then
            overrideHTTP = True
        Else
            overrideS4 = True
        End If
        'set overridetype
        theRange = Left$(theRange, InStr(theRange, "@") - 1)
    End If
    'set range type if specified
    rangeSplit = Split(theRange, "   ")
    '( 0 ) is start ( 1 ) is end : portet
    FrmMain.txtRangeStart.Text = rangeSplit(0)
    FrmMain.txtRangePort.Text = Mid$(rangeSplit(1), InStr(rangeSplit(1), ":") + 1)
    FrmMain.txtRangeEnd.Text = Left$(rangeSplit(1), InStr(rangeSplit(1), ":") - 1)
    'sets displays as if it were a manual scanle
    set_currentIP FrmMain.txtRangeStart.Text
    set_NewRange = True
    FrmMain.lstRangeSelection.ListIndex = proxyRangesIndex
  End If

Err:
  If Err.Number <> 0 Then FrmMain.txtScanProgress.Text = "Ranges(NR) Error! " & Time$
End Function
'sets next range in line or returns false if end of rangesf

Public Function get_NextIP(ByVal Index As Integer) As String
Static raiseFactor As Byte
On Error GoTo Err

  raiseFactor = raiseFactor + 1
  
  Select Case methodType
    Case 0, 2

get_NextIP_RangeTOP:

      currentIP(3) = currentIP(3) + 1
      If currentIP(3) > 255 Then
        currentIP(3) = 0
        currentIP(2) = currentIP(2) + 1
      End If
      If currentIP(2) > 255 Then
        currentIP(2) = 0
        currentIP(1) = currentIP(1) + 1
      End If
      If currentIP(1) > 255 Then
        currentIP(1) = 0
        currentIP(0) = currentIP(0) + 1
      End If
      'inc by 1
      If currentIP(0) = 256 Then
      'do nothing, null get_nextip = the end
      Else
        get_NextIP = Join(currentIP, ".")
      'sets next
      End If
      '
      If get_NextIP = FrmMain.txtRangeEnd.Text Then
        If methodType = 0 Then
          FrmMain.cmdScanStop.Caption = "Stop Scan"
          FrmMain.txtScanProgress.Text = "Finished"
          FrmMain.cmdScanStop_Click
        'manual scan, finished
        Else
          proxyRangesIndex = proxyRangesIndex + 1
          If Not set_NewRange Then
            FrmMain.cmdScanStop.Caption = "Stop Scan"
            FrmMain.txtScanProgress.Text = "Finished"
            FrmMain.cmdScanStop_Click
            Exit Function
          'ranges complete
          Else
           GoTo get_NextIP_RangeTOP
         'new range set fine, goto top
          End If
        'rangelist scan check if new range avail
        End If
      'reached end of range, set new scan range or end if manual
      End If
      '
      If raiseFactor >= 175 Then
        raiseFactor = 0
        FrmMain.txtScanProgress.Text = get_NextIP
      'for display
      End If
        
    'manual scanning method
    Case 1
      Dim inLoc As Integer
      get_NextIP = proxyScanlist(proxyScanlistIndex)
      inLoc = InStr(LCase$(get_NextIP), "@")
      If inLoc <> 0 Then
        If Len(get_NextIP) - inLoc < 3 Then
          proxyHTTP(Index) = False
        Else
          proxyHTTP(Index) = True
        End If
        get_NextIP = Left$(get_NextIP, inLoc - 1)
     End If
     'sets if to use http or not
      If InStrB(get_NextIP, ":") Then
        proxyPort = Mid$(get_NextIP, InStr(get_NextIP, ":") + 1)
        get_NextIP = Left$(get_NextIP, InStr(get_NextIP, ":") - 1)
      End If
      'sets next
      
      If raiseFactor >= 10 Then
        raiseFactor = 0
        FrmMain.txtScanProgress.Text = proxyScanlistIndex & " / " & UBound(proxyScanlist)
      'for display
      End If
      
      proxyScanlistIndex = proxyScanlistIndex + 1
      If proxyScanlistIndex = UBound(proxyScanlist) Then
        FrmMain.cmdScanStop.Caption = "Stop Scan"
        FrmMain.txtScanProgress.Text = "Finished"
        FrmMain.cmdScanStop_Click
      'scanlist complete
      End If
      
      
    'scanlist
  End Select
      
Err:
  If Err.Number <> 0 Then FrmMain.txtScanProgress.Text = "NextIP Error! " & Time$
End Function
'grabs next ip to scan

Public Sub connect_ProxySet()
Static subPasses As Byte
Dim l_tickCount As Long
Dim threadIndex As Integer
Dim sectionCount As Integer

On Error GoTo Err

  KillTimer FrmMain.hwnd, 0
  '
  subPasses = subPasses + 1
  If subPasses = 100 Then
    subPasses = 0
    FrmMain.txtScanTimeout.Text = original_TimeoutTick
    '
    If Not FrmMain.Visible Then
      notifID.szTip = "BetaScanner - Scanning... V:" & FrmMain.lblBnetVerified & ", G:" & FrmMain.lblGrantedUse & vbNullChar 'tooltip text
      Shell_NotifyIcon NIM_MODIFY, notifID
    End If
    'update tray
  End If
  'reset timeout timer every now and then incase it was increased from no buffer space error

  l_tickCount = GTC
  '
  For threadIndex = socketQuarterStart To (FrmMain.winSock.Count - 1)
    If cancelNext Then
      If Not FrmMain.Visible Then
        notifID.hIcon = FrmMain.icoFinished.Picture
        notifID.szTip = "BetaScanner - Complete! V:" & FrmMain.lblBnetVerified & ", G:" & FrmMain.lblGrantedUse & vbNullChar 'tooltip text
        Shell_NotifyIcon NIM_MODIFY, notifID
      End If
      'icon update
      Exit Sub
    End If
    'scanning complete and or stopped
    If l_tickCount - proxyTicks(threadIndex) > FrmMain.txtScanTimeout Then
      proxyTicks(threadIndex) = l_tickCount
      'set new timeout
      'sss1 = sss1 + 1
      DoEvents
      FrmMain.winSock(threadIndex).Close
      FrmMain.winSock(threadIndex).Connect get_NextIP(threadIndex), proxyPort
    End If
    '
    sectionCount = sectionCount + 1
    If sectionCount = socketQuarter Then GoTo Err
    'thread quota met
  Next threadIndex
  'connects timed out sockets to next ip in range

Err:
  socketQuarterStart = socketQuarterStart + socketQuarter
  If socketQuarterStart > (FrmMain.winSock.Count - 1) Then socketQuarterStart = 0
  'raise start point for next burst, or reset if at end
  If Err.Number = 10055 Then
    If Val#(FrmMain.txtScanTimeout.Text) <= (original_TimeoutTick * 4) Then
      Dim newTick As String
      newTick = Val#(FrmMain.txtScanTimeout.Text * 1.3)
      If InStrB(newTick, ".") Then newTick = Left$(newTick, InStr(newTick, ".") - 1)
      FrmMain.txtScanTimeout.Text = newTick
    End If
  End If
  SetTimer FrmMain.hwnd, 0, 400, AddressOf connect_ProxySet

End Sub
'controls when sockets connect to their next ip, and when ip dest is reached



Public Function GetINI(sKey As String) As String
Dim Ret As String, NC As Long

  Ret = String(200, 0)
  NC = GetPrivateProfileString("BETA Scanner", sKey, sKey, Ret, 200, App.Path & "\BETAScannerINI.INI")
  If NC <> 0 Then Ret = Left$(Ret, NC)
  
  If Ret = sKey Or Len(Ret) = 200 Then Ret = vbNullString
  
  GetINI = Ret

End Function
'Read from INI

Public Sub WriteINI(sKey As String, sValue As String)
  
  If LenB(FrmMain.txtScanProgress.Text) = 0 Then Exit Sub
  'stops from writing after loading(initial load)
  WritePrivateProfileString "BETA Scanner", sKey, sValue, App.Path & "\BETAScannerINI.INI"

End Sub
'Write to INI

Private Sub printFile(fileName As String, fileData As String)
Dim newFile As Integer

  DoEvents
  newFile = FreeFile
  Open fileName For Append As #newFile
  Print #newFile, fileData
  Close #newFile

End Sub
'prints to a text file

'Public Sub checkChecked()
'Dim percentResult As String

'  percentResult = sss1 / Val#(FrmMain.txtScanThreads.Text)
'
'  Dim inStt As Integer
'  inStt = InStr(percentResult, ".")
'  Select Case inStt
'    Case 0
'      percentResult = "100"
'    Case 2
'      percentResult = Mid$(percentResult, 3, 2)
'      If Len(percentResult) = 1 Then percentResult = percentResult & "0"
'  End Select
'  printFile App.Path & "\testPercent.txt", percentResult & "% Efficiency"
'
'  FrmMain.Caption = percentResult & "%  " & sss1
'  sss1 = 0
'
'End Sub
'determins performance of scanning, how close to threads per rate.
