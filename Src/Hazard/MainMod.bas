Attribute VB_Name = "MainMod"
Option Explicit

Public Enum AdvDirty
  AD_None
  Ad_Updated
  Ad_New
  Ad_Deleted
End Enum

Public Const MAX_PATH = 260

Public Const US_SP_Calc1 = 1
Public Const US_SP_Calc2 = 2

Public Const S_RegKey = "Hazard.Document"
Public Const S_RegName = "Модель ЧМС Hazard 2000/3.0"

Public Const S_RegAppKey = "SOFTWARE\AlexCorp\Hazard 2000\3.00\"

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageRect Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As RECT) As Long
Public Const WM_ACTIVATE = &H6
Public Const WA_CLICKACTIVE = 2
Public Const WM_SETTEXT = &HC
Public Const TVM_GETITEMRECT = 4356

Public Const LB_SETHORIZONTALEXTENT = &H194

Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)

Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYHSCROLL = 3
Public Const SM_CXVSCROLL = 2
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


Public Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type


Public Declare Function LoadModule Lib "kernel32" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long



Public haApp As HazardApp

Public frmCalibrate As CfrmCalibrate
Public frmMEditor As CfrmMEditor
Public frmMon As CfrmMon
'Public frmProcs As CfrmProcs

Public frmComplSP As CfrmComplSP
Public frmOptimiz As CfrmOptimiz
Public frmSP As CfrmSP

'Dim wh As New WIN32HOOKLib.Win32Hk

Public Sub Main()
 'MsgBox IIf(App.StartMode = vbSModeAutomation, "Auto", "Alone")
 'MsgBox IIf(haApp Is Nothing, "Nothing", "No Nothing")
  
 'LoadLibrary "f:\OTX\OtxShortcutBar.dll"

 'wh.AddHook "f:\OTX\OtxShortcutBar.dll"
 'wh.AddHook "g:\work\mag\alexf\Hazard\hazard.exe"
 'wh.StartLog "g:\OtxShortcutBar.hlog"
 
 If App.StartMode = vbSModeStandalone Then
   Set haApp = New HazardApp
 End If
 
 'wh.StopLog
 'wh.RemoveHook
End Sub


Public Sub SortCollX(ByVal coll As CollectionX)
  Dim i As Long, j As Long, i2 As Long
  i2 = coll.Count
  For i = 1 To i2 - 1
    For j = i + 1 To i2
      If coll(i) < coll(j) Then
        Dim lTmp As Long
        lTmp = coll(i)
        coll(i) = coll(j)
        coll(j) = lTmp
      End If
    Next
  Next
End Sub

Public Sub ExpandAll(ByVal pvTree As PVTreeViewX)
  Dim node As PVBranch

  Set node = pvTree.Branches.Get(pvtGetChild, 0)
  While node.IsValid
      node.Open pvtAllChildren
      Set node = node.Get(pvtGetNextSibling, 0)
  Wend
End Sub

Public Function TSign(ByVal i As Integer)
  TSign = IIf(i < 0, "-", "+")
End Function


Public Function FindL(ByVal lVal As Long, arrL() As Long) As Boolean
  Dim l As Long
  FindL = False
  For l = LBound(arrL) To UBound(arrL)
    If arrL(l) = lVal Then
      FindL = True
      Exit Function
    End If
  Next l
End Function

Public Function CountSelected(ByVal pv As PVTreeViewX) As Long
  Dim lCnt As Long: lCnt = 0
  Dim node As PVBranch, node2 As PVBranch

  Set node = pv.Branches.Get(pvtGetChild, 0)
  While node.IsValid
    Set node2 = node.Get(pvtGetChild, 0)
    While node2.IsValid
      If node2.IsSelected Then lCnt = lCnt + 1
      Set node2 = node2.Get(pvtGetNextSibling, 0)
    Wend
    Set node = node.Get(pvtGetNextSibling, 0)
  Wend
  
  CountSelected = lCnt
End Function


Public Sub HighLight(ByVal btn As SSCommand, ByVal bEnter As Boolean)
  'btn.Font.Bold = bEnter
  Dim bAct As Boolean
  bAct = (GetForegroundWindow() = frmMain.hWnd)
  bAct = bAct And (bEnter Or btn.hWnd = GetFocus())
  'btn.Font3D = ssRaisedLight
  btn.Font3D = IIf(bAct, ssRaisedLight, ssNoneFont3D)
  btn.ForeColor = IIf(bAct, &HFF0000, 0)
End Sub

Public Sub HighLight2(ByVal btn As SSCommand, ByVal bEnter As Boolean, ByVal hWnd As Long)
  Dim bAct As Boolean
  bAct = (GetForegroundWindow() = hWnd)
  bAct = bAct And (bEnter Or btn.hWnd = GetFocus())
  'btn.Font3D = ssRaisedLight
  btn.Font3D = IIf(bAct, ssRaisedLight, ssNoneFont3D)
  btn.ForeColor = IIf(bAct, &HFF0000, 0)
End Sub

Public Sub RegisterShellext(ByVal bFlCheck As Boolean)
  If bFlCheck And "" = GetClassesSetting(".hzd", "") Or Not bFlCheck Then
      Dim sExe As String
      sExe = Qq(App.Path & "\" & App.EXEName & ".exe") ' full path name of .exe
      SaveClassesSetting ".hzd", "", S_RegKey
      SaveClassesSetting S_RegKey, "", S_RegName
      SaveClassesSetting S_RegKey, "DefaultIcon", sExe & ",0"
      SaveClassesSetting S_RegKey & "\shell", "", ""
      SaveClassesSetting S_RegKey & "\shell\new", "", "Новая модель"
      SaveClassesSetting S_RegKey & "\shell\new\command", "", sExe & " /n "
      SaveClassesSetting S_RegKey & "\shell\open", "", "Открыть модель"
      SaveClassesSetting S_RegKey & "\shell\open\command", "", sExe & " /o " & "%1"
  End If
End Sub

Public Sub UnregisterShellext()
Dim lr As Long
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, ".hzd")
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey & "\shell\open\command")
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey & "\shell\open")
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey & "\shell\new\command")
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey & "\shell\new")
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey & "\shell")
  
  lr = ERROR_SUCCESS
  
  Dim rc As Long, hKey As Long
  rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, S_RegKey, 0, KEY_ALL_ACCESS, hKey)
  If rc <> ERROR_SUCCESS Then hKey = 0
  
  If hKey <> 0 Then
     lr = RegDeleteKey(hKey, "DefaultIcon")
     lr = RegCloseKey(hKey)
  End If
  
  lr = RegDeleteKey(HKEY_CLASSES_ROOT, S_RegKey)
  
End Sub

Private Function Qq(ByVal s As String) As String
  Qq = Chr(34) & s & Chr(34)
End Function

Public Function StrFromTCH(ByVal tch As TrustChange, Optional ByVal bFlSeq As Boolean = True) As String
  Select Case tch
    Case TR_NoChange
      StrFromTCH = ""
    Case TR_MinusOne
      StrFromTCH = IIf(bFlSeq, ",-1", "-1")
    Case TR_MinusTwo
      StrFromTCH = IIf(bFlSeq, ",-2", "-2")
    Case TR_PlusOne
      StrFromTCH = IIf(bFlSeq, ",+1", "+1")
    Case TR_PlusTwo
      StrFromTCH = IIf(bFlSeq, ",+2", "+2")
    Case TR_SetLow
      StrFromTCH = IIf(bFlSeq, ",Низкий", "Низкий")
    Case TR_SetNormal
      StrFromTCH = IIf(bFlSeq, ",Средний", "Средний")
    Case TR_SetHigh
      StrFromTCH = IIf(bFlSeq, ",Высокий", "Высокий")
    Case Else
      StrFromTCH = ""
  End Select
End Function

Public Function StrFromTrust(ByVal tch As TrustLevel) As String
  Select Case tch
    Case TL_Low
      StrFromTrust = "Низкий"
    Case TL_Normal
      StrFromTrust = "Средний"
    Case TL_High
      StrFromTrust = "Высокий"
    Case Else
      StrFromTrust = "<>"
  End Select
End Function

Public Sub SetupPVNumeric(ByVal nf As GERTNETLib.NumberFormat, ByVal NumberDigits As Integer, ParamArray CtlArray())
  If IsMissing(CtlArray) Then Exit Sub
  Dim pv As PVNumeric, l As Long
  If nf = NF_Normal Then
    For l = LBound(CtlArray) To UBound(CtlArray)
      Set pv = CtlArray(l)
      With pv
        .Format = pvnNormal
        .LimitValueByType = pvnCustom
        .DecimalMin = 0
        .DecimalMax = 15
      End With
    Next l
  Else
    For l = LBound(CtlArray) To UBound(CtlArray)
      Set pv = CtlArray(l)
      With pv
        .Format = pvnPower
        .LimitValueByType = pvnCustom
        .DecimalMin = 0
        .DecimalMax = 15
      End With
    Next l
  End If
End Sub

Public Sub RemoveCollItems(ByVal coll As CollectionX, ByVal collKeys As CollectionX)
  Dim vvt
  On Error Resume Next
  For Each vvt In collKeys
    coll.Remove vvt
  Next vvt
End Sub

Public Function UniqueKey(ByVal coll As CollectionX) As String
  
  Dim sTmp As String: sTmp = CStr(Fix(Rnd() * 9))
  Do While Not coll.Item(sTmp) Is coll.DefaultItem
    sTmp = sTmp & CStr(Fix(Rnd() * 9))
  Loop
  
  UniqueKey = sTmp
End Function

Public Function GetResolveReady(ByVal coll As CollectionX) As CollectionX
  Static arrSrt(3) As AdvDirty
  Static iFlInit As Integer
  If iFlInit = 0 Then
    iFlInit = 1
    arrSrt(0) = Ad_Deleted
    arrSrt(1) = Ad_Updated
    arrSrt(2) = Ad_New
    arrSrt(3) = AD_None
  End If
  
  Dim cllRet As CollectionX: Set cllRet = New CollectionX
  
  Dim l As Long, lC As Long: lC = coll.Count
  Dim l2 As Long, adCurr As AdvDirty
  For l = 0 To 3
    adCurr = arrSrt(l)
    For l2 = 1 To lC
      With coll.Item(l2)
        If .Dirty = adCurr Then
          cllRet.Add coll.Item(l2), coll.Key(l2)
        End If
      End With
    Next l2
  Next l
  
  Set GetResolveReady = cllRet
  Set cllRet = Nothing
  
End Function
