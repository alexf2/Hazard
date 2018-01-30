Attribute VB_Name = "GNGlobals"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
Public Declare Function ScrollWindowEx Lib "user32" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As Any, lprcClip As Any, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function ValidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)


Public Const MM_ISOTROPIC = 7

Public Const SM_CXVSCROLL = 2
Public Const SM_CYVSCROLL = 20
Public Const SM_CXBORDER = 5



Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_INVALIDATE = &H2

Public Const SB_BOTTOM = 7
Public Const SB_ENDSCROLL = 8
Public Const SB_LINELEFT = 0
Public Const SB_LINERIGHT = 1
Public Const SB_PAGELEFT = 2
Public Const SB_PAGERIGHT = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6


Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Type SCROLLINFO
  cbSize As Long
  fMask As Long
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type

Public Const GWL_STYLE = (-16)
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000

Public Const SB_HORZ = 0
Public Const SB_VERT = 1

Public Const SIF_ALL = &H1 Or &H2 Or &H4 Or &H8 Or &H10

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_SHOWWINDOW = &H18


Public Type WPBnd
  hw As Long
  ref As IMsgHandl
  'ref As Long
  m_l_OldWndProc As Long
End Type


Private Const incrArrBnd = 10
Private arrBnd() As WPBnd
Private lArrBnd_UBound As Long

Public Sub CBOff()
  Dim lInd As Long
  For lInd = 1 To lArrBnd_UBound
    With arrBnd(lInd)
      If .hw <> 0 Then
        If .m_l_OldWndProc <> 0 Then _
          SetWindowLong .hw, GWL_WNDPROC, .m_l_OldWndProc: .m_l_OldWndProc = 0
        Set .ref = Nothing
        'ZeroMemory .ref, 4
        '.ref = 0
        .hw = 0
        Exit Sub
      End If
    End With
  Next lInd
End Sub

Public Sub BindCtl(ByVal hw As Long, ByVal ref As CtlRepeater)
  Dim lInd As Long
  For lInd = 1 To lArrBnd_UBound
    If arrBnd(lInd).hw = 0 Then Exit For
  Next lInd
  If lInd > lArrBnd_UBound Then
    lArrBnd_UBound = lArrBnd_UBound + incrArrBnd
    ReDim Preserve arrBnd(1 To lArrBnd_UBound) As WPBnd
  End If
  
  With arrBnd(lInd)
    .hw = hw
    Set .ref = ref
    'CopyMemory .ref, ref, 4
'    Dim imsg As IMsgHandl
'    Set imsg = ref
'    .ref = ObjPtr(imsg)
    .m_l_OldWndProc = SetWindowLong(hw, GWL_WNDPROC, AddressOf WindowProc)
  End With
End Sub

Public Sub UnbindCtl(ByVal hw As Long)
'MsgBox CStr(hw), vbOKOnly, "Unbind reqvest"
  Dim lInd As Long
  For lInd = 1 To lArrBnd_UBound
    With arrBnd(lInd)
      If .hw = hw Then
        If .m_l_OldWndProc <> 0 Then _
          SetWindowLong hw, GWL_WNDPROC, .m_l_OldWndProc: .m_l_OldWndProc = 0
        Set .ref = Nothing
        'ZeroMemory .ref, 4
        '.ref = 0
        .hw = 0
'MsgBox CStr(hw), vbOKOnly, "Unbind OK"
        Exit Sub
      End If
    End With
  Next lInd
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error GoTo ERR_H
  
  'Dim refShd0 As IMsgHandl
  'Dim refShd As IMsgHandl
  
  Dim lInd As Long
  For lInd = 1 To lArrBnd_UBound
    With arrBnd(lInd)
      If .hw = hw Then
        'Dim fl As Long
        'fl = .ref
        'CopyMemory refShd0, fl, 4
        'Set refShd = refShd0
        'ZeroMemory refShd0, 4
        'Set refShd = object(.ref)
        
        
        Select Case uMsg
          Case WM_HSCROLL
            WindowProc = .ref.HandleHScroll(LoLong(wParam), HiLong(wParam))
          Case WM_VSCROLL
            WindowProc = .ref.HandleVScroll(LoLong(wParam), HiLong(wParam))
          Case WM_SHOWWINDOW
            WindowProc = CallWindowProc(.m_l_OldWndProc, hw, uMsg, wParam, lParam)
            .ref.HandleSHW wParam, lParam
            
            'ZeroMemory refShd, 4
            'Set refShd = Nothing
            Exit Function
          Case Else
            WindowProc = CallWindowProc(.m_l_OldWndProc, hw, uMsg, wParam, lParam)
        End Select
        
        'ZeroMemory refShd, 4
        'Set refShd = Nothing
        Exit Function
      End If
    End With
  Next lInd
  
  WindowProc = 0
  'ZeroMemory refShd, 4
  'Set refShd = Nothing
  Exit Function
ERR_H:
  'ZeroMemory refShd, 4
  'Set refShd = Nothing
  MsgBox Err.Description, vbOKOnly, "Error: " & CStr(lInd) & ": hw=" & CStr(hw)
End Function

