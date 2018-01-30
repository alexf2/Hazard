Attribute VB_Name = "ModMain"
Option Explicit

Public Const StreamStgGostsName As String = "StgGostsName"
Public Const MainGOSTStgName As String = "MainGosts.htg"

Public Enum EnNodes
  Node_GOST
  Node_Table
  Node_Empty
  Node_GOST_Ctx
End Enum


Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As Long, ByVal nCount As Long, ByVal lpDx As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Const ETO_OPAQUE = 2
Public Const TA_CENTER = 6
Public Const VTA_CENTER = TA_CENTER
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTextAlign Lib "gdi32" (ByVal hdc As Long) As Long

Public Const DT_CENTER = &H1
Public Const DT_NOPREFIX = &H800
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const TRANSPARENT = 1
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long



Public Const WM_CLOSE = &H10
Public Const WM_USER = &H400
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_PENDING_OP = WM_USER + 1
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const LB_ITEMFROMPOINT = &H1A9
Public Const LB_GETITEMRECT = &H198

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYHSCROLL = 3
Public Const SM_CXVSCROLL = 2
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

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



Public Sub HighLight2(ByVal btn As SSCommand, ByVal bEnter As Boolean, ByVal hWnd As Long)
  Dim bAct As Boolean
  bAct = (GetForegroundWindow() = hWnd)
  bAct = bAct And (bEnter Or btn.hWnd = GetFocus())
  'btn.Font3D = ssRaisedLight
  btn.Font3D = IIf(bAct, ssRaisedLight, ssNoneFont3D)
  btn.ForeColor = IIf(bAct, &HFF0000, 0)
End Sub

Public Function IsVarEmpty(ByRef v As Variant) As Boolean
  On Error GoTo ERR_H
  If IsEmpty(v) Then IsVarEmpty = True Else IsVarEmpty = False
  Exit Function
ERR_H:
  IsVarEmpty = True
  Err.Clear
End Function

Public Sub CenterForm(ByVal f1 As Form, ByVal f2 As Form)
    
  With f1
    .Move f2.Left - (.Width - f2.Width) / 2, f2.Top - (.Height - f2.Height) / 2
    
    If .Left < 0 Or .Top < 0 Or .Left + .Width > Screen.Width Or .Top + .Height > Screen.Height Then
      .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
    End If
  End With
  
End Sub
