VERSION 5.00
Begin VB.UserControl Bulb 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipBehavior    =   0  'None
   HitBehavior     =   0  'None
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   34
   ToolboxBitmap   =   "bulb.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "Bulb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Enums
Public Enum Mode_Constants
  mdRound = 0
  mdSQuare = 1
End Enum

Public Enum State_Constants
  stOn = 0
  stOff = 1
End Enum

'Default Property Values:
Const m_def_ColorHi = &HFF&
Const m_def_ColorLo = &H80&
Const m_def_ColorLight = &HC0C0C0
Const m_def_ColorDark = &H808080
Const m_def_DrawMode = mdRound
Const m_def_AutoResize = False
Const m_def_CantWidth = 1
Const m_def_TimeOff = 500
Const m_def_TimeOn = 100

'Internal constants
Const m_percRounding = 15#
Const m_percLeftBottom = 20#
Const m_percRightTop = 20#

'Property Variables:
Private m_ColorHi As OLE_COLOR
Private m_ColorLo As OLE_COLOR
Private m_ColorLight As OLE_COLOR
Private m_ColorDark As OLE_COLOR
Private m_md_DrawMode As Mode_Constants
Private m_b_AutoResize As Boolean
Private m_i_CantWidth As Integer
Private m_l_TimeOff As Long
Private m_l_TimeOn As Long

'Internal Variables
Private mlHRGN As Long
Private mbrHi As Long
Private mbrLo As Long

Private mpenLight As Long
Private mbrDark As Long
Private m_bState As Boolean
Private m_rArc As RECT

Private m_bLockResize As Boolean

Private WithEvents m_xt1 As XTimer
Attribute m_xt1.VB_VarHelpID = -1
Private WithEvents m_xt2 As XTimer
Attribute m_xt2.VB_VarHelpID = -1

'Event Declarations:
Event Click(ByVal aocxSender As Bulb)
Attribute Click.VB_UserMemId = -600

Public Property Let TimeOff(ByVal tm As Long)
  m_l_TimeOff = tm
  m_xt1.Interval = tm
  PropertyChanged "TimeOff"
End Property

Public Property Get TimeOff() As Long
Attribute TimeOff.VB_ProcData.VB_Invoke_Property = ";Behavior"
  TimeOff = m_l_TimeOff
End Property

Public Property Let TimeOn(ByVal tm As Long)
  m_l_TimeOn = tm
  m_xt2.Interval = tm
  PropertyChanged "TimeOn"
End Property

Public Property Get TimeOn() As Long
Attribute TimeOn.VB_ProcData.VB_Invoke_Property = ";Behavior"
  TimeOn = m_l_TimeOn
End Property


Public Property Let State(ByVal st As State_Constants)
  m_bState = IIf(st = stOff, False, True)
  RepaintControl
End Property

Public Property Get State() As State_Constants
Attribute State.VB_ProcData.VB_Invoke_Property = ";Misc"
  State = IIf(m_bState = False, stOff, stOn)
End Property


Public Property Let CantWidth(ByVal iW As Integer)
  If iW < 0 Then
    Err.Raise vbObjectError, "", "Bad argument"
    Exit Property
  End If
  m_i_CantWidth = iW
  Recreate_Pen mpenLight, m_ColorLight
  PropertyChanged "CantWidth"
  RepaintControl
End Property
Public Property Get CantWidth() As Integer
Attribute CantWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
  CantWidth = m_i_CantWidth
End Property

Public Property Let AutoResize(ByVal bRsz As Boolean)
  m_b_AutoResize = bRsz
  PropertyChanged "AutoResize"
  UserControl_Resize
  RepaintControl
End Property

Public Property Get AutoResize() As Boolean
Attribute AutoResize.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute AutoResize.VB_UserMemId = -500
  AutoResize = m_b_AutoResize
End Property


Public Property Let DrawMode(ByVal drwMd As Mode_Constants)
  m_md_DrawMode = drwMd
  PropertyChanged "DrawMode"
  UserControl_Resize
  RepaintControl
End Property
Public Property Get DrawMode() As Mode_Constants
Attribute DrawMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
  DrawMode = m_md_DrawMode
End Property

Public Property Let ColorLight(ByVal clrLi As OLE_COLOR)
  m_ColorLight = clrLi
  Recreate_Pen mpenLight, m_ColorLight
  PropertyChanged "ColorLight"
  RepaintControl
End Property

Public Property Get ColorLight() As OLE_COLOR
Attribute ColorLight.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ColorLight = m_ColorLight
End Property

Public Property Let ColorDark(ByVal clrDk As OLE_COLOR)
  m_ColorDark = clrDk
  Recreate_Brush mbrDark, m_ColorDark
  PropertyChanged "ColorDark"
  RepaintControl
End Property

Public Property Get ColorDark() As OLE_COLOR
Attribute ColorDark.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ColorDark = m_ColorDark
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
  Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
  UserControl.Appearance = New_Appearance
  PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
  UserControl.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H000000FF&
Public Property Get ColorHi() As OLE_COLOR
Attribute ColorHi.VB_Description = "Color pulse"
Attribute ColorHi.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ColorHi = m_ColorHi
End Property

Public Property Let ColorHi(ByVal New_ColorHi As OLE_COLOR)
  m_ColorHi = New_ColorHi
  Recreate_Brush mbrHi, m_ColorHi
  PropertyChanged "ColorHi"
  RepaintControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00000080&
Public Property Get ColorLo() As OLE_COLOR
Attribute ColorLo.VB_Description = "Color off"
Attribute ColorLo.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ColorLo = m_ColorLo
End Property

Public Property Let ColorLo(ByVal New_ColorLo As OLE_COLOR)
  m_ColorLo = New_ColorLo
  Recreate_Brush mbrLo, m_ColorLo
  PropertyChanged "ColorLo"
  RepaintControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub PulseStart()
  State = stOff
  
  m_xt1.Interval = m_l_TimeOff
  m_xt2.Interval = m_l_TimeOn
  m_xt1.Enabled = True
  
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub PulseStop()
  State = stOff
  m_xt1.Enabled = False
  m_xt2.Enabled = False
End Sub

Private Sub m_xt1_Tick()
  If m_l_TimeOff <> m_l_TimeOn Then
    State = stOn
    m_xt1.Enabled = False
    m_xt2.Enabled = True
  Else
    State = IIf(State = stOff, stOn, stOff)
  End If
End Sub

Private Sub m_xt2_Tick()
  State = stOff
  m_xt2.Enabled = False
  m_xt1.Enabled = True
End Sub

Private Sub UserControl_Initialize()
  m_bLockResize = False
  mlHRGN = 0
  mbrHi = 0
  mbrLo = 0

  mpenLight = 0
  mbrDark = 0
  m_bState = False
  
  Set m_xt1 = New XTimer
  Set m_xt2 = New XTimer
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ColorHi = m_def_ColorHi
  m_ColorLo = m_def_ColorLo
  m_ColorLight = m_def_ColorLight
  m_ColorDark = m_def_ColorDark
  m_md_DrawMode = m_def_DrawMode
  m_b_AutoResize = m_def_AutoResize
  m_i_CantWidth = m_def_CantWidth
  m_l_TimeOff = m_def_TimeOff
  m_l_TimeOn = m_def_TimeOn
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click(Me)
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
  m_ColorHi = PropBag.ReadProperty("ColorHi", m_def_ColorHi)
  m_ColorLo = PropBag.ReadProperty("ColorLo", m_def_ColorLo)
  
  m_ColorLight = PropBag.ReadProperty("ColorLight", m_def_ColorLight)
  m_ColorDark = PropBag.ReadProperty("ColorDark", m_def_ColorDark)
  
  m_md_DrawMode = PropBag.ReadProperty("DrawMode", m_def_DrawMode)
  m_b_AutoResize = PropBag.ReadProperty("AutoResize", m_def_AutoResize)
  
  m_i_CantWidth = PropBag.ReadProperty("CantWidth", m_def_CantWidth)
  m_l_TimeOff = PropBag.ReadProperty("TimeOff", m_def_TimeOff)
  m_l_TimeOn = PropBag.ReadProperty("TimeOn", m_def_TimeOn)
End Sub

Private Sub UserControl_Terminate()
  m_xt1.Enabled = False
  m_xt2.Enabled = False
  Set m_xt1 = Nothing
  Set m_xt2 = Nothing
  
  DestroyRegion mlHRGN
    
  DestroyBrush mbrHi
  DestroyBrush mbrLo
  DestroyPen mpenLight
  DestroyBrush mbrDark
  
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
  Call PropBag.WriteProperty("ColorHi", m_ColorHi, m_def_ColorHi)
  Call PropBag.WriteProperty("ColorLo", m_ColorLo, m_def_ColorLo)
  
  PropBag.WriteProperty "ColorLight", m_ColorLight, m_def_ColorLight
  PropBag.WriteProperty "ColorDark", m_ColorDark, m_def_ColorDark
  
  PropBag.WriteProperty "DrawMode", m_md_DrawMode, m_def_DrawMode
  PropBag.WriteProperty "AutoResize", m_b_AutoResize, m_def_AutoResize
  PropBag.WriteProperty "CantWidth", m_i_CantWidth, m_def_CantWidth
  PropBag.WriteProperty "TimeOff", m_l_TimeOff, m_def_TimeOff
  PropBag.WriteProperty "TimeOn", m_l_TimeOn, m_def_TimeOn
End Sub


Private Sub RepaintControl()
'   ****************************************************************************
'   * Purpose:  To force a repaint of the control.
'   ****************************************************************************
    
  UserControl.Refresh
    
End Sub

Private Sub UserControl_Resize()
  If Not m_bLockResize Then
    m_bLockResize = True
    If m_b_AutoResize And ScaleWidth <> ScaleHeight Then
      Dim sQ As Single
      sQ = IIf(Extender.Width > Extender.Height, Extender.Width, Extender.Height)
      
'      Extender.Width = sQ
'      Extender.Height = sQ
      'Extender.Move Width:=sQ, Height:=sQ
      Extender.Move Extender.Left, Extender.Top, sQ, sQ
    End If
    
    RecreateRegions
    
    m_bLockResize = False
  End If
End Sub

Private Sub UserControl_Paint()
  If mbrLo = 0 Then Recreate_Brush mbrLo, m_ColorLo
  If mbrHi = 0 Then Recreate_Brush mbrHi, m_ColorHi
  If mbrDark = 0 Then Recreate_Brush mbrDark, m_ColorDark
  
  If mpenLight = 0 Then Recreate_Pen mpenLight, m_ColorLight
  

  Dim rRect As RECT, lRgnClip As Long, lRgnRes As Long
  GetClipBox hdc, rRect
'  If IsRectEmpty(rRect) <> 0 Then
'    rRect.Left = ScaleLeft: rRect.Right = rRect.Left + ScaleWidth
'    rRect.Top = ScaleTop: rRect.Bottom = rRect.Top + ScaleHeight
'  End If
  lRgnClip = CreateRectRgnIndirect(rRect)
  lRgnRes = CreateRectRgnIndirect(rRect)
    
  CombineRgn lRgnRes, mlHRGN, lRgnClip, RGN_AND
    
  Dim lbrOld As Long
  lbrOld = SelectObject(hdc, IIf(m_bState = True, mbrHi, mbrLo))
  
  PaintRgn hdc, lRgnRes
  FrameRgn hdc, mlHRGN, mbrDark, m_i_CantWidth, m_i_CantWidth
  
  Dim lpenOld As Long
  lpenOld = SelectObject(hdc, mpenLight)
  If m_md_DrawMode = mdRound Then _
    Arc hdc, ScaleLeft + 2 * m_i_CantWidth, ScaleTop + 2 * m_i_CantWidth, ScaleLeft + ScaleWidth - 2 * m_i_CantWidth, ScaleTop + ScaleHeight - 2 * m_i_CantWidth, _
      m_rArc.Left, m_rArc.Top, m_rArc.Right, m_rArc.Bottom
  'kkkkkkkkkkkkk
  SelectObject hdc, lpenOld
  SelectObject hdc, lbrOld
  DeleteObject lRgnClip
  DeleteObject lRgnRes
  
End Sub


Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
  If PtInRegion(mlHRGN, X, Y) = 1 Then _
    HitResult = vbHitResultHit: Exit Sub
    
  Dim r As RECT
  r.Left = ScaleLeft: r.Right = r.Left + ScaleWidth
  r.Top = ScaleTop: r.Bottom = r.Top + ScaleHeight
  
  If X >= r.Left And X <= r.Right And Y >= r.Top And Y <= r.Bottom Then
    HitResult = vbHitResultTransparent
    Exit Sub
  End If
    
  HitResult = vbHitResultOutside
End Sub

Private Sub RecreateRegions()
  DestroyRegion mlHRGN
    
  If m_md_DrawMode = mdRound Then
    Dim rRect As RECT
    rRect.Left = ScaleLeft: rRect.Top = ScaleTop
    rRect.Right = ScaleLeft + ScaleWidth: rRect.Bottom = ScaleTop + ScaleHeight
    mlHRGN = CreateEllipticRgnIndirect(rRect)
    'lllllllllllll
    
    m_rArc.Left = rRect.Left + ScaleWidth / 100# * m_percLeftBottom
    m_rArc.Top = rRect.Bottom
    m_rArc.Right = rRect.Right
    m_rArc.Bottom = rRect.Top + ScaleHeight / 100# * m_percRightTop
  ElseIf m_md_DrawMode = mdSQuare Then
    Dim sRndX As Single, sRndY As Single
    sRndX = CSng(ScaleWidth) / 100# * m_percRounding
    sRndY = CSng(ScaleHeight) / 100# * m_percRounding
    mlHRGN = CreateRoundRectRgn(ScaleLeft, ScaleTop, ScaleLeft + ScaleWidth, ScaleTop + ScaleHeight, CLng(sRndX), CLng(sRndY))
  End If
    
  
  
End Sub

Private Sub DestroyRegion(ByRef rgn As Long)
  If rgn <> 0 Then
    DeleteObject rgn
    rgn = 0
  End If
End Sub

Private Sub DestroyBrush(ByRef br As Long)
'  If br <> 0 Then
'    DeleteObject br
'    br = 0
'  End If
  DestroyRegion br
End Sub

Private Sub DestroyPen(ByRef br As Long)
'  If br <> 0 Then
'    DeleteObject br
'    br = 0
'  End If
  DestroyRegion br
End Sub


Private Sub Recreate_Brush(ByRef br As Long, ByVal clr As OLE_COLOR)
  DestroyBrush br
  
  Dim lbr As LOGBRUSH
  lbr.lbStyle = BS_SOLID
  lbr.lbColor = ConvertSysColor(clr)
  lbr.lbHatch = 0
  
  br = CreateBrushIndirect(lbr)
End Sub

Private Sub Recreate_Pen(ByRef br As Long, ByVal clr As OLE_COLOR)
  DestroyBrush br
    
  br = CreatePen(PS_SOLID, m_i_CantWidth, ConvertSysColor(clr))
End Sub

