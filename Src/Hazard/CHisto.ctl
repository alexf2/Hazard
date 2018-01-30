VERSION 5.00
Begin VB.UserControl CHisto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Private Const TA_CENTER = 6
Private Const VTA_CENTER = TA_CENTER
Private Const TA_RIGHT = 2
Private Const TA_TOP = 0
Private Const TA_BOTTOM = 8
Private Const TA_LEFT = 0


Private Const OPAQUE = 2
Private Const TRANSPARENT = 1


Private Type Size
        cx As Long
        cy As Long
End Type



Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_BOTTOM = &H8




'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private m_def_LPad As Single '= 900
Private Const m_def_UPad = 100#
Private Const m_def_RPad = 100#
Private m_def_BPad As Single  '= 700
Private m_s_BH As Single

Private m_af As AngledFont
Private m_f_Lbl As Font

Private Const sYMask = "#0.0#####;;0"
Private Const sYMask0 = "#0.000000;;0"

Private Const sXMask = "#0.############;;0"
Private Const sXMask0 = "#0.000000000000;;0"

Private Type TPoint
  x1 As Double
  x2 As Double
  y As Double
  lShift As Long
  ocClr As OLE_COLOR
End Type

Private m_arrPts() As TPoint
Private m_l_NPts As Long

Private m_d_MaxX As Double
Private m_d_MaxY As Double

Private m_d_MulX As Double
Private m_d_MulY As Double

Private m_d_WorkW As Double
Private m_d_WorkH As Double

Private m_d_Trunc As Double

Private m_s_X0 As Single 'left top
Private m_s_Y0 As Single

Private m_s_X1 As Single 'left bottom
Private m_s_Y1 As Single

Private m_s_X2 As Single 'right bottom
Private m_s_Y2 As Single

Private Sub UserControl_Initialize()
  
  m_l_NPts = 0
  m_d_MaxX = 0
  m_d_MaxY = 0
  m_d_MulX = 1
  m_d_MulY = 1
  
  m_s_X0 = 0: m_s_Y0 = 0
  m_s_X1 = 0: m_s_Y1 = 0
  m_s_X2 = 0: m_s_Y2 = 0
    
  Set m_af = New AngledFont
  m_af.Angle = 90
  
  m_def_LPad = 900
  m_def_BPad = 700

End Sub

Public Property Get Trunc() As Double
  Trunc = m_d_Trunc
End Property

Public Property Let Trunc(ByVal New_Tr As Double)
  m_d_Trunc = New_Tr
  PropertyChanged "Trunc"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Get FontLbl() As Font
  Set FontLbl = m_f_Lbl
End Property
Public Property Set FontLbl(ByVal fnt As Font)
  Set m_f_Lbl = fnt
  Set m_af.Font = fnt
  m_af.Angle = 90
  PropertyChanged "FontLbl"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  UserControl.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
  UserControl.Appearance() = New_Appearance
  PropertyChanged "Appearance"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
  m_d_Trunc = 1
  
  Set m_f_Lbl = Ambient.Font
  Set m_af.Font = Ambient.Font
  m_af.Angle = 90
  
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
  
  m_d_Trunc = PropBag.ReadProperty("Trunc", 1)
  
  Set m_f_Lbl = PropBag.ReadProperty("FontLbl", Ambient.Font)
  Set m_af.Font = m_f_Lbl
  m_af.Angle = 90
End Sub



'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
  Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
  
  PropBag.WriteProperty "Trunc", m_d_Trunc, 1
  PropBag.WriteProperty "FontLbl", m_f_Lbl, Ambient.Font
  
End Sub


Public Sub SetArrData(arr() As Double)
  On Error Resume Next
  Dim l1 As Long, l2 As Long, lSz As Long
  l1 = LBound(arr)
  If Err.Number <> 0 Then lSz = 0 Else lSz = (UBound(arr) - LBound(arr) + 1) / 4 * Trunc
  On Error GoTo 0
  
  m_d_MaxX = 0: m_d_MaxY = 0
  
  If lSz = 0 Then
    ReDim m_arrPts(0) As TPoint
    m_l_NPts = 0
    RepaintControl
    Exit Sub
  Else
    ReDim m_arrPts(lSz - 1) As TPoint
    m_l_NPts = lSz
  End If
    
  Dim clr As OLE_COLOR, clrNew As OLE_COLOR
  
  Dim sPX1 As Single, sPY1 As Single
  sPX1 = ScaleX(2, vbPixels, vbTwips)
  sPY1 = ScaleY(2, vbPixels, vbTwips)
    
  m_def_LPad = 0
  m_def_BPad = 0
  m_s_BH = 0
  
  Dim m_def_LPad0 As Single, m_def_BPad0 As Single
  
  'm_af.Angle = 0
  
  Dim lIdx As Long: lIdx = 0
  Dim lShft As Long: lShft = 0
  For l1 = LBound(arr) To LBound(arr) + 4 * lSz - 1 Step 4
    With m_arrPts(lIdx)
      .x1 = arr(l1 + 1)
      .x2 = arr(l1 + 2)
      .y = arr(l1)
      .lShift = lShft
                      
      If .x2 + lShft + 1 > m_d_MaxX Then m_d_MaxX = .x2 + lShft + 1
      If .y > m_d_MaxY Then m_d_MaxY = .y
      
      Do
       clrNew = RndClr()
      Loop While clrNew = clr
      clr = clrNew
      .ocClr = clr
              
      m_def_LPad0 = UserControl.TextWidth(Format(.y, sYMask0)) + 7 * sPX1
      'Dim sz As Size
      'GetTextExtentPoint32 hdc, Format(.y, sYMask), Len(Format(.y, sYMask)), sz
      'sz.cx = sz.cx
      'm_def_LPad0 = ScaleX(sz.cx + 12, vbPixels, vbTwips)
      
      Dim sz As Size
      
      If m_def_LPad < m_def_LPad0 Then m_def_LPad = m_def_LPad0
      
      Dim hFontOld As Long: hFontOld = SelectObject(hdc, m_af.hFont)
      Dim ssTmp As String
      ssTmp = CutPoint(Format(.x2, sXMask))
      GetTextExtentPoint32 hdc, ssTmp, Len(ssTmp), sz
      m_def_BPad0 = ScaleY(2 * sz.cx + 16, vbPixels, vbTwips)
      If m_def_BPad < m_def_BPad0 Then
        m_def_BPad = m_def_BPad0
        m_s_BH = ScaleY(sz.cx, vbPixels, vbTwips)
      End If
      SelectObject hdc, hFontOld
            
      lShft = lShft + 1
    End With
    lIdx = lIdx + 1
  Next l1
      
  'm_af.Angle = 90
    
  UserControl_Resize
  RepaintControl
  
End Sub

Private Sub RepaintControl()
    
  UserControl.Refresh
    
End Sub

Private Sub UserControl_Resize()
  m_d_WorkW = ScaleWidth - (m_def_LPad + m_def_RPad) - 100
  m_d_WorkH = ScaleHeight - (m_def_UPad + m_def_BPad) - 100
  
  If m_d_MaxX = 0 Then
    m_d_MulX = 1
  Else
    m_d_MulX = m_d_WorkW / m_d_MaxX
  End If
  
  If m_d_MaxY = 0 Then
    m_d_MulY = 1
  Else
    m_d_MulY = m_d_WorkH / m_d_MaxY
  End If
  
'  UserControl.ScaleMode = vbUser
'  UserControl.ScaleHeight = -UserControl.Height
'  UserControl.ScaleTop = UserControl.Height - m_def_BPad
'  UserControl.ScaleLeft = m_def_LPad
  
  m_s_X0 = m_def_LPad
  m_s_Y0 = m_def_UPad
  
  m_s_X1 = m_def_LPad
  m_s_Y1 = ScaleHeight - m_def_BPad
  
  m_s_X2 = ScaleWidth - m_def_RPad
  m_s_Y2 = ScaleHeight - m_def_BPad
  

End Sub

Private Sub UserControl_Paint()
  'Line (0, 0)-(200, 200), RGB(255, 100, 50)
  Dim x As Single, l As Long
        
  UserControl.FillStyle = vbFSSolid
  UserControl.DrawStyle = vbSolid
  UserControl.DrawWidth = 2
  
  Dim sPX1 As Single, sPY1 As Single
  sPX1 = ScaleX(2, vbPixels, vbTwips)
  sPY1 = ScaleY(2, vbPixels, vbTwips)
    
  Line (m_s_X0 - sPX1, m_s_Y0)-(m_s_X0 - sPX1, m_s_Y1), RGB(0, 0, 0)
  Line (m_s_X1, m_s_Y1 + sPY1)-(m_s_X2, m_s_Y1 + sPY1), RGB(0, 0, 0)
  
  UserControl.DrawWidth = 1
  UserControl.DrawStyle = vbDash
  
  Dim y As Double
  For y = m_s_Y1 - 400 To m_s_Y0 Step -400
    Dim sTmp As String
    sTmp = Format((m_s_Y1 - y) / m_d_MulY, sYMask)
    
    UserControl.CurrentX = m_s_X0 - UserControl.TextWidth(sTmp) - 4 * sPX1
    UserControl.CurrentY = y - UserControl.TextHeight(sTmp) / 2
    
    Print sTmp
    
    Line (m_s_X1, y)-(m_s_X2, y), RGB(0, 0, 0)
  Next y
    
  UserControl.DrawStyle = vbSolid
    
  On Error GoTo ERR_H
  
  Dim hFontOld As Long: hFontOld = SelectObject(hdc, m_af.hFont)
  Dim xLbl As Double, sDelta As Double, sz As Size
  GetTextExtentPoint32 hdc, "5", 3, sz
  sDelta = ScaleX(sz.cy, vbPixels, vbTwips) / 2
  xLbl = -sDelta
  
  Dim lclrOld As Long, lMOld As Long, lAlnOld As Long
  lclrOld = SetTextColor(hdc, RGB(0, 0, 0))
  lMOld = SetBkMode(hdc, TRANSPARENT)
  
  'setBkMode hdc, OPAQUE
  'SetBkColor hdc, RGB(255, 0, 0)
  'SetTextAlign hdc, TA_RIGHT Or VTA_CENTER
  'SetTextAlign hdc, TA_TOP Or TA_RIGHT
  lAlnOld = SetTextAlign(hdc, TA_TOP Or TA_RIGHT)
  
  Dim bFlOdd As Boolean:  bFlOdd = False
  For l = 0 To m_l_NPts - 1
        
    With m_arrPts(l)
      UserControl.FillColor = .ocClr
      'x = .x1 * m_d_MulX
      
      Dim dYh As Double
      dYh = .y * m_d_MulY
      If .y <> 0 And ScaleY(dYh, vbTwips, vbPixels) < 1 Then dYh = sPY1
      
      Dim dXTmp As Double: dXTmp = m_s_X1 + (.x1 + .lShift) * m_d_MulX
      Line (dXTmp, m_s_Y1)-(m_s_X1 + (.x2 + .lShift + 1) * m_d_MulX, m_s_Y1 - .y * m_d_MulY), .ocClr, BF
      
      If dXTmp - xLbl >= sDelta Then
        
        Dim sTxt As String, rc As RECT:
        sTxt = CutPoint(Format(.x1, sXMask))
        rc.Left = ScaleX(dXTmp, vbTwips, vbPixels) - sz.cy / 2
        
        Dim sS1 As Single, sS2 As Single
        If bFlOdd Then
          sS1 = sPY1 * 4
          sS2 = 3 * sPY1
        Else
          sS1 = m_s_BH + 3 * sPY1 + sPY1 * 3
          sS2 = m_s_BH + 2 * sPY1 + sPY1 * 3
        End If
        
        bFlOdd = Not bFlOdd
        
        rc.Top = ScaleY(m_s_Y1 + sS1, vbTwips, vbPixels)
        
        TextOut hdc, rc.Left, rc.Top, sTxt, Len(sTxt)
        
        Line (dXTmp, m_s_Y1)-(dXTmp, m_s_Y1 + sS2), RGB(0, 0, 0)
        
'        rc.Left = ScaleX(m_s_X1, vbTwips, vbPixels)
'        rc.Top = ScaleY(m_s_Y1, vbTwips, vbPixels)
'        rc.Bottom = rc.Top + 50
'        rc.Right = rc.Left + 200
        
        'Line (ScaleX(rc.Left, vbPixels, vbTwips), ScaleY(rc.Top, vbPixels, vbTwips))-(ScaleX(rc.Right, vbPixels, vbTwips), ScaleY(rc.Bottom, vbPixels, vbTwips)), RGB(0, 255, 0), BF
        
        'DrawText hdc, "PPP Tst", 7, rc, DT_LEFT Or DT_TOP Or DT_NOCLIP Or DT_SINGLELINE
        'TextOut hdc, rc.Left, rc.Top, "ppp Tst", 7
        
        xLbl = dXTmp
      End If
    End With
  Next l
  
  SetTextColor hdc, lclrOld
  SetBkMode hdc, lMOld
  SetTextAlign hdc, lAlnOld
  SelectObject hdc, hFontOld
  Exit Sub
  
ERR_H:
  SelectObject hdc, hFontOld
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Function RndClr() As OLE_COLOR
  Static arrConst As Variant
  If IsEmpty(arrConst) Then _
    arrConst = Array(RGB(255, 127, 255), RGB(127, 255, 255), RGB(255, 255, 255), _
      RGB(255, 127, 127), RGB(0, 0, 0), _
      RGB(63, 63, 128), RGB(127, 127, 255), RGB(128, 63, 128), _
      RGB(63, 128, 63), RGB(127, 255, 127), RGB(255, 255, 127), _
      RGB(63, 128, 128), RGB(128, 128, 0), RGB(128, 0, 0))
      
        
  RndClr = arrConst(CInt(UBound(arrConst) - LBound(arrConst)) * Rnd + LBound(arrConst))
        
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
  hWnd = UserControl.hWnd
End Property

Private Function CutPoint(ByVal sStr As String) As String
  Dim sStr1 As String, lSz As Long
  lSz = Len(sStr)
  sStr1 = Right(sStr, 1)
  
  If lSz < 1 Then
    CutPoint = ""
  ElseIf sStr1 = "." Or sStr1 = "," Then
    CutPoint = Left(sStr, lSz - 1)
  Else
    CutPoint = sStr
  End If
End Function
