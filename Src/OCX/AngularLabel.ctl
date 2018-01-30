VERSION 5.00
Begin VB.UserControl AngularLabel 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   85
   ToolboxBitmap   =   "AngularLabel.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "AngularLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
'Const m_def_Angle = 0
Const m_def_AutoSize = False
'Property Variables:
Private m_af_AFont As AngledFont
Private m_s_Caption As String
Private m_b_AutoSize As Boolean
'Event Declarations:
Event Click(ByVal sender As Object)  'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_UserMemId = -600

'Int. var.
Private m_bLockResize As Boolean
Private Const pi = 3.14159265358979

Private Sub UserControl_Click()
  RaiseEvent Click(Me)
End Sub

Public Property Get Bold() As Boolean
  Bold = m_af_AFont.Bold
End Property

Public Property Let Bold(ByVal bFl As Boolean)
  m_af_AFont.Bold = bFl
  PropertyChanged "FontCtx"
  UserControl.Refresh
End Property


Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
  Caption = m_s_Caption
End Property

Public Property Let Caption(ByVal sCap As String)
  m_s_Caption = sCap
  PropertyChanged "Caption"
    
  If Not m_b_AutoSize Then
    UserControl_Resize
  End If
  
  UserControl.Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = m_af_AFont.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set m_af_AFont.Font = New_Font
  PropertyChanged "FontCtx"
  UserControl.Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  UserControl.Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Angle() As Integer
  Angle = m_af_AFont.Angle
End Property

Public Property Let Angle(ByVal New_Angle As Integer)
  m_af_AFont.Angle = New_Angle
  PropertyChanged "FontCtx"
  UserControl_Resize
  If Not m_b_AutoSize Then UserControl.Refresh
End Property

Private Sub UserControl_Initialize()
  m_bLockResize = False
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set m_af_AFont = New AngledFont
  m_af_AFont.Font = UserControl.Font
  UserControl.ForeColor = &H0
  m_s_Caption = Extender.Name
  m_b_AutoSize = m_def_AutoSize
End Sub

Private Sub UserControl_Paint()
  Dim hFontOld As Long
  hFontOld = SelectObject(hdc, m_af_AFont.hFont)
  Dim r As RECT
  r.Left = ScaleLeft: r.Top = ScaleTop
  r.Right = r.Left + ScaleWidth: r.Bottom = r.Top + ScaleHeight
  
'  Dim iDrawModeOld As Integer
'  iDrawModeOld = UserControl.DrawMode
'  UserControl.DrawMode = vbCopyPen
  
  'DrawText hdc, m_s_Caption, -1, r, DT_CENTER Or DT_VCENTER Or DT_NOCLIP Or DT_SINGLELINE
  Dim lAln As Long
  lAln = SetTextAlign(hdc, TA_CENTER Or VTA_CENTER)
  
  Dim dAngl As Double
  dAngl = CDbl(m_af_AFont.Angle) * pi / 180#
  TextOut hdc, ScaleLeft + (CDbl(ScaleWidth) - CDbl(m_af_AFont.FHeight) * Sin(dAngl)) / 2#, ScaleTop + (CDbl(ScaleHeight) - CDbl(m_af_AFont.FHeight) * Cos(dAngl)) / 2#, m_s_Caption, Len(m_s_Caption)
  SetTextAlign hdc, lAln
    
  SelectObject hdc, hFontOld
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
  'Set m_af_AFont = New AngledFont
  Set m_af_AFont = PropBag.ReadProperty("FontCtx")
  m_s_Caption = PropBag.ReadProperty("Caption", Extender.Name)
  m_b_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
End Sub

Private Sub UserControl_Resize()
  If Not m_b_AutoSize Then Exit Sub
  
  If Not m_bLockResize Then
    m_bLockResize = True
    
    On Error GoTo ERR_H
    
    Dim lX As Long, lY As Long
    m_af_AFont.CalcExtent m_s_Caption, lX, lY
            
    lX = ScaleX(lX, vbPixels, ContainerScale)
    lY = ScaleY(lY, vbPixels, ContainerScale)
    
    Dim H As Long, L As Long
    Dim dAngl As Double, dX As Double, dY As Double
    dAngl = m_af_AFont.Angle
    If dAngl > 90 And dAngl < 180 Or dAngl > 270 And dAngl < 360 Then
      dAngl = dAngl - 90
      dX = lY
      dY = lX
    Else
      dX = lX
      dY = lY
    End If
    
    dAngl = dAngl * pi / 180#
    H = Abs(dX * Sin(dAngl) + dY * Cos(dAngl))
    L = Abs(dX * Cos(dAngl) + dY * Cos(pi / 2 - dAngl))
            
    Dim lCX As Long, lCY As Long
    lCX = Extender.Left + Extender.Width / 2
    lCY = Extender.Top + Extender.Height / 2
    
'    Extender.Top = lCY - H / 2: Extender.Height = H + 2
'    Extender.Left = lCX - L / 2: Extender.Width = L + 2
    
    Extender.Move lCX - L / 2, lCY - H / 2, L + 2, H + 2
        
    m_bLockResize = False
  End If
  Exit Sub
ERR_H:
  m_bLockResize = False
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub UserControl_Terminate()
  Set m_af_AFont = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H0)
   PropBag.WriteProperty "FontCtx", m_af_AFont
   PropBag.WriteProperty "Caption", m_s_Caption, Extender.Name
   PropBag.WriteProperty "AutoSize", m_b_AutoSize, m_def_AutoSize
End Sub


Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
  AutoSize = m_b_AutoSize
End Property

Public Property Let AutoSize(ByVal bNewValue As Boolean)
  m_b_AutoSize = bNewValue
  PropertyChanged "AutoSize"
  UserControl_Resize
  Refresh
End Property


Private Function ContainerScale() As Integer
' **** Gets the proper mode to scale to
    Dim eScaleTo As ScaleModeConstants
    
    On Error Resume Next
    
    ContainerScale = Extender.Container.ScaleMode
    
    If Err Then
        Err = 0
        ContainerScale = Parent.ScaleMode
        If Err Then
            Err = 0
            ContainerScale = vbTwips
        End If
    End If
    
    On Error GoTo 0
End Function

