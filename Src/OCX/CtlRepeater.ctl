VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.UserControl CtlRepeater 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ControlContainer=   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6180
   Begin Threed.SSPanel SSPanel1 
      Height          =   2205
      Left            =   615
      TabIndex        =   0
      Top             =   735
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   3889
      _Version        =   131074
      BevelOuter      =   0
   End
End
Attribute VB_Name = "CtlRepeater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements FacEditLib.IFECallback
Implements IMsgHandl

'Default Property Values:
Const m_def_TypeMask = ""
Const m_def_FirstPaint = 2
'Property Variables:
Private m_gn_GertNet As GERTNETLib.MGertNet
Private m_s_TypeMask As String


'Event Declarations:
Event XClickAssEnum(ByVal feSender As Factor)
Event XClickPr(ByVal feSender As Factor)
Event XClickAskVal(ByVal feSender As Factor)
Event XClickSetupValue(ByVal feSender As Factor)
Event XFactorChanged(ByVal fac As Factor)
Event XFocus(ByVal fac As Factor)

'Int. var.
Private m_bLockResize As Boolean
Private m_lSummHeightPix As Long
Private m_bHS As Boolean
Private m_bVS As Boolean
Private m_i_FirstPaint As Integer

Private m_l_OldWndProc As Long

Private m_l_ShiftXPix As Long, m_l_ShiftYPix As Long
Private m_l_MaxShiftX As Long, m_l_MaxShiftY As Long

Private m_b_LockH As Boolean
Private m_b_LockV As Boolean

Private m_f_CurrFacEd As IFacED

Private m_b_Hooked As Boolean

Public Property Get CurrFac() As Factor
  If CntControls() < 2 Then
    Set CurrFac = Nothing
  ElseIf m_f_CurrFacEd Is Nothing Then
    'Dim vbeCtl As VBControlExtender
    Set m_f_CurrFacEd = Controls(2).object
    Set CurrFac = m_f_CurrFacEd.HFactor
  Else
    Set CurrFac = m_f_CurrFacEd.HFactor
  End If
End Property

Public Property Let TypeMask(ByVal sMask As String)
  m_s_TypeMask = sMask
  ClearAll
  PropertyChanged "TypeMask"
End Property

Public Property Get TypeMask() As String
  TypeMask = m_s_TypeMask
End Property

Private Function CntMask(ByVal f As CollFac) As Long
  Dim ff As Factor, lLen As Long
  lLen = Len(m_s_TypeMask)
  CntMask = 0
  For Each ff In f
    With ff
      Dim ibsk As IBSTRKey
      Set ibsk = ff
      If Left(ibsk.Get(), lLen) = m_s_TypeMask Then _
        CntMask = CntMask + 1
    End With
  Next ff
End Function

Private Function CntControls() As Long
  CntControls = 0
  Dim ctl As Control
  For Each ctl In UserControl.Controls
    If TypeOf ctl Is FacEditLib.FacEdit Then _
      CntControls = CntControls + 1
  Next ctl
End Function


Public Property Set GNet(ByVal f As MGertNet)
  
  
  If m_s_TypeMask = "" Then
    MsgBox "Сначала нужно назначить маску", vbOKOnly Or vbExclamation
    Exit Property
  End If
  Set m_gn_GertNet = f
  If Not f Is Nothing Then
    
    If CntMask(f.Factors) <> CntControls() Then _
      ClearAll
      
    BuildContent
  Else
    Dim ctl As Control
    For Each ctl In UserControl.Controls
      If TypeOf ctl Is FacEditLib.FacEdit Then
        Dim fTmp As FacEditLib.FacEdit
        Set fTmp = ctl
        Set fTmp.GNet = Nothing
        Set fTmp.HFactor = Nothing
        'MsgBox "yyy"
      End If
    Next ctl
  End If
'  If Not f Is Nothing Then _
'    BuildContent
End Property

Public Property Get GNet() As MGertNet
  Set GNet = m_gn_GertNet
End Property



Private Sub UserControl_EnterFocus()
  If Not m_f_CurrFacEd Is Nothing Then
    'Dim ext As VBControlExtender
    'Set ext = m_f_CurrFac
    'm_f_CurrFacEd.SetFocus
    
  End If
End Sub

Private Sub UserControl_Initialize()
  m_b_Hooked = False
  
  m_b_LockH = False
  m_b_LockV = False
  m_bLockResize = False
  m_lSummHeightPix = 0
  m_bHS = False
  m_bVS = False
  m_l_OldWndProc = 0
  m_i_FirstPaint = 0
  m_l_ShiftXPix = 0: m_l_ShiftYPix = 0
  m_l_MaxShiftX = 0: m_l_MaxShiftY = 0
          
End Sub

Private Sub UserControl_InitProperties()
  m_s_TypeMask = m_def_TypeMask
  SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
  
  If UserControl.Ambient.UserMode And Not m_b_Hooked Then _
    Hook: m_b_Hooked = True
End Sub

Private Sub UserControl_Paint()

  If m_i_FirstPaint <= m_def_FirstPaint Then
'    If m_i_FirstPaint < m_def_FirstPaint Then
'      m_i_FirstPaint = m_i_FirstPaint + 1
'      Exit Sub
'    End If
    
    If IsWindowVisible(hwnd) <> 0 And (Not Parent Is Nothing) Then
      If IsWindowVisible(Parent.hwnd) = 0 Then Exit Sub
      'If Not Parent.Visible Then Exit Sub
      
                
'      If m_bHS Then _
'        ShowScrollBar hwnd, SB_HORZ, 1
'
'      If m_bVS Then _
'        ShowScrollBar hwnd, SB_VERT, 1
            
      m_i_FirstPaint = m_def_FirstPaint + 1
      
      If CntControls() > 0 Then
        Dim fed As FacEdit
        Set fed = Controls.Item(1)
        fed.SetFocus
      End If
  
      Dim h As Single
      h = Height
      Height = h + 10
      Height = h
      
    End If
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_s_TypeMask = PropBag.ReadProperty("TypeMask", m_def_TypeMask)
  
  SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
  
  If UserControl.Ambient.UserMode And Not m_b_Hooked Then _
    Hook: m_b_Hooked = True
End Sub

Private Sub UserControl_Resize()

  If SSPanel1.Height < ScaleHeight Then
    SSPanel1.Height = ScaleHeight
  End If

  If CntControls() = 0 Then
    If m_bHS Or m_bVS Then
      SetupVScroller False
      SetupHScroller False
      ShowScrollBar hwnd, SB_VERT, 0
      ShowScrollBar hwnd, SB_HORZ, 0
      m_bHS = False
      m_bVS = False
      m_l_MaxShiftX = 0: m_l_MaxShiftY = 0
      m_lSummHeightPix = 0
    End If
    Exit Sub
  End If
  
  m_bLockResize = True
  On Error GoTo ERR_H
  
  Dim vbeCtl As VBControlExtender
  Set vbeCtl = Controls(1)
    
  If vbeCtl.Width <> ScaleWidth Then
    Dim fe As FacEditLib.FacEdit
    Set fe = vbeCtl
    If ScaleWidth >= fe.MinimumWidthTwip Or _
       (ScaleWidth < vbeCtl.Width And vbeCtl.Width <> fe.MinimumWidthTwip) Then
      
      'm_l_MaxShiftX = 0
      
      For Each vbeCtl In Controls
        vbeCtl.Width = ScaleWidth
      Next vbeCtl
      For Each vbeCtl In Controls
        If TypeOf vbeCtl Is FacEdit Then
          Exit For
        End If
      Next vbeCtl
            
    End If
  End If
    
  If vbeCtl.Width > ScaleWidth Then
    m_l_MaxShiftX = ScaleX(vbeCtl.Width - ScaleWidth + 1, vbTwips, vbPixels)
    SetupHScroller True, m_l_MaxShiftX
    If Not m_bHS Then
      ShowScrollBar hwnd, SB_HORZ, 1
      m_bHS = True
    End If
  Else
    SetupHScroller False
    ShowScrollBar hwnd, SB_HORZ, 0
    m_bHS = False
    m_l_MaxShiftX = 0
  End If
  
  Dim sH As Single
  sH = ScaleY(ScaleHeight, vbTwips, vbPixels)
  m_l_MaxShiftY = m_lSummHeightPix - sH + 1
  If m_lSummHeightPix > sH Then
    SetupVScroller True, m_l_MaxShiftY
    If Not m_bVS Then
      ShowScrollBar hwnd, SB_VERT, 1
      m_bVS = True
    End If
  Else
    SetupVScroller False
    ShowScrollBar hwnd, SB_VERT, 0
    m_bVS = False
    m_l_MaxShiftY = 0
  End If
  
  Dim bFlShft As Boolean
  bFlShft = False
  If Abs(m_l_ShiftXPix) > m_l_MaxShiftX Then _
    m_l_ShiftXPix = -m_l_MaxShiftX: bFlShft = True
  If Abs(m_l_ShiftYPix) > m_l_MaxShiftY Then _
    m_l_ShiftYPix = -m_l_MaxShiftY: bFlShft = True

  If bFlShft Then ShiftContent
    'llllllllll
  
  m_bLockResize = False
  Exit Sub
ERR_H:
  m_bLockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub UserControl_Terminate()
'MsgBox "UserControl_Terminate", vbOKOnly, "TT"
  'If UserControl.Ambient.UserMode Then
    'Unhook
  Set m_f_CurrFacEd = Nothing
  Set m_gn_GertNet = Nothing
  ClearAll
  
End Sub

Private Sub ClearAll()
    
  Dim lCnt As Long
  For lCnt = UserControl.Controls.Count To 2 Step -1
    UserControl.Controls.Remove "XCtl" & CStr(lCnt)
  Next lCnt
  
  m_lSummHeightPix = 0
  UserControl_Resize
  
  Set m_f_CurrFacEd = Nothing
  
End Sub

Private Sub BuildContent()
    
  'ClearAll
  Dim bCre As Boolean
  bCre = IIf(CntControls() = 0, True, False)
  
  Dim lCnt As Long, lLen As Long, lTop As Long
  Dim fac As Factor, cextF As VBControlExtender
  lLen = Len(m_s_TypeMask)
  lCnt = 2: lTop = 0
  For Each fac In m_gn_GertNet.Factors
    Dim ibsk As IBSTRKey
    Set ibsk = fac
    If Left(ibsk.Get(), lLen) = m_s_TypeMask Then
            
      Dim facEd As FacEditLib.FacEdit
      Dim facExt As VBControlExtender
      Dim sNamC As String: sNamC = "XCtl" & CStr(lCnt)
      If bCre Then
        Set facExt = Controls.Add("FacEditLib.FacEdit", sNamC, SSPanel1)
      Else
        Set facExt = Controls(sNamC)
      End If
      
      If cextF Is Nothing Then _
        Set cextF = facExt
        
      Set facEd = facExt
        
      lCnt = lCnt + 1
      Set facEd.GNet = m_gn_GertNet
      Set facEd.HFactor = fac
      
      If bCre Then
        facExt.Move 0, lTop, ScaleWidth
        lTop = lTop + facEd.MinimumHeightTwip
        
        facExt.Visible = True
        
        m_lSummHeightPix = m_lSummHeightPix + ScaleY(facExt.Height, vbTwips, vbPixels)
      End If
    End If
  Next fac
  
  UserControl_Resize
      
'  If Controls.Count > 0 Then
'    Dim fed As FacEdit
'    Set fed = Controls.Item(0)
'    fed.SetFocus
'  End If
  
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "TypeMask", m_s_TypeMask, m_def_TypeMask
End Sub

Private Sub SetupVScroller(ByVal bFl As Boolean, Optional ByVal lLong As Long = -1)
  If IsWindow(hwnd) = 0 Then Exit Sub
  If bFl Then
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_VSCROLL
  Else
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And (Not WS_VSCROLL)
  End If
  If lLong <> -1 Then
    'lLong = ScaleY(lLong, vbTwips, vbPixels)
    SetScrollRange hwnd, SB_VERT, 0, lLong, 1
  End If
End Sub

Private Sub SetupHScroller(ByVal bFl As Boolean, Optional ByVal lLong As Long = -1)
  If IsWindow(hwnd) = 0 Then Exit Sub
  If bFl Then
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_HSCROLL
  Else
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And (Not WS_HSCROLL)
  End If
  If lLong <> -1 Then
    'lLong = ScaleX(lLong, vbTwips, vbPixels)
    SetScrollRange hwnd, SB_HORZ, 0, lLong, 1
  End If
End Sub


Private Sub Hook()
  BindCtl hwnd, Me
End Sub

Public Sub Unhook()
  If m_b_Hooked = False Then
    Err.Raise vbObjectError + 1, "Unhook", "Фильтр не установлен"
  Else
    m_b_Hooked = False
    UnbindCtl hwnd
  End If
End Sub


Public Function IMsgHandl_HandleHScroll(ByVal nScrollCode As Long, ByVal nPos As Long) As Long
  IMsgHandl_HandleHScroll = 0
    
  'Debug.Print nScrollCode, nPos, m_l_ShiftXPix
  
  On Error GoTo ERR_H
  If m_b_LockH Then Exit Function
  m_b_LockH = True
  

  Select Case nScrollCode
    Case SB_BOTTOM
      
    Case SB_ENDSCROLL
      'm_l_ShiftXPix = -m_l_MaxShiftX
    Case SB_LINELEFT
      m_l_ShiftXPix = m_l_ShiftXPix + 3
    Case SB_LINERIGHT
      m_l_ShiftXPix = m_l_ShiftXPix - 3
    Case SB_PAGELEFT
      m_l_ShiftXPix = m_l_ShiftXPix + ScaleX(ScaleWidth, vbTwips, vbPixels)
    Case SB_PAGERIGHT
      m_l_ShiftXPix = m_l_ShiftXPix - ScaleX(ScaleWidth, vbTwips, vbPixels)
    Case SB_THUMBPOSITION
      m_l_ShiftXPix = -nPos
    Case SB_THUMBTRACK
      m_l_ShiftXPix = -nPos
    Case SB_TOP
    Case Else
      'HandleHScroll = 1
      m_b_LockH = False
      Exit Function
  End Select
  
  If m_l_ShiftXPix > 0 Then
    m_l_ShiftXPix = 0
  ElseIf Abs(m_l_ShiftXPix) > m_l_MaxShiftX Then
    m_l_ShiftXPix = -m_l_MaxShiftX
  End If
  
  SetScrollPos hwnd, SB_HORZ, Abs(m_l_ShiftXPix), 1
  
  ShiftContent
  
  m_b_LockH = False
  Exit Function
  
ERR_H:
  m_b_LockH = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


Public Function IMsgHandl_HandleVScroll(ByVal nScrollCode As Long, ByVal nPos As Long) As Long
  IMsgHandl_HandleVScroll = 0
  
  'Debug.Print nScrollCode, nPos, m_l_ShiftXPix
  
  On Error GoTo ERR_H
  If m_b_LockV Then Exit Function
  m_b_LockV = True
  

  Select Case nScrollCode
    Case SB_BOTTOM
      
    Case SB_ENDSCROLL
      'm_l_ShiftXPix = -m_l_MaxShiftX
    Case SB_LINELEFT
      m_l_ShiftYPix = m_l_ShiftYPix + 7
    Case SB_LINERIGHT
      m_l_ShiftYPix = m_l_ShiftYPix - 7
    Case SB_PAGELEFT
      m_l_ShiftYPix = m_l_ShiftYPix + ScaleY(ScaleHeight, vbTwips, vbPixels)
    Case SB_PAGERIGHT
      m_l_ShiftYPix = m_l_ShiftYPix - ScaleY(ScaleHeight, vbTwips, vbPixels)
    Case SB_THUMBPOSITION
      m_l_ShiftYPix = -nPos
    Case SB_THUMBTRACK
      m_l_ShiftYPix = -nPos
    Case SB_TOP
    Case Else
      'HandleHScroll = 1
      m_b_LockV = False
      Exit Function
  End Select
  
  If m_l_ShiftYPix > 0 Then
    m_l_ShiftYPix = 0
  ElseIf Abs(m_l_ShiftYPix) > m_l_MaxShiftY Then
    m_l_ShiftYPix = -m_l_MaxShiftY
  End If
  
  SetScrollPos hwnd, SB_VERT, Abs(m_l_ShiftYPix), 1
  
  ShiftContent
  
  m_b_LockV = False
  Exit Function
  
ERR_H:
  m_b_LockV = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub IMsgHandl_HandleSHW(ByVal lFlag As Long, ByVal lStatus As Long)
  If lFlag = 0 Then m_i_FirstPaint = 0
    
    
End Sub

Private Sub ShiftContent()
  'LockWindowUpdate hwnd
  
  Dim cextF As VBControlExtender
  Dim lTop As Long, lDelta As Long, lX As Long, lTop2 As Long
  lDelta = Controls(1).MinimumHeightTwip
  lTop = ScaleY(m_l_ShiftYPix, vbPixels, vbTwips)
  'lTop2 = lTop
  lX = ScaleX(m_l_ShiftXPix, vbPixels, vbTwips)
'  For Each cextF In Controls
'    'cextF.Move lX, lTop
'    lTop2 = lTop2 + lDelta
'  Next cextF
  
  SSPanel1.Move lX, lTop, ScaleWidth, lDelta * (Controls.Count - 1)
  
  'DoEvents
  'ExcludeClipRect hdc, 0, 0, ScaleX(ScaleWidth, vbTwips, vbPixels), ScaleY(ScaleHeight, vbTwips, vbPixels)
  'LockWindowUpdate 0
  UpdateWindow hwnd
End Sub

Public Sub IFECallback_ClickPr(ByVal feSender As Factor)
  RaiseEvent XClickPr(feSender)
End Sub


Public Sub IFECallback_ClickAssEnum(ByVal feSender As Factor)
  RaiseEvent XClickAssEnum(feSender)
End Sub

Public Sub IFECallback_ClickAskVal(ByVal feSender As Factor)
  RaiseEvent XClickAskVal(feSender)
End Sub

Public Sub IFECallback_ClickSetupValue(ByVal feSender As Factor)
  RaiseEvent XClickSetupValue(feSender)
End Sub

Public Sub IFECallback_FactorChanged(ByVal fac As Factor)
  RaiseEvent XFactorChanged(fac)
End Sub

Public Sub IFECallback_FactorFocus(ByVal facEd As IFacED)
  If Not m_f_CurrFacEd Is Nothing Then
    If Not m_f_CurrFacEd Is facEd Then _
      m_f_CurrFacEd.ClearFocus
  End If
  Set m_f_CurrFacEd = facEd
  RaiseEvent XFocus(facEd.HFactor)
End Sub

Public Sub UpdateValues(Optional ByVal fFac As Factor = Nothing)
  Dim cextF As VBControlExtender
  Dim fe As FacEditLib.FacEdit
  For Each cextF In Controls
    If TypeOf cextF Is FacEdit Then
      Set fe = cextF
      If fFac Is Nothing Then
        fe.RefreshValues
      ElseIf fe.HFactor Is fFac Then
        fe.RefreshValues
        Exit Sub
      End If
    End If
  Next cextF
End Sub
