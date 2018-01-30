VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.UserControl CtlRepeater2 
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ControlContainer=   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   6225
   Begin vsViewLib.vsViewPort vsViewPort1 
      Height          =   3930
      Left            =   450
      TabIndex        =   0
      Top             =   510
      Width           =   5220
      _Version        =   196608
      _ExtentX        =   9208
      _ExtentY        =   6932
      _StockProps     =   225
      Appearance      =   1
      Track           =   -1  'True
   End
End
Attribute VB_Name = "CtlRepeater2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements FacEditLib.IFECallback

'Default Property Values:
Const m_def_TypeMask = ""
Const m_def_FirstPaint = 2
'Property Variables:
Private m_gn_GertNet As GERTNETLib.MGertNet
Private m_s_TypeMask As String


'Event Declarations:
Event XClickAssEnum(ByVal feSender As Factor)
Event XClickSetupValue(ByVal feSender As Factor)
Event XFactorChanged(ByVal fac As Factor)

'Int. var.
Private m_bLockResize As Boolean
Private m_lSummHeightTwip As Long
Private m_i_FirstPaint As Integer


Public Property Let TypeMask(ByVal sMask As String)
  m_s_TypeMask = sMask
  ClearAll
  PropertyChanged "TypeMask"
End Property

Public Property Get TypeMask() As String
  TypeMask = m_s_TypeMask
End Property

Public Property Set GNet(ByVal f As MGertNet)
  If m_s_TypeMask = "" Then
    MsgBox "Сначала нужно назначить маску", vbOKOnly Or vbExclamation
    Exit Property
  End If
  Set m_gn_GertNet = f
  ClearAll
  If Not f Is Nothing Then _
    BuildContent
End Property

Public Property Get GNet() As MGertNet
  Set GNet = m_gn_GertNet
End Property


Private Sub UserControl_Initialize()
  m_bLockResize = False
  m_lSummHeightTwip = 0
  
  m_i_FirstPaint = 0
  
End Sub

Private Sub UserControl_InitProperties()
  m_s_TypeMask = m_def_TypeMask
End Sub

Private Sub UserControl_Paint()
  If m_i_FirstPaint <= m_def_FirstPaint Then
    
    If IsWindowVisible(hwnd) <> 0 And (Not Parent Is Nothing) Then
      If IsWindowVisible(Parent.hwnd) = 0 Then Exit Sub
                                  
      m_i_FirstPaint = m_def_FirstPaint + 1
      
      If Controls.Count > 1 Then
        Dim fed As FacEdit
        Set fed = Controls.Item(1)
        fed.SetFocus
      End If
            
    End If
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_s_TypeMask = PropBag.ReadProperty("TypeMask", m_def_TypeMask)
End Sub

Private Sub UserControl_Resize()
    
  m_bLockResize = True
  On Error GoTo ERR_H
  
  If vsViewPort1.Width <> ScaleWidth Or vsViewPort1.Height <> ScaleHeight Then
    vsViewPort1.Move 0, 0, ScaleWidth, ScaleHeight
  End If
  
  If Controls.Count = 1 Then
    m_lSummHeightTwip = 0
    vsViewPort1.VirtualHeight = ScaleHeight
    vsViewPort1.VirtualWidth = ScaleWidth
    m_bLockResize = False
    Exit Sub
  End If
    
  
  Dim vbeCtl As VBControlExtender
  Set vbeCtl = Controls(1)
    
  Dim sWidth As Single
  sWidth = vsViewPort1.Width - ScaleX(2 * GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips)
    
  If vbeCtl.Width <> sWidth Then
    Dim fe As FacEditLib.FacEdit
    Set fe = vbeCtl
    If sWidth >= fe.MinimumWidthTwip Or _
       (sWidth < vbeCtl.Width And vbeCtl.Width <> fe.MinimumWidthTwip) Then
      
      'm_l_MaxShiftX = 0
      sWidth = sWidth - ScaleX(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
      For Each vbeCtl In Controls
        vbeCtl.Width = sWidth
      Next vbeCtl
      Set vbeCtl = Controls(1)
      
    End If
  End If
    
  
  vsViewPort1.VirtualHeight = IIf(vsViewPort1.Height < m_lSummHeightTwip, m_lSummHeightTwip, vsViewPort1.Height)
  'vsViewPort1.VirtualWidth = IIf(ScaleWidth < vbeCtl.Width, vbeCtl.Width, ScaleWidth)
  vsViewPort1.VirtualWidth = vbeCtl.Width + 2 * ScaleX(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
    
  
  m_bLockResize = False
  Exit Sub
ERR_H:
  m_bLockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub UserControl_Terminate()
  Set m_gn_GertNet = Nothing
  ClearAll
End Sub

Private Sub ClearAll()
    
  Dim lCnt As Long
  For lCnt = UserControl.Controls.Count - 1 To 1 Step -1
    UserControl.Controls.Remove "XCtl" & CStr(lCnt)
  Next lCnt
  
  m_lSummHeightTwip = 0
  UserControl_Resize
  
End Sub

Private Sub BuildContent()
  
  ClearAll
  
  Dim lCnt As Long, lLen As Long, lTop As Long
  Dim fac As Factor, cextF As VBControlExtender
  lLen = Len(m_s_TypeMask)
  lCnt = 1: lTop = 0
  For Each fac In m_gn_GertNet.Factors
    Dim ibsk As IBSTRKey
    Set ibsk = fac
    If Left(ibsk.Get(), lLen) = m_s_TypeMask Then
            
      Dim facEd As FacEditLib.FacEdit
      Dim facExt As VBControlExtender
      Set facExt = Controls.Add("FacEditLib.FacEdit", "XCtl" & CStr(lCnt), vsViewPort1)
      
      If cextF Is Nothing Then _
        Set cextF = facExt
        
      Set facEd = facExt
        
      lCnt = lCnt + 1
      Set facEd.GNet = m_gn_GertNet
      Set facEd.HFactor = fac
      
      facExt.Move 0, lTop, ScaleWidth
      lTop = lTop + facEd.MinimumHeightTwip
      
      facExt.Visible = True
      
      m_lSummHeightTwip = m_lSummHeightTwip + facExt.Height
    End If
  Next fac
  
  UserControl_Resize
        
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "TypeMask", m_s_TypeMask, m_def_TypeMask
End Sub


Public Sub IFECallback_ClickAssEnum(ByVal feSender As Factor)
  RaiseEvent XClickAssEnum(feSender)
End Sub

Public Sub IFECallback_ClickSetupValue(ByVal feSender As Factor)
  RaiseEvent XClickSetupValue(feSender)
End Sub

Public Sub IFECallback_FactorChanged(ByVal fac As Factor)
  RaiseEvent XFactorChanged(fac)
End Sub



