VERSION 5.00
Object = "{CC52DF59-28C5-11D4-8F1B-00504E02C39D}#53.0#0"; "GNControls.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Begin VB.Form frmMEditor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Редактор модели"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VsOcxLib.VideoSoftIndexTab tabEditor 
      Height          =   3375
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   5940
      _Version        =   327680
      _ExtentX        =   10477
      _ExtentY        =   5953
      _StockProps     =   102
      Caption         =   "Человек|Машина|Рабочая среда|Технология"
      ConvInfo        =   1418783674
      BackColor       =   12632256
      ForeColor       =   0
      FrontTabColor   =   12632256
      BackTabColor    =   12632256
      Style           =   3
      AutoScroll      =   -1  'True
      ShowFocusRect   =   0   'False
      BackSheets      =   1
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      FrontTabForeColor=   16711680
      New3D           =   0   'False
      MouseIcon       =   "frmMEditor.frx":0000
      Picture(0)      =   "frmMEditor.frx":001C
      Picture(1)      =   "frmMEditor.frx":0186
      Picture(2)      =   "frmMEditor.frx":02F0
      Picture(3)      =   "frmMEditor.frx":045A
      Begin GNControls.CtlRepeater CtlRepeater4 
         Height          =   2925
         Left            =   75645
         TabIndex        =   9
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "T"
      End
      Begin GNControls.CtlRepeater CtlRepeater3 
         Height          =   2925
         Left            =   75345
         TabIndex        =   8
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "C"
      End
      Begin GNControls.CtlRepeater CtlRepeater2 
         Height          =   2925
         Left            =   75045
         TabIndex        =   7
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "M"
      End
      Begin GNControls.CtlRepeater CtlRepeater1 
         Height          =   2925
         Left            =   45
         TabIndex        =   6
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "H"
      End
   End
   Begin Threed.SSFrame ssfAlho 
      Height          =   1575
      Left            =   540
      TabIndex        =   1
      ToolTipText     =   "Алгоритм, используемый для прогонов модели"
      Top             =   3720
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   2778
      _Version        =   131074
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Алгоритм"
      ShadowStyle     =   1
      Begin Threed.SSOption ssoptAnalytic2 
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   1185
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Аналитический устойчивый"
      End
      Begin Threed.SSOption ssoptAnalytic 
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   870
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Аналитический"
      End
      Begin Threed.SSOption ssoptImitate2 
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   555
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Имитационный устойчивый"
      End
      Begin Threed.SSOption ssoptImitate 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   255
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Имитационный"
         Value           =   -1
      End
   End
   Begin Threed.SSCheck sscheckPreview 
      Height          =   540
      Left            =   3915
      TabIndex        =   5
      Top             =   4065
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   953
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      BackColor       =   12632256
      Caption         =   "Показать функцию индекса опасности"
   End
End
Attribute VB_Name = "frmMEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_LPad = 100
Private Const m_UPad = 70
Private Const m_L2Pad = 200

Private m_mg_GertNet As MGertNet
Private m_b_LockGUI As Boolean

Public Property Get MinW() As Single
  MinW = 7000
End Property

Public Property Get MinH() As Single
  MinH = 4500
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
    
  If bFlShow Then
    'EnblControls Me, Not haApp.IsCalcAny
  End If
End Function


Private Sub CtlRepeater1_XClickAssEnum(ByVal feSender As GERTNETLib.IFactor)
  MakeAss feSender, m_mg_GertNet
End Sub

Private Sub CtlRepeater2_XClickAssEnum(ByVal feSender As GERTNETLib.IFactor)
  MakeAss feSender, m_mg_GertNet
End Sub

Private Sub CtlRepeater3_XClickAssEnum(ByVal feSender As GERTNETLib.IFactor)
  MakeAss feSender, m_mg_GertNet
End Sub

Private Sub CtlRepeater4_XClickAssEnum(ByVal feSender As GERTNETLib.IFactor)
  MakeAss feSender, m_mg_GertNet
End Sub


Private Sub MakeAss(ByVal f As Factor, ByVal m As MGertNet)
  On Error GoTo ERR_H
  
  frmEnumAss.AssData f, m
  frmEnumAss.Show vbModal, frmMain
  frmEnumAss.UnassData
  
  If frmEnumAss.ModalResult Then
    Dim bm As New CBeam
    bm.Beam Me
    DoEvents
    UpdateValuesAll
  End If
  
  Exit Sub
ERR_H:
  frmEnumAss.UnassData
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Friend Sub UpdateValuesAll()
  CtlRepeater1.UpdateValues
  CtlRepeater2.UpdateValues
  CtlRepeater3.UpdateValues
  CtlRepeater4.UpdateValues
End Sub

Private Sub CtlRepeater1_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater1_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater2_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater2_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater3_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater3_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater4_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater4_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub Form_Initialize()
  m_b_LockGUI = False
End Sub

Private Sub Form_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxModelMon.GetExtraWH sWExtr, sHExtr
  
  tabEditor.Move m_LPad, m_UPad, ScaleWidth - 2 * m_LPad - sWExtr, ScaleHeight - 3 * m_UPad - ssfAlho.Height - sHExtr
  ssfAlho.Move tabEditor.Left, tabEditor.Top + tabEditor.Height + m_UPad
  sscheckPreview.Move ssfAlho.Left + ssfAlho.Width + 2 * m_L2Pad, ssfAlho.Top + (ssfAlho.Height - sscheckPreview.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  HandsOffModel
  CtlRepeater1.Unhook
  CtlRepeater2.Unhook
  CtlRepeater3.Unhook
  CtlRepeater4.Unhook
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set CtlRepeater1.GNet = Nothing
  Set CtlRepeater2.GNet = Nothing
  Set CtlRepeater3.GNet = Nothing
  Set CtlRepeater4.GNet = Nothing
  
  Set m_mg_GertNet = Nothing
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
  
    
  On Error GoTo ERR_H
  m_b_LockGUI = True
  
  If mgn.RunMode = MT_Analytical Then
    ssoptAnalytic.Value = True
  ElseIf mgn.RunMode = MT_ImitateIndistinct Then
    ssoptImitate2.Value = True
  ElseIf mgn.RunMode = MT_Imitate Then
    ssoptImitate.Value = True
  Else
    ssoptAnalytic2.Value = True
  End If
  
    
  Set CtlRepeater1.GNet = mgn
  Set CtlRepeater2.GNet = mgn
  Set CtlRepeater3.GNet = mgn
  Set CtlRepeater4.GNet = mgn
  
  If sscheckPreview.Value = True Then
    m_mg_GertNet.Run 5, 10, False, -1, True
    ShowSPFunc CurrFac
  End If
      
  m_b_LockGUI = False
  Exit Sub
ERR_H:
  m_b_LockGUI = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Private Sub sscheckPreview_Click(Value As Integer)
  If Value = True Then
    frmGraphBar.Visible = True
    If Not m_mg_GertNet Is Nothing Then _
      m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
    
  Else
    frmGraphBar.Visible = False
  End If
End Sub

Private Sub ssoptAnalytic_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptAnalytic.Value = True Then
      frmMon.ssoptAvto.Enabled = False
      frmMon.ssoptVal.Enabled = False
      frmMon.pvnRnd.Enabled = False
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_Analytical
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
    End If
  End If
End Sub

Private Sub ssoptAnalytic2_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptAnalytic2.Value = True Then
      frmMon.ssoptAvto.Enabled = False
      frmMon.ssoptVal.Enabled = False
      frmMon.pvnRnd.Enabled = False
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_AnalyticalIndistinct
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
    End If
  End If
End Sub

Private Sub ssoptImitate_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptImitate.Value = True Then
      frmMon.ssoptAvto.Enabled = True
      frmMon.ssoptVal.Enabled = True
      frmMon.pvnRnd.Enabled = True
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_Imitate
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
    End If
  End If
End Sub

Private Sub ssoptImitate2_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptImitate2.Value = True Then
      frmMon.ssoptAvto.Enabled = True
      frmMon.ssoptVal.Enabled = True
      frmMon.pvnRnd.Enabled = True
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_ImitateIndistinct
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
  
    End If
  End If
End Sub

Friend Sub ShowSPFunc(ByVal fac As Factor)

  If sscheckPreview.Value = True And IsBoundModel And Not fac Is Nothing Then
    Dim mgDbg As IMGertNet_Debug
    Set mgDbg = m_mg_GertNet
    Dim arrdPts() As Double, sFactName As String
    Dim iKey As IBSTRKey
    Set iKey = fac
    sFactName = iKey.Get
    arrdPts = mgDbg.GetPtsForFactor(sFactName)
    Dim l As Long, lSz As Long, lPts As Long
    lSz = UBound(arrdPts) - LBound(arrdPts) + 1
    lPts = lSz / 2
    
    With frmGraphBar.Graph1
      .DrawMode = GraphLib.DrawModeConstants.gphNoAction
      
      .AutoInc = AutoIncConstants.gphOff
      .RandomData = RandomDataConstants.gphOff
      .DataReset = gphGraphData
      .NumSets = 1
      .ThisSet = 1
      .NumPoints = lPts
      '.TickEvery = 2
      
      '.Background = gphBlue
      
      .ThisSet = 1
      
      l = 0
      .ThisPoint = 1
      Dim lPtCnt As Long
      lPtCnt = 1
      Do While l < lSz
            
        .ThisPoint = lPtCnt
        lPtCnt = lPtCnt + 1
        .XPosData = arrdPts(l)
        l = l + 1
        .GraphData = arrdPts(l)
        l = l + 1
          
      Loop
              
      .PrintStyle = PrintStyleConstants.gphColor
      .GridStyle = GridStyleConstants.gphBoth
      
      Dim sDescr As String, sTrust As String, dValue As Double
      m_mg_GertNet.GetFactorPresent fac, sDescr, sTrust, dValue
      
      If m_mg_GertNet.RunMode <> MT_ImitateIndistinct And _
        m_mg_GertNet.RunMode <> MT_AnalyticalIndistinct Then sTrust = "*"
      
      Dim ss As String
      .BottomTitle = sFactName & "(" & CStr(lPts) & " pts) <" & sTrust & "> = " & FormatNumber(dValue) & "(" & sDescr & ")"
      .GraphTitle = fac.Name
      '.Labels = LabelsConstants.gphOn
      
      .FontUse = gphOtherTitles
      .FontStyle = GraphLib.FontStyleConstants.gphItalic
      .DrawMode = GraphLib.DrawModeConstants.gphBlit
    End With
    
  End If
End Sub

Friend Function CurrFac() As Factor
  If m_mg_GertNet Is Nothing Then
    Set CurrFac = Nothing
  End If
  
  Dim rep As CtlRepeater
  Set rep = CurrRepeater()
  If rep Is Nothing Then
    Set CurrFac = Nothing
  Else
    Set CurrFac = rep.CurrFac
  End If
  
End Function

Friend Function CurrRepeater() As CtlRepeater
  Set CurrRepeater = XRepeater(tabEditor.CurrTab)
End Function

Friend Function XRepeater(ByVal l As Long) As CtlRepeater
  Select Case l
    Case 0
      Set XRepeater = CtlRepeater1
    Case 1
      Set XRepeater = CtlRepeater2
    Case 2
      Set XRepeater = CtlRepeater3
    Case 3
      Set XRepeater = CtlRepeater4
  End Select
End Function


Private Sub tabEditor_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
  If Not XRepeater(NewTab) Is Nothing Then _
    ShowSPFunc XRepeater(NewTab).CurrFac
End Sub
