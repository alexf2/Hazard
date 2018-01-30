VERSION 5.00
Object = "{CC52DF59-28C5-11D4-8F1B-00504E02C39D}#76.0#0"; "GNControls.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Begin VB.UserControl CfrmMEditor 
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   ScaleHeight     =   5835
   ScaleWidth      =   6330
   Begin VsOcxLib.VideoSoftIndexTab tabEditor 
      Height          =   3375
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5940
      _Version        =   327680
      _ExtentX        =   10477
      _ExtentY        =   5953
      _StockProps     =   102
      Caption         =   "Человек|Машина|Рабочая среда|Технология"
      ConvInfo        =   1418783674
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
      MouseIcon       =   "CfrmMEditor.ctx":0000
      Picture(0)      =   "CfrmMEditor.ctx":001C
      Picture(1)      =   "CfrmMEditor.ctx":0186
      Picture(2)      =   "CfrmMEditor.ctx":02F0
      Picture(3)      =   "CfrmMEditor.ctx":045A
      Begin GNControls.CtlRepeater CtlRepeater4 
         Height          =   2925
         Left            =   75645
         TabIndex        =   12
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "T"
      End
      Begin GNControls.CtlRepeater CtlRepeater3 
         Height          =   2925
         Left            =   75345
         TabIndex        =   11
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "C"
      End
      Begin GNControls.CtlRepeater CtlRepeater2 
         Height          =   2925
         Left            =   75045
         TabIndex        =   10
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "M"
      End
      Begin GNControls.CtlRepeater CtlRepeater1 
         Height          =   2925
         Left            =   45
         TabIndex        =   9
         Top             =   405
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   5159
         TypeMask        =   "H"
      End
   End
   Begin Threed.SSFrame ssfAlho 
      Height          =   1350
      Left            =   510
      TabIndex        =   8
      ToolTipText     =   "Алгоритм, используемый для прогонов модели"
      Top             =   3600
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   2381
      _Version        =   131074
      Font3D          =   1
      ForeColor       =   0
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
      Begin Threed.SSCommand ssOptions 
         Height          =   345
         Left            =   105
         TabIndex        =   6
         ToolTipText     =   "Смотреть числовые характеристики модели"
         Top             =   900
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         PictureUseMask  =   -1  'True
         Picture         =   "CfrmMEditor.ctx":05C4
         Caption         =   "Параметры"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSOption ssoptAnalytic22 
         Height          =   285
         Left            =   2805
         TabIndex        =   5
         Top             =   945
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         Caption         =   "Ан. уст. приближенный"
      End
      Begin Threed.SSOption ssoptImitate 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   255
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         Caption         =   "Имитационный"
         Value           =   -1
      End
      Begin Threed.SSOption ssoptImitate2 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         Caption         =   "Имитационный устойчивый"
      End
      Begin Threed.SSOption ssoptAnalytic 
         Height          =   285
         Left            =   2805
         TabIndex        =   3
         Top             =   255
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         Caption         =   "Аналитический"
      End
      Begin Threed.SSOption ssoptAnalytic2 
         Height          =   285
         Left            =   2805
         TabIndex        =   4
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   131074
         ForeColor       =   0
         Caption         =   "Аналитический устойчивый"
      End
   End
   Begin Threed.SSCheck sscheckPreview 
      Height          =   465
      Left            =   525
      TabIndex        =   7
      Top             =   5100
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   820
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      Caption         =   "Функция принадлежности индексу опасности"
   End
End
Attribute VB_Name = "CfrmMEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const m_LPad = 100
Private Const m_UPad = 70
Private Const m_L2Pad = 200

Private m_mg_GertNet As MGertNet
Private m_b_LockGUI As Boolean

Public Caption As String

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get MousePointer() As Long
  MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal lP As Long)
  UserControl.MousePointer = lP
End Property


Public Property Get PssoptImitate() As SSOption
  Set PssoptImitate = ssoptImitate
End Property

Public Property Get PssoptImitate2() As SSOption
  Set PssoptImitate2 = ssoptImitate2
End Property

Public Property Get PssoptAnalytic() As SSOption
  Set PssoptAnalytic = ssoptAnalytic
End Property

Public Property Get PssoptAnalytic2() As SSOption
  Set PssoptAnalytic2 = ssoptAnalytic2
End Property


Private Sub CtlRepeater1_XClickAskVal(ByVal feSender As GERTNETLib.IFactor)
  MkAskVal feSender
End Sub

Private Sub CtlRepeater1_XClickPr(ByVal feSender As GERTNETLib.IFactor)
  MkEditPr feSender
End Sub

Private Sub CtlRepeater2_XClickAskVal(ByVal feSender As GERTNETLib.IFactor)
  MkAskVal feSender
End Sub

Private Sub CtlRepeater2_XClickPr(ByVal feSender As GERTNETLib.IFactor)
  MkEditPr feSender
End Sub

Private Sub CtlRepeater3_XClickAskVal(ByVal feSender As GERTNETLib.IFactor)
  MkAskVal feSender
End Sub

Private Sub CtlRepeater3_XClickPr(ByVal feSender As GERTNETLib.IFactor)
  MkEditPr feSender
End Sub

Private Sub CtlRepeater4_XClickAskVal(ByVal feSender As GERTNETLib.IFactor)
  MkAskVal feSender
End Sub

Private Sub CtlRepeater1_XClickSetupValue(ByVal feSender As GERTNETLib.IFactor)
  MkAssIdx feSender
End Sub

Private Sub CtlRepeater2_XClickSetupValue(ByVal feSender As GERTNETLib.IFactor)
  MkAssIdx feSender
End Sub

Private Sub CtlRepeater3_XClickSetupValue(ByVal feSender As GERTNETLib.IFactor)
  MkAssIdx feSender
End Sub

Private Sub CtlRepeater4_XClickPr(ByVal feSender As GERTNETLib.IFactor)
  MkEditPr feSender
End Sub

Private Sub CtlRepeater4_XClickSetupValue(ByVal feSender As GERTNETLib.IFactor)
  MkAssIdx feSender
End Sub

Private Sub MkAskVal(ByVal feSender As GERTNETLib.IFactor)
  'MsgBox "Назначить " & feSender.Name, vbOKOnly, "tt"
  
  If haApp.GertNetMain Is Nothing Then
    MsgBox "Нет модели", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If haApp.GertNetMain.ModuleProgID = "" Then
    MsgBox "Не назначен модуль оценки значений факторов опасности", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  Dim sPath As String
  sPath = frmAssMgr.GetPathForCfg(haApp.GertNetMain.ModuleProgID, haApp.GertNetMain.ModuleConfig)
  
  'On Error Resume Next
  
  Dim ifass As IFactorAssign
  If frmPlugin.m_tIfass Is Nothing Then
    Set frmPlugin.m_tIfass = CreateObject(haApp.GertNetMain.ModuleProgID)
    Set ifass = frmPlugin.m_tIfass
    frmPlugin.Move ScaleX(50, vbPixels, vbTwips), ScaleY(50, vbPixels, vbTwips)
    frmPlugin.Show vbModeless, frmMain
    frmPlugin.ssCancel.Enabled = False
    DoEvents
  End If
  
  If ifass Is Nothing Then Set ifass = frmPlugin.m_tIfass
  
'  If Err.Number <> 0 Then
'    MsgBox "Не удаётся загрузить модуль оценки ФО '" & haApp.GertNetMain.ModuleProgID & "':" & vbCrLf & Err.Description, vbOKOnly Or vbCritical, "Ошибка"
'    Exit Sub
'  End If
    
  On Error GoTo ERR_H
    
  ifass.AssignFactor frmMain, frmMain.hWnd, haApp.GertNetMain, feSender, haApp.GertNetMain.ModuleConfig, sPath
  DoEvents
  
  If ifass.ModalResult <> 0 Then
    Dim bm As New CBeam
    bm.Beam Me
    DoEvents
    
    UpdateValuesAll feSender
    
    If sscheckPreview.Value = True Then _
      m_mg_GertNet.Run 5, 10, False, -1, True
    
    ShowSPFunc feSender
  End If
  
  'Set ifass = Nothing
  frmPlugin.ssCancel.Enabled = True
  Exit Sub
ERR_H:
  'Set ifass = Nothing
  'Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
  Unload frmPlugin
End Sub

Private Sub MkEditPr(ByVal feSender As GERTNETLib.IFactor)
        
  On Error GoTo ERR_H
  
  With frmPr
    .AssData haApp.GertNetMain, feSender
    .Show vbModal, frmMain
    DoEvents
      
    .UnassData
  End With
  
  Exit Sub
ERR_H:
  frmPr.UnassData
  'Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub MkAssIdx(ByVal feSender As GERTNETLib.IFactor)
        
  On Error GoTo ERR_H
  
  frmIdx.AssData feSender
  Dim ibsk As IBSTRKey: Set ibsk = feSender
  frmIdx.Caption = "Назначить индексы опасности для '" & ibsk.Get() & "'"
  frmIdx.Show vbModal, frmMain
  DoEvents
  
  If frmIdx.ModalResult Then
    Dim bm As New CBeam
    bm.Beam Me
    DoEvents
    
    UpdateValuesAll feSender
    
    If sscheckPreview.Value = True Then _
      m_mg_GertNet.Run 5, 10, False, -1, True
    
    ShowSPFunc feSender
  End If
  
  frmIdx.UnassData
  
  Exit Sub
ERR_H:
  frmIdx.UnassData
  'Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssOptions_Click()
  On Error GoTo ERR_H
  
  frmAlhoOpt.AssData haApp.GertNetMain
  frmAlhoOpt.Show vbModal, frmMain
  DoEvents
  
  frmAlhoOpt.UnassData
  
  Exit Sub
ERR_H:
  frmAlhoOpt.UnassData
  'Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub UserControl_Initialize()
  m_b_LockGUI = False
  Caption = "Редактор модели"
End Sub


Public Property Get PsscheckPreview() As SSCheck
  Set PsscheckPreview = sscheckPreview
End Property

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
  'Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Friend Sub UpdateValuesAll(Optional ByVal fFac As GERTNETLib.IFactor = Nothing)
  CtlRepeater1.UpdateValues fFac
  CtlRepeater2.UpdateValues fFac
  CtlRepeater3.UpdateValues fFac
  CtlRepeater4.UpdateValues fFac
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
  If sscheckPreview.Value = True And IsBoundModel Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater2_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater3_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True And IsBoundModel Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater3_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater4_XFactorChanged(ByVal fac As GERTNETLib.IFactor)
  If sscheckPreview.Value = True And IsBoundModel Then _
    m_mg_GertNet.Run 5, 10, False, -1, True
    
  ShowSPFunc fac
End Sub

Private Sub CtlRepeater4_XFocus(ByVal fac As GERTNETLib.IFactor)
  ShowSPFunc fac
End Sub


Private Sub UserControl_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxModelMon.GetExtraWH sWExtr, sHExtr
  
  tabEditor.Move m_LPad, m_UPad, ScaleWidth - 2 * m_LPad - sWExtr, ScaleHeight - 3 * m_UPad - ssfAlho.Height - sHExtr
  ssfAlho.Move tabEditor.Left, tabEditor.Top + tabEditor.Height + m_UPad
  sscheckPreview.Move ssfAlho.Left + ssfAlho.Width + m_L2Pad, ssfAlho.Top + (ssfAlho.Height - sscheckPreview.Height) / 2
End Sub


Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set CtlRepeater1.GNet = Nothing
  Set CtlRepeater2.GNet = Nothing
  Set CtlRepeater3.GNet = Nothing
  Set CtlRepeater4.GNet = Nothing
  
  ssOptions.Enabled = False
  
  Set m_mg_GertNet = Nothing
End Sub

Public Sub SetupNumbers()

End Sub


Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
  SetupNumbers
    
  On Error GoTo ERR_H
  m_b_LockGUI = True
  
  If mgn.RunMode = MT_Analytical Then
    ssoptAnalytic.Value = True
  ElseIf mgn.RunMode = MT_ImitateIndistinct Then
    ssoptImitate2.Value = True
  ElseIf mgn.RunMode = MT_Imitate Then
    ssoptImitate.Value = True
  ElseIf mgn.RunMode = MT_Analytical2 Then
    ssoptAnalytic22.Value = True
  Else
    ssoptAnalytic2.Value = True
  End If
  
  ssOptions.Enabled = True
    
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
      frmMon.PssoptAvto.Enabled = False
      frmMon.PssoptVal.Enabled = False
      frmMon.PpvnRnd.Enabled = False
      
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
      frmMon.PssoptAvto.Enabled = False
      frmMon.PssoptVal.Enabled = False
      frmMon.PpvnRnd.Enabled = False
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_AnalyticalIndistinct
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
    End If
  End If
End Sub

Private Sub ssoptAnalytic22_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptAnalytic22.Value = True Then
      frmMon.PssoptAvto.Enabled = False
      frmMon.PssoptVal.Enabled = False
      frmMon.PpvnRnd.Enabled = False
      
      If Not m_mg_GertNet Is Nothing Then
        m_mg_GertNet.RunMode = MT_Analytical2
        
      If sscheckPreview.Value = True Then _
        m_mg_GertNet.Run 5, 10, False, -1, True: ShowSPFunc CurrFac
      End If
    End If
  End If
End Sub


Private Sub ssoptImitate_Click(Value As Integer)
  If Not m_b_LockGUI Then
    If ssoptImitate.Value = True Then
      frmMon.PssoptAvto.Enabled = True
      frmMon.PssoptVal.Enabled = True
      frmMon.PpvnRnd.Enabled = True
      
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
      frmMon.PssoptAvto.Enabled = True
      frmMon.PssoptVal.Enabled = True
      frmMon.PpvnRnd.Enabled = True
      
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
    Dim arrdPts() As Double, arrdPts2() As Double, sFactName As String
    Dim iKey As IBSTRKey
    Set iKey = fac
    sFactName = iKey.Get
    arrdPts = mgDbg.GetPtsForFactor(sFactName)
    Dim l As Long, lSz As Long, lSz2 As Long, lPts As Long, lPts2 As Long
    lSz = UBound(arrdPts) - LBound(arrdPts) + 1
    lPts = lSz / 2
    
    arrdPts2 = mgDbg.GetDensityForFactor(sFactName)
    lSz2 = UBound(arrdPts2) - LBound(arrdPts2) + 1
    lPts2 = lSz2 / 2
    
    With frmGraphBar.Graph1
      .DrawMode = GraphLib.DrawModeConstants.gphNoAction
      
      .ThisSet = 1
      .DrawStyle = 1
      .PrintStyle = PrintStyleConstants.gphColor
      .GridStyle = GridStyleConstants.gphBoth
      
      .AutoInc = AutoIncConstants.gphOff
      .RandomData = RandomDataConstants.gphOff
      .DataReset = gphLabelText Or gphGraphData
      .Labels = LabelsConstants.gphOn
      .NumSets = 2
      .ThisSet = 2
      '.ColorData = ColorDataConstants.gphLightCyan
      .NumPoints = IIf(lPts > lPts2, lPts, lPts2)
      '.TickEvery = 2
      '.Ticks = gphTicksOn
      '.LabelEvery = 1
      
      '.Background = gphBlue
      
      '.ThisSet = 2
      
      l = 0
      .ThisPoint = 1
      Dim lPtCnt As Long, vLastX, lBnd As Long
      lBnd = UBound(arrdPts)
      For lPtCnt = 1 To .NumPoints
        .ThisPoint = lPtCnt
        .XPosData = arrdPts(lBnd - 1)
        .GraphData = arrdPts(lBnd)
      Next lPtCnt
      
      lPtCnt = 1
      Do While l < lSz
            
        .ThisPoint = lPtCnt
        lPtCnt = lPtCnt + 1
        .XPosData = arrdPts(l)
        
'        If IsEmpty(vLastX) Or vLastX <> arrdPts(l) Then _
'          .LabelText = FormatNumber(arrdPts(l), 3, vbTrue, vbFalse, vbFalse)
'        vLastX = arrdPts(l)
         
        
        l = l + 1
        .GraphData = arrdPts(l)
        l = l + 1
          
      Loop
      
            
      '.ColorData = ColorDataConstants.gphLightGreen
      .ThisSet = 1
      .DrawStyle = 1
      .PrintStyle = PrintStyleConstants.gphColor
      '.ColorData = ColorDataConstants.gphLightGreen
      '.NumPoints = lPts
            
      lBnd = UBound(arrdPts2)
      For lPtCnt = 1 To .NumPoints
        .ThisPoint = lPtCnt
        .XPosData = arrdPts2(lBnd - 1)
        .GraphData = arrdPts2(lBnd)
      Next lPtCnt
      
      l = 0
      .ThisPoint = 1
      lPtCnt = 1
      Do While l < lSz2
            
        .ThisPoint = lPtCnt
        lPtCnt = lPtCnt + 1
        .XPosData = arrdPts2(l)
                
        l = l + 1
        .GraphData = arrdPts2(l)
        l = l + 1
          
      Loop
              
      
      
      Dim sDescr As String, sTrust As String, dValue As Double
      m_mg_GertNet.GetFactorPresent fac, sDescr, sTrust, dValue
      
      If m_mg_GertNet.RunMode <> MT_ImitateIndistinct And _
        m_mg_GertNet.RunMode <> MT_AnalyticalIndistinct And _
        m_mg_GertNet.RunMode <> MT_Analytical2 _
          Then sTrust = "*"
      
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


Private Sub UserControl_Terminate()
  HandsOffModel
  CtlRepeater1.Unhook
  CtlRepeater2.Unhook
  CtlRepeater3.Unhook
  CtlRepeater4.Unhook
End Sub

Private Sub ssOptions_LostFocus()
  HighLight ssOptions, False
End Sub

Private Sub ssOptions_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssOptions, True
End Sub

Private Sub ssOptions_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssOptions, False
End Sub

