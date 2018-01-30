VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HazardApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_gn_GertNetMain As GERTNETLib.MGertNet

Private WithEvents m_gn_GertNetRun As GERTNETLib.MGertNet
Attribute m_gn_GertNetRun.VB_VarHelpID = -1
Private WithEvents m_gn_GertNetRang As GERTNETLib.MGertNet
Attribute m_gn_GertNetRang.VB_VarHelpID = -1
Private WithEvents m_gn_GertNetOpt As GERTNETLib.MGertNet
Attribute m_gn_GertNetOpt.VB_VarHelpID = -1
Private m_ic_Run As ICallBack
Private m_ic_Rang As ICallBack
Private m_ic_Opt As ICallBack

Private m_l_CurrSFColl As Long
Private m_s_NameDoc As String
Private m_coll_SFSF As CollectionX
Private m_b_DirtySFSF As Boolean
Private m_b_NewDoc As Boolean

Private m_cx_Reports1 As CollectionX, m_cx_Reports2 As CollectionX

Private Const sFmt_Normal = "#0.0##############;;0"
Private Const sFmt2_Normal = "#0.0###;;0"
Private Const sFmt_Scientific = "0.0##############e-###;;0"
Private Const sFmt2_Scientific = "0.0###e-###;;0"


Public CreRepRun As Boolean
Public CreRepOpt As Boolean
Public CreRepVApply As Boolean

Public Property Get FmtDbl() As String
  If HaveDoc Then
    FmtDbl = m_gn_GertNetMain.NFormat(NFT_Pr)
  Else
    FmtDbl = sFmt_Normal
  End If
End Property
Public Property Get FmtDbl2() As String
  If HaveDoc Then
    FmtDbl2 = m_gn_GertNetMain.NFormat(NFT_Other)
  Else
    FmtDbl2 = sFmt2_Normal
  End If
End Property


Public Property Get IsNewDoc() As Boolean
  IsNewDoc = m_b_NewDoc
End Property

Friend Sub SetNewDoc()
  m_b_NewDoc = True
End Sub

Public Property Get Rep1() As CollectionX
  Set Rep1 = m_cx_Reports1
End Property

Public Property Get Rep2() As CollectionX
  Set Rep2 = m_cx_Reports2
End Property

Public Sub CloseCalc()
  Set m_gn_GertNetRun = Nothing
End Sub
Public Sub CloseRang()
  Set m_gn_GertNetRang = Nothing
End Sub
Public Sub CloseOpt()
  Set m_gn_GertNetOpt = Nothing
End Sub

Public Sub StartOpt(ByVal dP0 As Double, ByVal mgBase As MGertNet, ByVal collTmp As collSF, _
  ByVal otTask As OptTask, ByVal vDta As Variant, _
  ByVal otType As OptType, ByVal iRetN As Integer, ByVal tpPrior As GERTNETLib.ThrdPriority, _
  ByVal ic As ICallBack)
  
  On Error GoTo ERR_H
      
  Dim icln As IClonable
  Set icln = mgBase
  Set m_gn_GertNetOpt = icln.Clone
  
  With m_gn_GertNetOpt
    Set m_ic_Opt = ic
    .CalcSpeed = tpPrior
    .OptimizationType = otType
    .UserProp = otType
    
    If otTask = OT_FixMoney_MinQ Then
      .OptimizeSelectSF otTask, collTmp, CCur(vDta), 0, dP0, iRetN
    Else
      .OptimizeSelectSF otTask, collTmp, mgBase.MoneyForSF, CDbl(vDta), dP0, iRetN
    End If
    
  End With
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetOpt = Nothing
  Set m_ic_Opt = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub VirtualApply(ByVal mgBase As MGertNet, ByVal collTmp As collSF, ByVal tpPrior As GERTNETLib.ThrdPriority, ByVal ic As ICallBack)
  On Error GoTo ERR_H
  
  If m_coll_SFSF Is Nothing Then
    Err.Raise vbObjectError + 1, "HazardApp.StartCalc", "Нет загруженной модели"
  End If
    
  Dim icln As IClonable
  Set icln = mgBase
  Set m_gn_GertNetRang = icln.Clone
  
  With m_gn_GertNetRang
    Set m_ic_Rang = ic
    .CalcSpeed = tpPrior
    .UserProp = US_SP_Calc2
    .ApplySF collTmp
    .Run .k, .N, False, .RndBase
  End With
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetRang = Nothing
  Set m_ic_Rang = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub StartRang(ByVal mgBase As MGertNet, ByVal coll As collSF, ByVal tpPrior As GERTNETLib.ThrdPriority, ByVal ic As ICallBack)
  On Error GoTo ERR_H
  
  If m_coll_SFSF Is Nothing Then
    Err.Raise vbObjectError + 1, "HazardApp.StartCalc", "Нет загруженной модели"
  End If
    
  Dim icln As IClonable
  Set icln = mgBase
  Set m_gn_GertNetRang = icln.Clone
  'Set m_gn_GertNetRang = mgBase
  
  With m_gn_GertNetRang
    Set m_ic_Rang = ic
    .CalcSpeed = tpPrior
    .UserProp = US_SP_Calc1
    .CalcDeltaQ coll, m_gn_GertNetRang.k, m_gn_GertNetRang.N, -1
  End With
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetRang = Nothing
  Set m_ic_Rang = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub StartCalc(ByVal lk As Long, ByVal lN As Long, ByVal lRnd As Long, ByVal tpPrior As GERTNETLib.ThrdPriority, ByVal ic As ICallBack)
  On Error GoTo ERR_H
  
  If m_gn_GertNetMain Is Nothing Then
    Err.Raise vbObjectError + 1, "HazardApp.StartCalc", "Нет загруженной модели"
  End If
  
  Dim icln As IClonable
  Set icln = m_gn_GertNetMain
  Set m_gn_GertNetRun = icln.Clone
  m_gn_GertNetRun.StatOn = TSF_Yes
  
  With m_gn_GertNetRun
    Set m_ic_Run = ic
    .CalcSpeed = tpPrior
    .Run lk, lN, False, lRnd
  End With
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetRun = Nothing
  Set m_ic_Run = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub StartCalibrate(ByVal iPhaze As Integer, ByVal dP1 As Double, ByVal dP2 As Double, ByVal dP3 As Double, ByVal dP4 As Double, ByVal tpPrior As GERTNETLib.ThrdPriority, ByVal ic As ICallBack)
  On Error GoTo ERR_H
  
  If m_gn_GertNetMain Is Nothing Then
    Err.Raise vbObjectError + 1, "HazardApp.StartCalibrate", "Нет загруженной модели"
  End If
  
  If iPhaze = 1 Then
    Dim icln As IClonable
    Set icln = m_gn_GertNetMain
    Set m_gn_GertNetRun = icln.Clone
  End If
  
  With m_gn_GertNetRun
    Set m_ic_Run = ic
    If iPhaze = 1 Then
      .CalcSpeed = tpPrior
            
      Dim fac As Factor
      For Each fac In .Factors
        'fac.TrustLvl = TL_Normal
        'fac.Value = 5
      Next fac
      .RunMode = MT_Analytical
    ElseIf iPhaze = 2 Then
      .RunMode = MT_AnalyticalIndistinct
    End If
    .CalibrateModel dP1, dP2, dP3, dP4
  End With
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetRun = Nothing
  Set m_ic_Run = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Property Get GN_Run() As MGertNet
  Set GN_Run = m_gn_GertNetRun
End Property
Public Property Get GN_Rang() As MGertNet
  Set GN_Rang = m_gn_GertNetRang
End Property
Public Property Get GN_Opt() As MGertNet
  Set GN_Opt = m_gn_GertNetOpt
End Property

Public Sub CancelCalc()
  If m_gn_GertNetRun Is Nothing Then Exit Sub
  On Error GoTo ERR_H
  If m_gn_GertNetRun.IsRunning Then
    m_gn_GertNetRun.Cancel
  End If
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Ошибка"
  Err.Clear
End Sub

Public Sub CancelRang()
  If m_gn_GertNetRang Is Nothing Then Exit Sub
  On Error GoTo ERR_H
  If m_gn_GertNetRang.IsRunning2 Then
    m_gn_GertNetRang.Cancel
  End If
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Ошибка"
  Err.Clear
End Sub
Public Sub CancelOpt()
  If m_gn_GertNetOpt Is Nothing Then Exit Sub
  On Error GoTo ERR_H
  If m_gn_GertNetOpt.IsRunningOpt Then
    m_gn_GertNetOpt.Cancel
  End If
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Ошибка"
  Err.Clear
End Sub

Friend Property Let DirtySFSF(ByVal bFl As Boolean)
  m_b_DirtySFSF = bFl
End Property

Friend Property Get DirtySFSF() As Boolean
  If m_b_DirtySFSF Or m_coll_SFSF Is Nothing Then
    DirtySFSF = m_b_DirtySFSF
  Else
    Dim cll As collSF
    For Each cll In m_coll_SFSF
      If cll.IsDirty Then
        DirtySFSF = True
        Exit Property
      End If
    Next cll
    DirtySFSF = False
  End If
End Property

Public Property Get IsCalcAny() As Boolean
  IsCalcAny = IsCalcRun Or IsCalcRang Or IsCalcOpt
End Property

Public Property Get IsCalcRun() As Boolean
  If m_gn_GertNetRun Is Nothing Then
    IsCalcRun = False
  Else
    With m_gn_GertNetRun
      IsCalcRun = .IsRunning Or .IsRunning2 Or .IsRunningOpt
    End With
  End If
End Property

Public Property Get IsCalcRang() As Boolean
  If m_gn_GertNetRang Is Nothing Then
    IsCalcRang = False
  Else
    With m_gn_GertNetRang
      IsCalcRang = .IsRunning Or .IsRunning2 Or .IsRunningOpt
    End With
  End If
End Property

Public Property Get IsCalcOpt() As Boolean
  If m_gn_GertNetOpt Is Nothing Then
    IsCalcOpt = False
  Else
    With m_gn_GertNetOpt
      IsCalcOpt = .IsRunning Or .IsRunning2 Or .IsRunningOpt
    End With
  End If
End Property

'Friend Property Set GertNetMain(ByVal gn As MGertNet)
'  Set m_gn_GertNetMain = gn
'End Property

Public Property Get GertNetMain() As MGertNet
  Set GertNetMain = m_gn_GertNetMain
End Property

Public Property Get GertNetMainDsp() As Object
  Set GertNetMainDsp = m_gn_GertNetMain
End Property


Public Property Get XCollection() As CollectionX
  Set XCollection = m_coll_SFSF
End Property

Public Property Let CurrSFColl(ByVal collN As Long)
  If m_coll_SFSF Is Nothing Then
    m_l_CurrSFColl = -1
    If collN <> -1 Then _
      Err.Raise vbObjectError + 1, "HazardApp", "Нельзя назначить текущий набор мер: они не загружены - " & collN
    Exit Property
  End If
  
  If collN <> -1 And m_coll_SFSF Is Nothing Then _
    Err.Raise vbObjectError + 1, "HazardApp", "Нельзя назначить текущий набор мер: они не загружены - " & collN
  If collN < -1 Or collN > m_coll_SFSF.Count Then _
    Err.Raise vbObjectError + 1, "HazardApp", "Номер набора мер выходит за пределы: " & collN
  
  m_l_CurrSFColl = collN
End Property

Public Property Get CurrSFColl() As Long
  CurrSFColl = m_l_CurrSFColl
End Property

Public Property Get SFnn(ByVal lIdx As Long) As collSF
  Set SFnn = m_coll_SFSF(lIdx)
End Property

Public Property Get SFCount() As Long
  SFCount = m_coll_SFSF.Count
End Property

Public Property Get NameDoc() As String
  NameDoc = m_s_NameDoc
End Property

Public Property Get HaveDoc() As Boolean
  HaveDoc = Not (m_gn_GertNetMain Is Nothing)
End Property

Public Sub OpenDoc(ByVal sName As String, Optional ByVal bAsk As Boolean = True)
  On Error GoTo ERR_H
  
  If Not (m_gn_GertNetMain Is Nothing) Or Not (m_coll_SFSF Is Nothing) Then
    If CloseDoc(bAsk) = vbCancel Then Exit Sub
  End If
      
  Dim iStg As IStorage
  Set iStg = OpenCF(sName, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Dim gnTmp As New MGertNet, ips As IPersistStorage
  Set ips = gnTmp
  
  ips.Load iStg
  
  Set m_gn_GertNetMain = gnTmp
  m_s_NameDoc = sName
  m_b_NewDoc = False
  m_b_DirtySFSF = False
  
  UpdateGUI_MonModel Me
  
  Set ips = Nothing: Set iStg = Nothing
  OpenSF bAsk
  
  Set m_cx_Reports1 = New CollectionX
  Set m_cx_Reports2 = New CollectionX
  
  'm_gn_GertNetMain.Factors("C02").Name = "Качество информационной модели состояния среды"
  
  UpdateGUI_MonSF Me
  UpdateGUI_Reports Me
  
  m_gn_GertNetMain.Snapshot True
  
  Exit Sub
ERR_H:
  Set m_gn_GertNetMain = Nothing
  m_s_NameDoc = "<пусто>"
  UpdateGUI_MonModel Me
  UpdateGUI_MonSF Me
  
  Set m_cx_Reports1 = Nothing
  Set m_cx_Reports2 = Nothing
  UpdateGUI_Reports Me
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Public Sub OpenSF(Optional ByVal bAsk As Boolean = True)
  On Error GoTo ERR_H
    
  Dim iStg As IStorage
  Set iStg = OpenCF(m_s_NameDoc, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Dim iStg2 As IStorage
  Set iStg2 = OpenStorageInt(iStg, "SFCollections", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  Dim arrStgs() As String, arrStrm() As String
  DirOfStorage iStg2, TDR_Stg, arrStgs, arrStrm
  
  On Error Resume Next
  Dim lU As Long: lU = UBound(arrStgs)
  Dim bFlErr As Boolean: bFlErr = Err.Number <> 0 Or lU < 0
  On Error GoTo ERR_H
  
  Set m_coll_SFSF = New CollectionX
  m_l_CurrSFColl = -1
  
  If bFlErr Then Exit Sub
    
  Dim l As Long
  For l = LBound(arrStgs) To UBound(arrStgs)
    Dim collTmp As collSF
    Set collTmp = New collSF
    Dim ipsTmp As IPersistStorage
    Set ipsTmp = collTmp
    Dim iStg3 As IStorage
    Set iStg3 = OpenStorageInt(iStg2, arrStgs(l), STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    ipsTmp.Load iStg3
    
    m_coll_SFSF.Add collTmp, arrStgs(l)
  Next l
  
  m_l_CurrSFColl = 1
  m_b_DirtySFSF = False
    
  UpdateGUI_MonSF Me
  
  Exit Sub
ERR_H:
  Set m_coll_SFSF = Nothing
  m_l_CurrSFColl = -1
  UpdateGUI_MonSF Me
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

Public Sub SelectSave(ByVal bAsk As Boolean)
  If Not m_b_NewDoc Then
    If m_gn_GertNetMain Is Nothing Then
      frmSave.SaveModel = False
    Else
      frmSave.SaveModel = m_gn_GertNetMain.IsDirty
    End If
    frmSave.ssModel.Enabled = frmSave.SaveModel
    
    If m_coll_SFSF Is Nothing Then
      frmSave.SaveSP = False
    Else
      frmSave.SaveSP = TestDirtySFSF
    End If
    frmSave.ssSP.Enabled = frmSave.SaveSP
    
    If Not frmSave.SaveModel And Not frmSave.SaveSP Then
      frmSave.ModalResult = vbNo
    ElseIf bAsk Then
      frmSave.ModalResult = vbCancel
      frmSave.Show vbModal, frmMain
      DoEvents
    End If
  Else
    frmSave.SaveModel = True
    frmSave.SaveSP = True
    frmSave.ssModel.Enabled = False
    frmSave.ssSP.Enabled = False
    
    If Not frmSave.SaveModel And Not frmSave.SaveSP Then
      frmSave.ModalResult = vbNo
    ElseIf bAsk Then
      frmSave.ModalResult = vbCancel
      frmSave.Show vbModal, frmMain
      DoEvents
    End If
  End If
  
  If Not bAsk Then frmSave.ModalResult = vbNo
End Sub

Public Function CloseDoc(ByVal bAsk As Boolean) As Long
  On Error GoTo ERR_H
  
'  Dim lTst As Long: lTst = vbNo
'  Dim bTst2 As Boolean: bTst2 = False
'  If Not (m_coll_SFSF Is Nothing) Then
'    If TestDirtySFSF Then
'      If bAsk Then
'        lTst = MsgBox("Набор мер модифицирован. Сохранить ?", vbYesNoCancel Or vbQuestion, "Запрос")
'        bTst2 = True
'      Else
'        lTst = vbYes
'      End If
'    End If
'  End If
'
'  If lTst = vbCancel Then _
'    CloseDoc = vbCancel: Exit Function
'
'  If Not (m_gn_GertNetMain Is Nothing) Then
'    If lTst = vbNo And Not bTst2 Then
'      If m_gn_GertNetMain.IsDirty Then
'        If bAsk Then
'          lTst = MsgBox("Модель модифицирована. Сохранить ?", vbYesNoCancel Or vbQuestion, "Запрос")
'        Else
'          lTst = vbYes
'        End If
'      End If
'    End If
'  End If
  
  If Not HaveDoc Then
    CloseDoc = vbNo
    Exit Function
  End If
  
  frmSave.ssNo.Enabled = True
  SelectSave bAsk
    
  If frmSave.ModalResult = vbCancel Then
    CloseDoc = vbCancel
    Exit Function
  ElseIf frmSave.ModalResult = vbOK Then
    If Not SaveDoc(bAsk, frmSave.SaveModel, frmSave.SaveSP) Then
      CloseDoc = vbCancel
      Exit Function
    End If
  End If
  
  Set m_cx_Reports1 = Nothing
  Set m_cx_Reports2 = Nothing
  
  Set m_coll_SFSF = Nothing
  Set m_gn_GertNetMain = Nothing
  m_b_DirtySFSF = False
  m_l_CurrSFColl = -1
  
          
  CloseDoc = frmSave.ModalResult
  UpdateGUI Me
  
  Exit Function
ERR_H:
  CloseDoc = vbCancel
  UpdateGUI Me
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Friend Function AskFileName() As Boolean
  With frmMain.CommonDialog1
    .CancelError = True
    .DialogTitle = "Сохранить под новым именем"
    .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
    
'    Dim sarr() As String, sName As String
'    sarr = Split(m_s_NameDoc, "\", -1, vbTextCompare)
'    sName = sarr(UBound(sarr)): sarr(UBound(sarr)) = ""
'    .InitDir = Join(sarr, "\")
'    .FileName = sName
    
    Dim sDir As String, sName As String, sExt As String
    CutPath m_s_NameDoc, sDir, sName, sExt
    .InitDir = sDir
    .FileName = sName & sExt
    
    On Error GoTo ERR_H
    .ShowSave
    
    m_s_NameDoc = .FileName
              
    CutPath .FileName, sDir
    .InitDir = sDir
    
    
    AskFileName = True
    'm_b_NewDoc = False
          
    
    Exit Function
    
ERR_H:
    AskFileName = False
    'Err.Clear
  End With
End Function

Public Function NewDoc(ByVal bAsk As Boolean) As Boolean
  On Error GoTo ERR_H
  If Not (m_gn_GertNetMain Is Nothing) Or Not (m_coll_SFSF Is Nothing) Then
    If CloseDoc(bAsk) = vbCancel Then _
      NewDoc = False: Exit Function
  End If
  
  m_s_NameDoc = ""
  m_l_CurrSFColl = -1
  
  Dim ips As IPersistStorage
  Set m_gn_GertNetMain = New MGertNet
  Set ips = m_gn_GertNetMain
  ips.InitNew Nothing
  Set m_coll_SFSF = New CollectionX
  
  InitDefaultMG m_gn_GertNetMain
  InitDefaultSF m_coll_SFSF
  m_l_CurrSFColl = 1
  
  m_b_DirtySFSF = True
  
  NewDoc = True
  m_b_NewDoc = True
  m_s_NameDoc = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "noname.hzd"
  
  Set m_cx_Reports1 = New CollectionX
  Set m_cx_Reports2 = New CollectionX
  
  UpdateGUI Me
  
  m_gn_GertNetMain.Snapshot True
  
  Exit Function
ERR_H:
  NewDoc = False
  Set m_coll_SFSF = Nothing
  m_l_CurrSFColl = -1
  Set m_gn_GertNetMain = Nothing
  m_b_NewDoc = False
  
  Set m_cx_Reports1 = Nothing
  Set m_cx_Reports2 = Nothing
  
  UpdateGUI Me
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Public Function SaveDoc(ByVal bAsk As Boolean, Optional bModel As Boolean = True, Optional bCollSF As Boolean = True) As Boolean
  On Error GoTo ERR_H
      
  If m_b_NewDoc Then bCollSF = True: bModel = True
      
  If Not HaveDoc Then
    If bAsk Then _
      MsgBox "Нет документа", vbOKOnly Or vbExclamation, "Предупреждение"
    SaveDoc = False
  End If
      
  Dim stg As IStorage
  If m_b_NewDoc Then
    If Not AskFileName Then
      SaveDoc = False: Exit Function
    End If
    Set stg = CreateCF(m_s_NameDoc, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Else
    'On Error Resume Next
    Set stg = OpenCF(m_s_NameDoc, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
'    If Err.Number <> 0 Then
'      On Error GoTo ERR_H
'      Set stg = CreateCF(m_s_NameDoc, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
'    Else
'      On Error GoTo ERR_H
'    End If
  End If
      
  If Not (m_coll_SFSF Is Nothing) And bCollSF Then
    Dim iStg2 As IStorage
    On Error Resume Next
    Set iStg2 = OpenStorageInt(stg, "SFCollections", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    Dim bFlres As Boolean: bFlres = (Err.Number = 0)
    On Error GoTo ERR_H
    
    If Not bFlres Then _
      Set iStg2 = CreateStorageInt(stg, "SFCollections", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    
    SynchronizeStg iStg2, m_coll_SFSF
    m_b_DirtySFSF = False
  End If
  
  If Not (m_gn_GertNetMain Is Nothing) And bModel Then
    If m_b_NewDoc Or m_gn_GertNetMain.IsDirty Then
      Dim ips As IPersistStorage
      Set ips = m_gn_GertNetMain
      ips.Save stg, 1
    End If
  End If
      
  SaveDoc = True
  m_b_NewDoc = False
  Exit Function
ERR_H:
  SaveDoc = False
  
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Private Sub Class_Initialize()
  
  m_l_CurrSFColl = -1
  m_b_NewDoc = False
  m_b_DirtySFSF = False
  
  CreRepRun = True
  CreRepOpt = True
  CreRepVApply = True
  
  'RegisterShellext True

  'If haApp Is Nothing Then Set haApp = Me
  
  If haApp Is Nothing Then _
    Set haApp = Me
    
  If App.StartMode = vbSModeStandalone Then
    Load frmMain
    frmMain.Show
  End If
  
End Sub

Public Property Get Visible() As Boolean
  Visible = frmMain.Visible
End Property

Public Property Let Visible(ByVal Vis As Boolean)
  If Vis Then
    If Not frmMain.Visible Then
      Load frmMain
      frmMain.Show
    End If
  Else
    Unload frmMain
  End If
End Property

Public Sub Quit()
  If Not haApp Is Nothing Then
  'MsgBox "1"
    If haApp.IsCalcAny Then Err.Raise vbObjectError + 1, "Quit", "Невозможно завершить, пока идут фоновые вычисления"
  ' MsgBox "2"
    haApp.CloseDoc False
 ' MsgBox "3"
  End If
  Visible = False
'MsgBox "4"
  Set haApp = Nothing
End Sub


Private Sub Class_Terminate()

  Unload frmMain
  'Set frmMain = Nothing
  Set m_gn_GertNetMain = Nothing
  Set m_coll_SFSF = Nothing
  
End Sub

Public Function TestDirtySFSF() As Boolean
  If m_b_DirtySFSF Then _
    TestDirtySFSF = True: Exit Function
  
  If m_coll_SFSF Is Nothing Then _
    TestDirtySFSF = False: Exit Function
  
  TestDirtySFSF = False
  
  Dim collTmp As Variant
  For Each collTmp In m_coll_SFSF
    Dim collSF As collSF
    Set collSF = collTmp
    If collSF.IsDirty Then _
      TestDirtySFSF = True: Exit For
  Next collTmp

End Function

Private Sub SynchronizeStg(ByVal iStg As IStorage, ByVal coll As OTXRTLib.CollectionX)

  Dim collTmp As New CollectionX, lTmp As Long
  For lTmp = 1 To coll.Count
    collTmp.Add coll(lTmp), coll.Key(lTmp)
  Next lTmp

  Dim arrStgs() As String, arrStrm() As String
  DirOfStorage iStg, TDR_Stg, arrStgs, arrStrm
  
  On Error Resume Next
  Dim lU As Long: lU = UBound(arrStgs)
  Dim bFlErr As Boolean: bFlErr = Err.Number <> 0 Or lU < 0
  
  
  Dim cUpd As New CollectionX, cDel As New CollectionX
  If Not bFlErr Then
    Dim l As Long
    Set collTmp.DefaultItem = Nothing
    collTmp.AutoDefaultGetKey = True
    For l = LBound(arrStgs) To UBound(arrStgs)
      Dim csf As collSF
      Set csf = collTmp.Item(arrStgs(l))
      If csf Is Nothing Then
        cDel.Add arrStgs(l)
        collTmp.Remove arrStgs(l)
      Else
        If csf.IsDirty Then cUpd.Add csf, arrStgs(l)
        collTmp.Remove arrStgs(l)
      End If
    Next l
    collTmp.AutoDefaultGetKey = False
  End If
  
  Dim sNameDel As Variant
  For Each sNameDel In cDel
    Dim sNDel As String
    sNDel = sNameDel
    iStg.DestroyElement sNDel
  Next sNameDel
  
  Dim lTo As Long
  lTo = cUpd.Count
  For l = 1 To lTo
    Set csf = cUpd(l)
    Dim iStg2 As IStorage
    Set iStg2 = OpenStorageInt(iStg, cUpd.Key(l), STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    Dim ipsStor As IPersistStorage
    Set ipsStor = csf
    ipsStor.Save iStg2, 1
  Next l
  
  lTo = collTmp.Count
  For l = 1 To lTo
    Set csf = collTmp(l)
    Set iStg2 = CreateStorageInt(iStg, collTmp.Key(l), STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    Set ipsStor = csf
    ipsStor.Save iStg2, 1
  Next l

End Sub
'iStg2 , m_coll_SFSF


Private Sub m_gn_GertNetRun_OnCancel(ByVal dtTime As Date)
  If Not m_ic_Run Is Nothing Then _
    m_ic_Run.Cancel dtTime, m_gn_GertNetRun
  Set m_ic_Run = Nothing
End Sub

Private Sub m_gn_GertNetRun_OnEndOfWork(ByVal dtTime As Date)
  If Not m_ic_Run Is Nothing Then _
    m_ic_Run.EndOfWork dtTime, m_gn_GertNetRun
  Set m_ic_Run = Nothing
End Sub

Private Sub m_gn_GertNetRun_OnErrorCalc(ByVal bsErrMsg As String)
  m_ic_Run.ErrorCalc bsErrMsg, m_gn_GertNetRun
  Set m_ic_Run = Nothing
End Sub

Private Sub m_gn_GertNetRun_OnStepOfWork(ByVal shPercent As Integer)
  If Not m_ic_Run Is Nothing Then _
    m_ic_Run.StepOfWork shPercent
End Sub


Private Sub m_gn_GertNetRang_OnCancel(ByVal dtTime As Date)
  If Not m_ic_Rang Is Nothing Then _
    m_ic_Rang.Cancel dtTime, m_gn_GertNetRang
  Set m_ic_Rang = Nothing
End Sub
Private Sub m_gn_GertNetRang_OnErrorCalc(ByVal bsErrMsg As String)
  m_ic_Rang.ErrorCalc bsErrMsg, m_gn_GertNetRang
  Set m_ic_Rang = Nothing
End Sub


Private Sub m_gn_GertNetRang_OnEndOfWork(ByVal dtTime As Date)
  If Not m_ic_Rang Is Nothing Then _
    m_ic_Rang.EndOfWork dtTime, m_gn_GertNetRang
  'Set m_ic_Rang = Nothing
End Sub
Private Sub m_gn_GertNetRang_OnEndOfWork2(ByVal dtTime As Date)
  If Not m_ic_Rang Is Nothing Then _
    m_ic_Rang.EndOfWork2 dtTime, m_gn_GertNetRang
  Set m_ic_Rang = Nothing
End Sub


Private Sub m_gn_GertNetRang_OnStepOfWork(ByVal shPercent As Integer)
  If Not m_ic_Rang Is Nothing Then _
    m_ic_Rang.StepOfWork shPercent
End Sub
Private Sub m_gn_GertNetRang_OnStepOfWork2(ByVal shPercent As Integer)
  If Not m_ic_Rang Is Nothing Then _
    m_ic_Rang.StepOfWork2 shPercent
End Sub



Private Sub m_gn_GertNetOpt_OnCancel(ByVal dtTime As Date)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.Cancel dtTime, m_gn_GertNetOpt
  'Set m_ic_Opt = Nothing
End Sub
Private Sub m_gn_GertNetOpt_OnCancel3(ByVal dtTime As Date)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.Cancel3 dtTime, m_gn_GertNetOpt
  Set m_ic_Opt = Nothing
End Sub


Private Sub m_gn_GertNetOpt_OnErrorCalc(ByVal bsErrMsg As String)
  m_ic_Opt.ErrorCalc bsErrMsg, m_gn_GertNetOpt
  Set m_ic_Opt = Nothing
End Sub

Private Sub m_gn_GertNetOpt_OnEndOfWork(ByVal dtTime As Date)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.EndOfWork dtTime, m_gn_GertNetOpt
  'Set m_ic_Opt = Nothing
End Sub
Private Sub m_gn_GertNetOpt_OnEndOfWork3(ByVal dtTime As Date)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.EndOfWork3 dtTime, m_gn_GertNetOpt
  Set m_ic_Opt = Nothing
End Sub


Private Sub m_gn_GertNetOpt_OnStepOfWork(ByVal shPercent As Integer)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.StepOfWork shPercent
End Sub

Private Sub m_gn_GertNetOpt_OnStepOfWork3(ByVal shPercent As Integer)
  If Not m_ic_Opt Is Nothing Then _
    m_ic_Opt.StepOfWork3 shPercent
End Sub

