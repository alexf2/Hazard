VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Begin VB.Form frmAnswer 
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmAnswer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter sssplit 
      Height          =   4260
      Left            =   135
      TabIndex        =   5
      Top             =   210
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   7514
      _Version        =   131074
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmAnswer.frx":000C
      Begin PVTreeView3Lib.PVTreeViewX tv 
         Height          =   2220
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   6720
         _Version        =   393216
         _ExtentX        =   11853
         _ExtentY        =   3916
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   14811135
         BorderStyle     =   1
         Appearance      =   1
         BorderStyle     =   1
         StandardDefaultPicture=   10
         AlwaysShowSelection=   -1  'True
         AllowInPlaceEditing=   0   'False
         BackColor       =   14811135
         ForeColor       =   0
         EnableToolTips  =   0   'False
         LineColor       =   12632256
         DataMember      =   ""
         DataField0      =   ""
         DataField1      =   ""
         DataField2      =   ""
         DataField3      =   ""
         DataField4      =   ""
         DataField5      =   ""
         DataField6      =   ""
         DataField7      =   ""
         DataField8      =   ""
         DataField9      =   ""
         DataField10     =   ""
         DataField11     =   ""
         DataField12     =   ""
         DataField13     =   ""
         DataField14     =   ""
         DataField15     =   ""
         DataField16     =   ""
         DataField17     =   ""
         DataField18     =   ""
         DataField19     =   ""
      End
      Begin TrueOleDBGrid60.TDBGrid tdbg 
         Height          =   1965
         Left            =   0
         OleObjectBlob   =   "frmAnswer.frx":005E
         TabIndex        =   1
         Top             =   2295
         Width           =   6720
      End
   End
   Begin VB.Image imgGOST 
      Height          =   315
      Left            =   405
      Picture         =   "frmAnswer.frx":6AAF
      Top             =   15
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgTable 
      Height          =   195
      Left            =   15
      Picture         =   "frmAnswer.frx":6BD9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSCommand ssCompute 
      Height          =   525
      Left            =   60
      TabIndex        =   2
      Top             =   4665
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAnswer.frx":6CF7
      Caption         =   "Вычислить"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssCancel 
      Height          =   525
      Left            =   4965
      TabIndex        =   4
      Top             =   4620
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmAnswer.frx":6E61
      Caption         =   "Отменить"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssOk 
      Height          =   525
      Left            =   2580
      TabIndex        =   3
      ToolTipText     =   "Вносит изменения и закрывает диалог"
      Top             =   4635
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAnswer.frx":6FCB
      Caption         =   "Назначить"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FirstStart As Integer

Private m_b_Res As Boolean
Private m_b_LockResize As Boolean

Private m_cf As CollFTables
Private m_cfRoot As IFactorTable

Private m_mg As MGertNet
Private m_f As Factor
Private m_iStgF As IStorage
Private m_iStgGOST As IStorage

Private m_fValuationSumm As Variant
Private m_fAveWeightedSumm As Variant
Private m_sValDescr As String
Private m_s_MsgErr2 As String

Private m_myFont As StdFont
Private m_iftCurrentTable As IFactorTable

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400


Property Get ValueOfFactor() As Variant
  ValueOfFactor = m_fAveWeightedSumm
End Property

Private Function GetValDescr(ByVal f As Single) As String
  Dim le As LingvoEnum
  Set le = m_mg.GetEnumForFactor(m_f)
  GetValDescr = le.MarkForValue(le.RoundS(f / 10!))
End Function

Private Sub UpdateValue()
  If m_cfRoot Is Nothing Then
    m_fValuationSumm = Null
    m_fAveWeightedSumm = Null
    m_sValDescr = ""
    ssOk.Caption = "Назначить"
    ssOk.Enabled = False
  Else
    On Error GoTo ERR_H
    m_fValuationSumm = 0!: m_fAveWeightedSumm = 0!
    m_cfRoot.CalcValues m_fValuationSumm, m_fAveWeightedSumm
    
    m_sValDescr = GetValDescr(m_fAveWeightedSumm)
    ssOk.Caption = m_sValDescr & " (" & FormatNumber(m_fAveWeightedSumm, 2) & ")"
    ssOk.Enabled = True
  End If
  Exit Sub
  
ERR_H:
  ssOk.Caption = "Назначить"
  m_sValDescr = ""
  ssOk.Enabled = False
  m_fValuationSumm = Null: m_fAveWeightedSumm = Null
  Err.Raise Err.Number
End Sub

Private Sub Form_Initialize()
  m_b_LockResize = False
  
  With tdbg
       
     With .Styles
       .Item("Normal").VerticalAlignment = dbgVertCenter
       .Item("HighlightRow").VerticalAlignment = dbgVertCenter
       .Item("Selected").VerticalAlignment = dbgVertCenter
       '.Item("HighlightRow").BackColor = vbBlue
       '.Item("HighlightRow").ForeColor = vbWhite
     End With
     '.AnchorRightColumn = True
     '.ExtendRightColumn = True
     '.MarqueeStyle = dbgHighlightCell
     With .EditorStyle
       .Font.Bold = True
       .VerticalAlignment = dbgTop
       .WrapText = True
     End With
  
     With .Columns(0) 'Название
      .Alignment = dbgLeft
      .DefaultValue = "<Пусто>"
      .WrapText = True
      .Locked = True
      .HeadAlignment = dbgCenter
      .Font.Bold = False
      .FooterAlignment = dbgRight
      .FooterFont.Bold = True
     End With
     
     With .Columns(1) 'Блок
      .Alignment = dbgCenter
      .DefaultValue = False
      .WrapText = False
      .Locked = False
      .HeadAlignment = dbgCenter
      .Font.Bold = False
      .FooterAlignment = dbgCenter
      .FooterFont.Bold = True
     End With
     
     With .Columns(2) 'лингвистическая
      .Alignment = dbgCenter
      .DefaultValue = "<пусто>"
      .WrapText = False
      .Locked = False
      .HeadAlignment = dbgCenter
      .Font.Bold = True
      .FooterAlignment = dbgRight
      .FooterFont.Bold = True
     End With
     
     With .Columns(3) 'Бальная
      .Alignment = dbgCenter
      .DefaultValue = 1
      .WrapText = False
      .Locked = False
      .HeadAlignment = dbgCenter
      .Font.Bold = True
      .FooterAlignment = dbgCenter
      .FooterFont.Bold = False
     End With
     
     With .Columns(4) 'Ср. взв.
      .Alignment = dbgCenter
      .DefaultValue = 0
      .WrapText = False
      .Locked = True
      .HeadAlignment = dbgCenter
      .Font.Bold = False
      .FooterAlignment = dbgCenter
      .FooterFont.Bold = True
     End With
     
     With .Columns(5) 'Комментарий
      .Alignment = dbgLeft
      .DefaultValue = ""
      .WrapText = True
      .Locked = False
      .HeadAlignment = dbgCenter
      .Font.Bold = False
      .FooterAlignment = dbgLeft
      .FooterFont.Bold = False
     End With
     
     .AlternatingRowStyle = True
     .EvenRowStyle.BackColor = &H80FFFF
     .OddRowStyle.BackColor = &HC0FFFF
          
     .RowHeight = 2 * .RowHeight
   End With
End Sub

Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 7000 Then Width = 7000
    If Height < 5000 Then Height = 5000
  Else
    m_b_LockResize = False
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  sssplit.Move L_PAD, T_PAD, ScaleWidth - 2 * L_PAD, ScaleHeight - 3 * T_PAD - ssOk.Height
  ssCompute.Move sssplit.Left, sssplit.Top + sssplit.Height + T_PAD
  ssCancel.Move sssplit.Left + sssplit.Width - ssCancel.Width, ssCompute.Top
  ssOk.Move ssCompute.Left + ssCompute.Width + L_PAD / 2, ssCancel.Top, sssplit.Width - ssCancel.Width - ssCompute.Width - L_PAD
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  tv.StandardDefaultPicture = pvtNone
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
  'tv.ForeColor = &HFFFFFF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssCancel_Click
  End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
  'Set m_gt = Nothing
  
End Sub


Private Sub ssCancel_Click()
  m_b_Res = False
  If tdbg.EditActive Then _
    tdbg.CurrentCellModified = False: tdbg.EditActive = False
  Hide
End Sub

Private Sub ssCompute_Click()
  On Error GoTo ERR_H
  UpdateValue
  UpdateFooter
  If tdbg.ApproxCount > 0 Then
    tdbg.RefetchCol 2
    tdbg.RefetchCol 4
  End If
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Информация"
  Err.Clear
  UpdateFooter
End Sub

Private Sub ssOk_Click()

  On Error Resume Next
  If tdbg.EditActive Or tdbg.CurrentCellModified Then _
    tdbg.Update
    
  If Err.Number <> 0 Then Exit Sub
  On Error GoTo ERR_H
  
  Dim isci As IStCollItem, ips As IPersistStorage
  Set isci = m_cf
  If isci.Status <> OS_None Then
    Dim iStg As IStorage
    Set iStg = OpenStorageInt(m_iStgF, CStr(isci.Key), STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    Set ips = m_cf
    ips.Save iStg, 1
  End If
    
  m_b_Res = True
  Hide
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub ssCancel_LostFocus()
  HighLight2 ssCancel, False, Me.hWnd
End Sub

Private Sub ssCancel_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, True, Me.hWnd
End Sub

Private Sub ssCancel_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, False, Me.hWnd
End Sub


Private Sub ssCompute_LostFocus()
  HighLight2 ssCompute, False, Me.hWnd
End Sub

Private Sub ssCompute_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCompute, True, Me.hWnd
End Sub

Private Sub ssCompute_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCompute, False, Me.hWnd
End Sub


Private Sub ssOk_LostFocus()
  HighLight2 ssOk, False, Me.hWnd
End Sub

Private Sub ssOk_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, True, Me.hWnd
End Sub

Private Sub ssOk_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, False, Me.hWnd
End Sub


Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Friend Sub AssData(ByVal mg As MGertNet, ByVal f As Factor, ByVal sPath As String, ByVal PathToGostsStg As String, ByVal sCfgName As String)
  FirstStart = 1
  Set m_iftCurrentTable = Nothing
  m_b_Res = False
  m_fValuationSumm = Null
  m_fAveWeightedSumm = Null
  m_sValDescr = ""

        
  Set m_mg = mg
  Set m_f = f
  
  Dim ibs As IBSTRKey
  Set ibs = f
  
  'Caption = sCfgName & ": '" & ibs.Get() & "'"
  Caption = sCfgName & ": '" & f.Name & "'"
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  
  Set m_iStgF = OpenCF(sPath, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Set m_iStgGOST = OpenCF(PathToGostsStg, STGM_DIRECT1 Or STGM_READ1 Or STGM_SHARE_EXCLUSIVE1)
    
  Dim cllFactors As CollFactors
  Set cllFactors = New CollFactors
  Dim ips As IPersistStorage
  Dim ci As IStCollItem
  Set ci = cllFactors: Set ips = cllFactors
  ci.DefaultCU = True
  cllFactors.SetUpdateMode Nothing, False
  
  ips.Load m_iStgF
  Set m_cf = cllFactors(ibs.Get())
  Set m_cfRoot = m_cf(m_cf.DetectRoot())
  Set cllFactors = Nothing
  
'  Set ci = m_cf
'  m_cfRoot.NComponentName(0) = "TT"
'  Debug.Print m_cf(m_cf.DetectRoot()).NComponentName(0)
  
  Dim cllTopics As CollTopics
  Set cllTopics = New CollTopics
  Set ci = cllTopics: Set ips = cllTopics
  ci.DefaultCU = True
  cllTopics.SetUpdateMode Nothing, False
  
  ips.Load m_iStgGOST
  
  Dim ift As IFactorTable, isci As IStCollItem
  For Each ift In m_cf
    With ift
      If Not .IsWeightsCorrect Then
        Set isci = ift
        Err.Raise vbObjectError + 1, "frmAnswer.AssData", "У таблицы '" & isci.Name & "' ошибочно расставлены веса"
      End If
            
            
      Dim lNSlots As Long: lNSlots = ift.NumberSlots - 1
      Dim l As Long, lKT As Long, lKG As Long
      Dim obj As Object, sPth(1) As String
      
      For l = lNSlots To 0 Step -1
      
        If .NSlotType(l) = FT_Gost Then
          .NSlotGost l, lKT, lKG
          sPth(0) = CStr(lKT): sPth(1) = CStr(lKG)
          On Error GoTo ERR_GOST
          Set obj = FetchGostTable(m_iStgGOST, sPth, STGM_DIRECT1 Or STGM_READ1 Or STGM_SHARE_EXCLUSIVE1, STGM_DIRECT1 Or STGM_READ1 Or STGM_SHARE_EXCLUSIVE1, True)
          On Error GoTo 0
          GoTo L_Next1
ERR_GOST:
          Err.Raise Err.Number, Err.Source, "Ошибка загрузки норматива: " & Err.Description
L_Next1:
          
        ElseIf .NSlotType(l) = FT_Self Then
          .NSlotSelf l, lKT
          Set obj = m_cf(lKT)
          
        Else
          Set isci = ift
          Err.Raise vbObjectError + 1, "frmAnswer.AssData", "У таблицы '" & isci.Name & "' имеются пустые слоты"
          
        End If
        
        .NBindSlot l, obj
        
      Next l
      
    End With
  Next ift
  
  On Error Resume Next
  UpdateValue
  On Error GoTo 0
  
  LoadTv
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  If pv.IsValid Then pv.Open pvtAllChildren

End Sub

Private Sub LoadTv()
  Dim pv As PVBranch
  With tv
    On Error GoTo ERR_H
    
    HandsOffTree
    .Redraw = False
    .Branches.Clear
              
    Dim isci As IStCollItem
    Set isci = m_cfRoot
        
    Set pv = .Branches.Add(pvtPositionInOrder, 0, isci.Name)
    LoadToBr pv, m_cfRoot, 0
      
    .Redraw = True
  End With
  
  Exit Sub
ERR_H:
  tv.Redraw = True
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub LoadToBr(ByVal pv As PVBranch, ByVal ift As IFactorTable, ByVal lIdx As Long)
  With pv
    If .Level = 0 Then
      .ForeColor = RGB(0, 128, 128)
    Else
      .ForeColor = RGB(0, 0, 255)
    End If
    Set .CustomItemPicture = imgTable.Picture
    
    Dim tbItem As CTableItem: Set tbItem = New CTableItem
    Set tbItem.Table = ift: tbItem.Index = lIdx
    Set .DataVariant = tbItem: Set tbItem = Nothing
    
    .Data = Node_Table
  End With
  
  Dim pv2 As PVBranch
  With ift
    Dim lNSlt As Long: lNSlt = .NumberSlots - 1
    Dim l As Long
    For l = 0 To lNSlt
      Set pv2 = pv.Add(pvtPositionInOrder, 0, .NComponentName(l))
                
      Select Case .NSlotType(l)
        Case FT_Gost
          LoadGOST pv2, ift, l
                        
        Case FT_Self
          Dim ift2 As IFactorTable
          .NSlotSelf l, , ift2
          LoadToBr pv2, ift2, l
          
        Case Else
          Err.Raise vbObjectError + 1, "frmAnswer.LoadToBr", "Внутренняя ошибка"
      End Select
    Next l
  End With
End Sub

Private Sub LoadGOST(ByVal pv2 As PVBranch, ByVal ift As IFactorTable, ByVal l As Long)
  With ift
    Dim isci As IStCollItem, ig As IGostTable
    .NSlotGost l, , , ig
    Set isci = ig
    
    With pv2
      .Text = ift.NComponentName(l) & " [ " & ig.ClmName & " ]"
      Set .CustomItemPicture = imgGOST.Picture
      .ForeColor = RGB(0, 0, 0)
      Set .Font = m_myFont
      
      Dim gtItem As CGOSTItem: Set gtItem = New CGOSTItem
      Set gtItem.GOST = ig: gtItem.Index = l
      Set .DataVariant = gtItem: Set gtItem = Nothing
          
      .Data = Node_GOST
    End With
  End With
    
  Dim pvT As PVBranch
  With ig
    Dim lNsl As Long: lNsl = .NumberSlots - 1
    Dim l2 As Long, bFlFounded As Boolean
    bFlFounded = False
    For l2 = 0 To lNsl
      Set pvT = pv2.Add(pvtPositionInOrder, 0, .NDescr(l2))
      With pvT
        .ForeColor = RGB(0, 0, 0)
        .Data = Node_GOST_Ctx
        .DataVariant = Null
        .CheckBoxType = pvtChkBoxRound
        .CheckBoxState = IIf(ift.NSlotValue(l) = l2, pvtChkBoxStateChecked, pvtChkBoxStateNotChecked)
        If .CheckBoxState = pvtChkBoxStateChecked Then bFlFounded = True
        .DataVariant = l2
      End With
    Next l2
    
    If Not bFlFounded And ift.NSlotValue(l) <> -1 Then ift.NSlotValue(l) = -1
    
  End With
End Sub


Friend Sub UnassData()
  tdbg.Close True
  Set m_cf = Nothing
  Set m_cfRoot = Nothing
  Set m_mg = Nothing
  Set m_f = Nothing
  HandsOffTree
  tv.Branches.Clear
  Set m_iftCurrentTable = Nothing
  Set m_iStgF = Nothing
  Set m_iStgGOST = Nothing
End Sub

Private Function IsVarEmpty2(ByRef v As Variant) As Boolean
  On Error GoTo ERR_H
  If v Is Nothing Then IsVarEmpty2 = True Else IsVarEmpty2 = False
  Exit Function
ERR_H:
  IsVarEmpty2 = True
  Err.Clear
End Function


Private Sub HandsOffTree()
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  Do While pv.IsValid
    With pv
      If .Data = Node_Table Then
        If Not IsVarEmpty2(.DataVariant) Then _
          .DataVariant.Table.HandsOffObjects
        Set .DataVariant = Nothing
        
      ElseIf .Data = Node_GOST Then
        Set .DataVariant = Nothing
      End If
    End With
        
    Set pv = pv.Get(pvtGetNext, 0)
  Loop
End Sub



Private Sub tdbg_BeforeUpdate(Cancel As Integer)
  With tdbg
    If .Columns(2).DataChanged Then
      If IsNull(.Columns(2).Value) Or IsEmpty(.Columns(2).Value) Then _
        m_s_MsgErr2 = "Лингвистическая оценка не должна быть пустой": GoTo ERR_H
      .Columns(2).Value = Trim(.Columns(2).Value)
    End If
    
    If .Columns(3).DataChanged Then
      If IsNull(.Columns(3).Value) Or IsEmpty(.Columns(3).Value) Then _
        m_s_MsgErr2 = "Бальная оценка обязательна": GoTo ERR_H
    End If
    
    If .Columns(5).DataChanged Then
      If Not IsNull(.Columns(5).Value) Then _
        .Columns(5).Value = Trim(.Columns(5).Value)
    End If
    
    Exit Sub
    
ERR_H:
    Cancel = True
    tdbg.PostMsg 1
  End With
End Sub

Private Sub tdbg_Change()

  With tdbg
    If .Col = 1 Then
      If Not IsNull(.Bookmark) And Not m_iftCurrentTable Is Nothing Then
         'm_iftCurrentTable.NLocked(.Bookmark) = .Columns(1).Value
         'LockClms
         If .Columns(1).Value Then
           .Columns(2).Locked = False
           .Columns(3).Locked = False
           .Columns(5).Locked = False
         Else
           .Columns(2).Locked = True
           .Columns(3).Locked = True
           .Columns(5).Locked = True
        End If
      End If
    End If
  End With
End Sub

Private Sub LockClms()

  If m_iftCurrentTable Is Nothing Then Exit Sub
  With tdbg
    If m_iftCurrentTable.NLocked(.Bookmark) Then
      .Columns(2).Locked = False
      .Columns(3).Locked = False
      .Columns(5).Locked = False
    Else
      .Columns(2).Locked = True
      .Columns(3).Locked = True
      .Columns(5).Locked = True
    End If
  End With
End Sub

Private Sub tdbg_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  With tdbg
    If .ApproxCount < 1 Then Exit Sub
    If LastRow <> .Bookmark Then
      LockClms
    End If
  End With
End Sub

Private Sub tv_AfterCheckBoxStateChanged(ByVal node As PVTreeView3Lib.PVIBranch)

  Dim pv As PVBranch, pv2 As PVBranch
  Set pv = node.Get(pvtGetParent, 0)
  Set pv2 = pv.Get(pvtGetParent, 0)
  If pv.Data <> Node_GOST Or pv2.Data <> Node_Table Then
    MsgBox "tv_AfterCheckBoxStateChanged: Внутренняя ошибка", vbCritical Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If node.CheckBoxState = pvtChkBoxStateChecked Then
  
    pv2.DataVariant.Table.NSlotValue(pv.DataVariant.Index) = node.DataVariant
    
    If Not pv2.DataVariant.Table.NLocked(pv.DataVariant.Index) Then _
      If Not (IsEmpty(tdbg.FirstRow) Or IsNull(tdbg.FirstRow)) Then _
        tdbg.RefetchRow tdbg.FirstRow + tdbg.Row
    
    UpdateFooter
    If ssOk.Enabled Then ssOk.Enabled = False
  Else
    'pv2.DataVariant.Table.NSlotValue(pv.DataVariant.Index) = -1
  End If
    
End Sub

Private Sub tdbg_AfterUpdate()
  UpdateFooter
End Sub

Private Function CurrentTable() As IFactorTable
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If pv.IsValid Then
    If pv.Data = Node_GOST Then
      Set pv = pv.Get(pvtGetParent, 0)
    ElseIf pv.Data = Node_GOST_Ctx Then
      Set pv = pv.Get(pvtGetParent, 0).Get(pvtGetParent, 0)
    ElseIf pv.Data = Node_Table Then
      If pv.Level <> 0 Then _
        Set pv = pv.Get(pvtGetParent, 0)
    End If
    
    If pv.Data <> Node_Table Then _
      Err.Raise vbObjectError + 1, "CurrentTable", "CurrentTable: внутренняя ошибка"
    
    Set CurrentTable = pv.DataVariant.Table
  Else
    Set CurrentTable = Nothing
  End If
End Function

Private Sub tdbg_UnboundReadDataEx( _
  ByVal RowBuf As TrueOleDBGrid60.RowBuffer, _
  StartLocation As Variant, _
  ByVal offset As Long, _
  ApproximatePosition As Long)

  If m_iftCurrentTable Is Nothing Then
    RowBuf.RowCount = 0
    Exit Sub
  End If
  
  Dim isci As IStCollItem
  Dim lRows As Long, lBookmark As Long, lStaRow As Long, l As Long
  Dim lRowsFetched As Long, lMaxRowNum As Long
  Dim lMaxRowIdx As Long, lMaxColumnIdx As Long
      
  lMaxRowNum = m_iftCurrentTable.NumberSlots
  lMaxRowIdx = RowBuf.RowCount - 1
  lMaxColumnIdx = RowBuf.ColumnCount - 1
  
  lRowsFetched = 0
  If IsNull(StartLocation) Then
      ' StartLocation reffers to either BOF (-1) or EOF (MaxRow)
      lStaRow = IIf(offset < 0, lMaxRowNum + offset, -1 + offset)
  Else
      ' StartLocation is an actual bookmark
      lStaRow = StartLocation + offset
  End If

  With RowBuf
    For l = 0 To lMaxRowIdx
      lBookmark = lStaRow + l
      If lBookmark < 0 Or lBookmark >= lMaxRowNum Then Exit For
      
      Dim lColIndex As Long, j As Long
      For j = 0 To lMaxColumnIdx
        lColIndex = .ColumnIndex(l, j)
        Select Case lColIndex
          Case 0
            .Value(l, 0) = m_iftCurrentTable.NComponentName(lBookmark)
          Case 1
            .Value(l, 1) = m_iftCurrentTable.NLocked(lBookmark)
          Case 2
            .Value(l, 2) = m_iftCurrentTable.NLingvoVal(lBookmark)
          Case 3
            .Value(l, 3) = m_iftCurrentTable.NValuation(lBookmark)
          Case 4
            .Value(l, 4) = m_iftCurrentTable.NAveWeighted(lBookmark)
          Case 5
            If m_iftCurrentTable.NLocked(lBookmark) Then
              .Value(l, 5) = m_iftCurrentTable.NComment(lBookmark)
            Else
              If m_iftCurrentTable.NSlotType(lBookmark) = FT_Gost Then
                Dim gt As IGostTable
                m_iftCurrentTable.NSlotGost lBookmark, , , gt
                Set isci = gt
                .Value(l, 5) = isci.Name
              End If
            End If
        End Select
      Next j
      
      
      .Bookmark(l) = lBookmark
      lRowsFetched = lRowsFetched + 1
    Next l
  End With
  
  RowBuf.RowCount = lRowsFetched
  If lStaRow >= 0 Then ApproximatePosition = lStaRow
End Sub

Private Sub tdbg_UnboundWriteData( _
  ByVal RowBuf As TrueOleDBGrid60.RowBuffer, _
  WriteLocation As Variant)

  Dim Col As Integer
  On Error GoTo NoWrite
  
  If m_iftCurrentTable Is Nothing Or RowBuf.ColumnCount < 1 Then
    RowBuf.RowCount = 0
  Else
    'm_iftCurrentTable.NWeight(WriteLocation) = RowBuf.Value(0, 0)
    With m_iftCurrentTable
      If Not IsNull(RowBuf.Value(0, 1)) Then
        .NLocked(WriteLocation) = RowBuf.Value(0, 1)
      End If
        
      If .NLocked(WriteLocation) Then
        If Not IsNull(RowBuf.Value(0, 2)) Then _
          .NLingvoVal(WriteLocation) = RowBuf.Value(0, 2)
          
        If Not IsNull(RowBuf.Value(0, 3)) Then
          .NValuation(WriteLocation) = CSng(RowBuf.Value(0, 3))
          ssOk.Enabled = False
        End If
          
        If Not IsNull(RowBuf.Value(0, 5)) Then _
          .NComment(WriteLocation) = RowBuf.Value(0, 5)
      End If
    End With
  End If
    
  Exit Sub
NoWrite:
    RowBuf.RowCount = 0
    m_s_MsgErr2 = Err.Description
    tdbg.PostMsg 1
    'MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
    Err.Clear
End Sub


Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  With tdbg
    On Error GoTo ERR_H
    
    Dim vt As IFactorTable
    Set vt = CurrentTable()

        
    If vt Is m_iftCurrentTable Then
      If node.Data = Node_GOST Or node.Data = Node_Table Then
        .Bookmark = node.DataVariant.Index
      ElseIf node.Data = Node_GOST_Ctx Then
        .Bookmark = node.Get(pvtGetParent, 0).DataVariant.Index
      End If
      
      Exit Sub
      
    End If
          

    'tdbgK.ReBind
      
    Set m_iftCurrentTable = vt
    If Not vt Is Nothing Then
      On Error Resume Next
      .Bookmark = Null
      On Error GoTo ERR_H
      .ReBind
      .ApproxCount = vt.NumberSlots
      UpdateFooter
    Else
      On Error Resume Next
      .Bookmark = Null
      On Error GoTo ERR_H
      .ReBind
      .ApproxCount = 0
      UpdateFooter
    End If

    If Not m_iftCurrentTable Is Nothing Then
      If m_iftCurrentTable.NumberSlots > 0 Then
        If node.Data = Node_GOST Or node.Data = Node_Table Then
          .Bookmark = node.DataVariant.Index
        ElseIf node.Data = Node_GOST_Ctx Then
          .Bookmark = node.Get(pvtGetParent, 0).DataVariant.Index
        End If
      End If
    End If
    
  End With
    
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Информация"
  Err.Clear
End Sub

Private Sub UpdateFooter()
  With tdbg
    If m_iftCurrentTable Is Nothing Then
      '.Columns(2).FooterText = ""
      .Columns(3).FooterText = ""
      .Columns(4).FooterText = ""
    Else
      Dim fValSumm As Variant, fAveWeightedSumm As Variant
      fValSumm = m_iftCurrentTable.ValuationSumm
      fAveWeightedSumm = m_iftCurrentTable.AveWeightedSumm
      .Columns(3).FooterText = IIf(IsEmpty(fValSumm), "<?>", FormatNumber(fValSumm, 2))
      .Columns(4).FooterText = IIf(IsEmpty(fAveWeightedSumm), "<?>", FormatNumber(fAveWeightedSumm, 2))
    End If
  End With
End Sub

Private Sub tdbg_LostFocus()
    
  On Error Resume Next
  'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
  With tdbg
    If .CurrentCellModified Then
      .Update
      If Err.Number <> 0 Then
        .CurrentCellModified = False
        .EditActive = False
      End If
    End If
  End With
    
End Sub

Private Sub tdbg_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 3662 Or DataError = 16389 Then
    Response = 0
    'tdbgCompl.PostMsg 1
    MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
    DoEvents
    m_s_MsgErr2 = ""
  End If
End Sub

Private Sub tdbg_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = vbKeyReturn Then
    tdbg.Update
  End If
End Sub


Private Sub tdbg_Validate(Cancel As Boolean)
   On Error Resume Next
   tdbg.Update
   If Err.Number <> 0 Then Cancel = True
End Sub


Private Sub tdbg_PostEvent(ByVal MsgID As Integer)
  Select Case MsgID
    Case 1
      MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
      m_s_MsgErr2 = ""
  End Select
End Sub

