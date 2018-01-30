VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Begin VB.Form frmEditFT 
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   Icon            =   "frmEditFT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin SSSplitter.SSSplitter split1 
      Height          =   2565
      Left            =   495
      TabIndex        =   2
      Top             =   135
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   4524
      _Version        =   131074
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmEditFT.frx":0442
      Begin PVTreeView3Lib.PVTreeViewX tv 
         Height          =   2565
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5460
         _Version        =   393216
         _ExtentX        =   9631
         _ExtentY        =   4524
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   9450828
         BorderStyle     =   1
         Appearance      =   1
         BorderStyle     =   1
         StandardDefaultPicture=   10
         AlwaysShowSelection=   -1  'True
         AllowInPlaceEditing=   0   'False
         BackColor       =   9450828
         ForeColor       =   0
         EnableToolTips  =   0   'False
         SelectedTextBackColor=   16711680
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
   End
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   1590
      Top             =   4695
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
      WindowText1     =   "frmEditFT"
   End
   Begin VB.TextBox txtInPlace 
      BackColor       =   &H00FFFBF0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2175
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "Вызов редактора - F2"
      Top             =   4665
      Visible         =   0   'False
      Width           =   2250
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgK 
      Height          =   1635
      Left            =   1245
      OleObjectBlob   =   "frmEditFT.frx":0474
      TabIndex        =   4
      Top             =   2745
      Visible         =   0   'False
      Width           =   4335
   End
   Begin Threed.SSCommand ssClose 
      Height          =   525
      Left            =   4560
      TabIndex        =   1
      Top             =   4680
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmEditFT.frx":2DE2
      Caption         =   "Закрыть"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin VB.Image imgBadGOST 
      Height          =   315
      Left            =   375
      Picture         =   "frmEditFT.frx":2F4C
      Top             =   4560
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgEmpty 
      Height          =   225
      Left            =   630
      Picture         =   "frmEditFT.frx":30CA
      Top             =   4905
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTable 
      Height          =   195
      Left            =   765
      Picture         =   "frmEditFT.frx":31C4
      Top             =   4575
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgGOST 
      Height          =   315
      Left            =   1140
      Picture         =   "frmEditFT.frx":32E2
      Top             =   4590
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmEditFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_Res As Boolean
Private m_b_LockResize As Boolean
Private m_cf As CollFTables
Public FirstStart As Integer

Private m_bSavedOK As Boolean
Private m_pvEdit As PVBranch
Private m_lEditSlot As Long
Private m_iftCurrentTable As IFactorTable
Private m_s_MsgErr2 As String
Private m_lLeftOfCell As Long

Private m_myFont As StdFont

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 100
Private Const L_PAD2 = 100


Private Enum Cmds
  cmTAddNew
  cmTAssEmpty
  cmTAssGOST
  cmTAssTable
  cmTRefresh
  CmTRemove
  CmTRename
  cmTSave
  cmTUniform
  cmTAssWeights
  cmTCheckW
  cmTNormalize
End Enum


Private Sub Form_Initialize()
  m_b_LockResize = False
  
  With GTMsgHook1
    .Windows.Item(1).Messages.Add _
      Key:="WM_PENDING_OP", Value:=WM_PENDING_OP, ProcessDefaultProc:=False
  
    .Subclass = True
  End With
  
End Sub

Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 5000 Then Width = 5000
    If Height < 3000 Then Height = 3000
  Else
    m_b_LockResize = False
    Exit Sub
  End If
  
  On Error GoTo ERR_H
    
  split1.Move L_PAD, T_PAD, ScaleWidth - 2 * L_PAD, ScaleHeight - 3 * T_PAD - ssClose.Height
  ssClose.Move split1.Left, split1.Top + split1.Height + T_PAD, split1.Width
    
'  tv.Move L_PAD, T_PAD, ScaleWidth - (3 / 2) * L_PAD - sspPercent.Width, ScaleHeight - 3 * T_PAD - ssClose.Height
'  sspPercent.Move tv.Left + tv.Width + L_PAD / 2, tv.Top, sspPercent.Width, tv.Height
'  ssClose.Move tv.Left, tv.Top + tv.Height + T_PAD, ScaleWidth - 2 * L_PAD
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    'ssOk_Click
  ElseIf KeyAscii = vbKeyReturn Then
    'ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  m_s_MsgErr2 = ""
  'Set tv.CustomDefaultPicture = Nothing
  tv.StandardDefaultPicture = pvtNone
  
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
    
  With tdbgK
    With .Styles
      .Item("Normal").VerticalAlignment = dbgVertCenter
      .Item("HighlightRow").VerticalAlignment = dbgVertCenter
      .Item("Selected").VerticalAlignment = dbgVertCenter
      '.Item("HighlightRow").BackColor = vbBlue
      '.Item("HighlightRow").ForeColor = vbWhite
    End With
    '.AnchorRightColumn = True
    '.ExtendRightColumn = True
    .MarqueeStyle = dbgHighlightCell
    With .EditorStyle
      .Font.Bold = True
      .Alignment = dbgCenter
      .BackColor = &H80FFFF
      .ForeColor = &H0
      .VerticalAlignment = dbgVertCenter
      .WrapText = True
    End With
  
    With .Columns(0)
     .Alignment = dbgLeft
     .Locked = False
     .DefaultValue = 0.1
     .HeadAlignment = dbgCenter
     .OwnerDraw = True
    End With
    .EvenRowStyle.BackColor = &H80FFFF
    .OddRowStyle.BackColor = &HC0FFFF
    .AlternatingRowStyle = True
    .RowHeight = 1.5 * .RowHeight
      
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    'txtInPlace_LostFocus
    On Error GoTo ERR_H
    With tdbgK
      If .Visible Then
        If .EditActive Or .CurrentCellModified Then _
         .Update
      End If
    End With
    ssClose_Click
  End If
  Exit Sub
ERR_H:
  Dim vbRes As VbMsgBoxResult
  If m_s_MsgErr2 = "" Then
    vbRes = MsgBox("[" & Err.Description & "] - Закрыть ?", vbExclamation Or vbYesNo, "Ошибка")
    If vbRes = vbYes Then ssClose_Click
  End If
End Sub


Private Sub Form_Terminate()
  GTMsgHook1.Subclass = False
End Sub

Private Sub GTMsgHook1_Message(ByVal hWnd_ As Long, ByVal MsgID As Long, ByVal wParam As Long, ByVal lParam As Long)
  If hWnd_ <> Me.hWnd Then Exit Sub
  If MsgID <> WM_PENDING_OP Then Exit Sub
    
  On Error GoTo ERR_H
    
  Select Case lParam
    Case cmTAddNew
      CmdTAddNew
    Case cmTAssEmpty
      CmdTAssEmpty
    Case cmTAssGOST
      CmdTAssGOST
    Case cmTAssTable
      CmdTAssTable
    Case cmTRefresh
      CmdTRefresh
    Case CmTRemove
      CmdTRemove
    Case CmTRename
      CmdTRename
    Case cmTSave
      CmdTSave
    Case cmTUniform
      CmdTUniform
    Case cmTAssWeights
      CmdTAssWeights
    Case cmTCheckW
      CmdTCheckW
    Case cmTNormalize
    CmdTNormalize
  End Select
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub ssClose_Click()
  CmdTSave
  If Not m_bSavedOK Then
    Dim vbr As VbMsgBoxResult
    vbr = MsgBox("Произошла ошибка при сохранении. Закрыть окно ?", vbQuestion Or vbYesNo, "Запрос")
    If vbr <> vbYes Then Exit Sub
  End If
  
  m_b_Res = True
  Me.Hide
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub ssClose_LostFocus()
  HighLight2 ssClose, False, Me.hWnd
End Sub

Private Sub ssClose_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, True, Me.hWnd
End Sub

Private Sub ssClose_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, False, Me.hWnd
End Sub


Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property


Friend Sub AssData(ByVal cf As CollFTables)
  FirstStart = 1
  m_b_Res = False
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  
  Set m_cf = cf
  
  Dim isci As IStCollItem
  Set isci = cf
  
  If cf.Count < 1 Then
    cf.Add "Оценка " & isci.Name
  End If
  
  LoadToTV
  
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  If pv.IsValid Then pv.Open pvtAllChildren
  
  Set m_iftCurrentTable = CurrentTable()
End Sub

Private Sub LoadGOST(ByVal pv2 As PVBranch, ByVal ift As IFactorTable, ByVal l As Long)
  With ift
    Dim isci As IStCollItem, lKT As Long, lKG As Long
    Dim ig As IGostTable
    .NSlotGost l, lKT, lKG
    Set isci = frmGostEdit.DirectFGOST(lKT, lKG)
    
    With pv2
      If Not isci Is Nothing Then
         .Text = ift.NComponentName(l) & " (" & isci.Name & ")"
         Set .CustomItemPicture = imgGOST.Picture
      Else
        .Text = ift.NComponentName(l)
        Set .CustomItemPicture = imgBadGOST.Picture
        pv2.Clear
        Exit Sub
      End If
      .DataVariant = Null
      .Data = Node_GOST
      Set .Font = m_myFont
      .ForeColor = RGB(14, 239, 61)
    End With
  End With
  
  Set ig = isci
  Dim pvT As PVBranch
  With ig
    Dim lNsl As Long: lNsl = .NumberSlots - 1
    For l = 0 To lNsl
      Set pvT = pv2.Add(pvtPositionInOrder, 0, .NDescr(l) & " = " & .NValue(l))
      With pvT
        .Data = Node_GOST_Ctx
        .DataVariant = Null
        .ForeColor = RGB(108, 245, 250)
      End With
      'pvT.ForeColor = RGB(14, 239, 61)
    Next l
  End With
End Sub

Private Sub LoadToBr(ByVal pv As PVBranch, ByVal ift As IFactorTable)
  With pv
    If .Level = 0 Then
      .ForeColor = RGB(237, 232, 83)
    Else
      '.ForeColor = RGB(108, 245, 250)
      .ForeColor = RGB(14, 239, 61)
    End If
    Set .CustomItemPicture = imgTable.Picture
    Set .DataVariant = ift
    .Data = Node_Table
  End With
  
  Dim pv2 As PVBranch
  With ift
    Dim lNSlt As Long: lNSlt = .NumberSlots - 1
    Dim l As Long
    For l = 0 To lNSlt
      Set pv2 = pv.Add(pvtPositionInOrder, 0, .NComponentName(l))
      pv2.ForeColor = RGB(108, 245, 250)
      'pv2.ForeColor = RGB(14, 239, 61)
      Select Case .NSlotType(l)
        Case FT_Gost
          LoadGOST pv2, ift, l
          
        Case FT_None
          Set pv2.CustomItemPicture = imgEmpty.Picture
          pv2.DataVariant = Null
          pv2.Data = Node_Empty
          
        Case FT_Self
          Dim lKT As Long
          .NSlotSelf l, lKT
          LoadToBr pv2, m_cf(lKT)
      End Select
    Next l
  End With
End Sub

Private Sub LoadToTV()
  Dim pv As PVBranch
  With tv
    On Error GoTo ERR_H
    
    .Redraw = False
    .Branches.Clear
              
    Dim ift As IFactorTable, isci As IStCollItem
    Set ift = m_cf(m_cf.DetectRoot)
    Set isci = ift
        
    Set pv = .Branches.Add(pvtPositionInOrder, 0, isci.Name)
    LoadToBr pv, ift
      
    .Redraw = True
  End With
  
  Exit Sub
ERR_H:
  tv.Redraw = True
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Friend Sub UnassData()
  On Error Resume Next
  tv.Branches.Clear
  tdbgK.Close True
  Set m_cf = Nothing
  Set m_pvEdit = Nothing
  Set m_iftCurrentTable = Nothing
End Sub


Friend Sub PostTNormalize()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTNormalize
End Sub


Friend Sub PostTCheckW()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTCheckW
End Sub

Friend Sub PostTAddNew()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTAddNew
End Sub
Friend Sub PostTAssEmpty()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTAssEmpty
End Sub
Friend Sub PostTAssGOST()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTAssGOST
End Sub
Friend Sub PostTAssTable()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTAssTable
End Sub
Friend Sub PostTRefresh()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTRefresh
End Sub
Friend Sub PostTRemove()
  PostMessage hWnd, WM_PENDING_OP, 0, CmTRemove
End Sub
Friend Sub PostTRename()
  PostMessage hWnd, WM_PENDING_OP, 0, CmTRename
End Sub
Friend Sub PostTSave()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTSave
End Sub
Friend Sub PostTUniform()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTUniform
End Sub

Friend Sub PostTAssWeights()
  PostMessage hWnd, WM_PENDING_OP, 0, cmTAssWeights
End Sub

Private Sub tdbgK_AfterUpdate()
  UpdateFooter
End Sub

Private Sub tv_RButtonUp(ByVal node As PVTreeView3Lib.PVIBranch, ByVal X_ As Single, ByVal Y_ As Single)
  Dim pv As PVBranch
  Dim x As Single, y As Single
  x = ScaleX(X_, vbTwips, vbPixels): y = ScaleX(Y_, vbTwips, vbPixels)
  Set pv = tv.HitTest(x, y)
  
  With frmShadow
  
    If pv.IsValid Then
      pv.Select pvtSelectNode
                            
      Select Case pv.Data
        Case Node_GOST, Node_Table, Node_Empty
          .mnuTAddNew.Enabled = True
          .mnuTRename.Enabled = True
          .mnuTAssGOST.Enabled = True
          .mnuTAssTable.Enabled = True
          .mnuTAssEmpty.Enabled = True
          .mnuTRefresh.Enabled = True
          .mnuTSave.Enabled = True
          .mnuTRemove.Enabled = True
          .mnuTUniform.Enabled = True
          .mnuTAssWeights = True
          .mnuTCheckW = True
          .mnuTNormalize = True
          
        Case Node_GOST_Ctx
          .mnuTAddNew.Enabled = True
          .mnuTRename.Enabled = False
          .mnuTAssGOST.Enabled = False
          .mnuTAssTable.Enabled = False
          .mnuTAssEmpty.Enabled = False
          .mnuTRefresh.Enabled = True
          .mnuTSave.Enabled = True
          .mnuTRemove.Enabled = False
          .mnuTUniform.Enabled = True
          .mnuTAssWeights = True
          .mnuTCheckW = True
          .mnuTNormalize = True
      End Select
                        
    Else
      .mnuTAddNew.Enabled = False
      .mnuTRename.Enabled = False
      .mnuTAssGOST.Enabled = False
      .mnuTAssTable.Enabled = False
      .mnuTAssEmpty.Enabled = False
      .mnuTRefresh.Enabled = True
      .mnuTSave.Enabled = True
      .mnuTRemove.Enabled = False
      .mnuTUniform.Enabled = IIf(tv.Count < 1, False, True)
      .mnuTAssWeights = .mnuTUniform.Enabled
      .mnuTCheckW = .mnuTUniform.Enabled
      .mnuTNormalize = .mnuTUniform.Enabled
    End If
    
    PopupMenu .mnuTabEditor, vbPopupMenuLeftAlign, X_, Y_
        
  End With
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim pv As PVBranch
  If KeyCode = vbKeyF5 Then
    KeyCode = 0: CmdTRefresh
    
  ElseIf KeyCode = vbKeyDelete Then
    KeyCode = 0: CmdTRemove
    
  ElseIf KeyCode = vbKeyF2 Then
    KeyCode = 0: CmdTRename
    
  ElseIf KeyCode = vbKeyF10 Then
    KeyCode = 0: CmdTNormalize
    
  End If
End Sub


Private Sub CmdTAddNew()
  Dim pv As PVBranch
  Set pv = ParentTableBr(tv.Branches.Get(pvtGetNextSelected, 0))
  If pv Is Nothing Then Exit Sub
    
  Dim ift As IFactorTable
  Set ift = pv.DataVariant
  ift.InsertSlots 1
  
  Set pv = pv.Add(pvtPositionInOrder, 0, ift.NComponentName(ift.NumberSlots - 1))
  With pv
    .ForeColor = RGB(108, 245, 250)
    Set .CustomItemPicture = imgEmpty.Picture
    .DataVariant = Null
    .Data = Node_Empty
    
    .Select pvtSelectNode
    .EnsureVisible
    CmdTRename
  End With
End Sub
    
Private Sub GatherTables(ByVal pv As PVBranch, ByVal cll As Collection)
  cll.Add pv
  Set pv = pv.Get(pvtGetChild, 0)
  Do While pv.IsValid
    If pv.Data = Node_Table Then GatherTables pv, cll
    Set pv = pv.Get(pvtGetNextSibling, 0)
  Loop
End Sub

    
Private Sub RemoveAllTables(ByVal pv As PVBranch, ByVal RemoveBr As Boolean)
  Dim cll As New Collection
  GatherTables pv, cll
  
'  Dim sTmp As String, v
'  For Each v In cll
'    sTmp = sTmp & v.Text & vbCrLf
'  Next v
'  MsgBox "Remove:" & vbCrLf & sTmp, vbOKOnly, "H"
  
  Dim bRedr As Boolean
  bRedr = tv.Redraw
  tv.Redraw = False
  On Error GoTo ERR_H
  
  Dim l As Long
  For l = cll.Count To 1 Step -1
    With cll(l)
      If l <> 1 Or RemoveBr Then
        m_cf.Remove .DataVariant
        .DataVariant = Null
      End If
      If l <> 1 Then .Remove
    End With
  Next l
  
    
  tv.Redraw = bRedr
  Exit Sub
  
ERR_H:
  tv.Redraw = bRedr
  Err.Raise Err.Number
End Sub
    
Private Sub CmdTAssEmpty()
  Dim pvPar As PVBranch, pv As PVBranch
  Dim lIdx As Long
  
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
  If pv.Level = 0 Then
    MsgBox "Нельзя очистить корневую таблицу", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If pv.Data = Node_GOST_Ctx Then Set pv = pv.Get(pvtGetParent, 0)
  Set pvPar = pv.Get(pvtGetParent, 0)
  
  pv.Select pvtSelectNode
  
  lIdx = GetSelectedIndex(pvPar)
  If lIdx = -1 Then
    MsgBox "Внутренняя ошибка - слот не найден", vbCritical Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  Select Case pv.Data
    Case Node_Empty
      Exit Sub
    Case Node_Table
      RemoveAllTables pv, True
      'm_cf.Remove pv.DataVariant
  End Select
  
  pvPar.DataVariant.NAssSlotType lIdx, FT_None
  With pv
    .Clear
    .Text = pvPar.DataVariant.NComponentName(lIdx)
    Set .CustomItemPicture = imgEmpty.Picture
    .DataVariant = Null
    .Data = Node_Empty
  End With
  
End Sub

Private Sub CmdTAssGOST()
  Dim pvPar As PVBranch, pv As PVBranch
  Dim lIdx As Long
  
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
  If pv.Level = 0 Then
    MsgBox "Нельзя заменить корневую таблицу на норматив", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If pv.Data = Node_GOST_Ctx Then Set pv = pv.Get(pvtGetParent, 0)
  Set pvPar = pv.Get(pvtGetParent, 0)
  
  pv.Select pvtSelectNode
  
  lIdx = GetSelectedIndex(pvPar)
  If lIdx = -1 Then
    MsgBox "Внутренняя ошибка - слот не найден", vbCritical Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
    
  With frmGostEdit
    If pv.Data = Node_GOST Then
      Dim lKT As Long, lKG As Long, ift As IFactorTable
      Set ift = pvPar.DataVariant
      ift.NSlotGost lIdx, lKT, lKG
      .KeyTopic = lKT
      .KeyGOST = lKG
    Else
      Dim vE As Variant
      .KeyTopic = vE
      .KeyGOST = vE
    End If
    .SelectSelectedGost
    '.UpdateTVHint
    .ClearModalResult
    .Show vbModal, Me
    DoEvents
    If Not .ModalResult Then Exit Sub
    
    pv.Select pvtSelectNode
    CmdTAssEmpty
    
    pvPar.DataVariant.NAssSlotType lIdx, FT_Gost, .KeyTopic, .KeyGOST
    LoadGOST pv, pvPar.DataVariant, lIdx
    pv.Open pvtNode
    pv.EnsureVisible
  End With
End Sub

Private Sub CmdTRename()
  Dim pvPar As PVBranch, pv As PVBranch
  Dim lIdx As Long
  
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
  If pv.Level = 0 Then
    MsgBox "Нельзя переименовать корневую таблицу", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If pv.Data = Node_GOST_Ctx Then Exit Sub
    
  Set pvPar = pv.Get(pvtGetParent, 0)
    
  lIdx = GetSelectedIndex(pvPar)
  If lIdx = -1 Then
    MsgBox "Внутренняя ошибка - слот не найден", vbCritical Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
        
  pv.Select pvtSelectNode
  tv.BeginInPlaceEdit
  Dim r As RECT, x As Single, y As Single, w As Single, h As Single
  GetWindowRect tv.GetEditHWND(), r
  tv.EndInPlaceEdit
  
    
  Dim pt As POINTAPI
  pt.x = r.Left: pt.y = r.Top
  ScreenToClient hWnd, pt
  x = pt.x: y = pt.y
  
  pt.x = r.Right: pt.y = r.Bottom
  ScreenToClient hWnd, pt
  w = pt.x - x: h = pt.y - y
  
  x = ScaleX(x, vbPixels, vbTwips)
  w = ScaleX(w, vbPixels, vbTwips)
  y = ScaleY(y, vbPixels, vbTwips)
  h = ScaleY(h, vbPixels, vbTwips)
  
  w = tv.Width - 2 * x
      
  With txtInPlace
    .Move x, y, w, h
          
    .Text = pvPar.DataVariant.NComponentName(lIdx)
    Set m_pvEdit = pv
    m_lEditSlot = lIdx
    .Visible = True
    .ZOrder 0
    .SetFocus
  End With
  
  Exit Sub
ERR_H:
  Set m_pvEdit = Nothing
  txtInPlace.Visible = False
  
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  
End Sub

Private Sub txtInPlace_KeyPress(KeyAscii As Integer)
  Dim ift As IFactorTable, lKT As Long, lKG As Long, isci As IStCollItem
  
  If m_pvEdit Is Nothing Then Exit Sub
  
  On Error GoTo ERR_H
  
  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    GoTo L_UNLOCK
    
  ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Dim szText As String: szText = Trim(txtInPlace.Text)
    If Len(szText) < 1 Then
      MsgBox "Нельзя использовать пустое имя", vbExclamation Or vbOKOnly, "Ошибка"
      GoTo L_UNLOCK
    End If
    
      Dim pvPar As PVBranch
      Set pvPar = m_pvEdit.Get(pvtGetParent, 0)
      
      If m_pvEdit.Data = Node_Table Then
        Set isci = m_pvEdit.DataVariant
        m_cf.NameKey(isci.Key, m_pvEdit.DataVariant) = szText
      End If
            
      pvPar.DataVariant.NComponentName(m_lEditSlot) = szText
                    
      If m_pvEdit.Data = Node_GOST Then
        Set ift = pvPar.DataVariant
        ift.NSlotGost m_lEditSlot, lKT, lKG
        Set isci = frmGostEdit.DirectFGOST(lKT, lKG)
                
        If Not isci Is Nothing Then
           m_pvEdit.Text = ift.NComponentName(m_lEditSlot) & " (" & isci.Name & ")"
        Else
          m_pvEdit.Text = ift.NComponentName(m_lEditSlot)
        End If
        
      Else
        m_pvEdit.Text = szText
      End If
      
      GoTo L_UNLOCK
          
  End If
  
  Exit Sub
  
L_UNLOCK:
  On Error GoTo 0
  Set m_pvEdit = Nothing
  txtInPlace.Visible = False
  tv.SetFocus
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  GoTo L_UNLOCK
End Sub


Private Sub txtInPlace_LostFocus()

  If Not m_pvEdit Is Nothing Then
    Set m_pvEdit = Nothing
    txtInPlace.Visible = False
    tv.SetFocus
  End If
  
End Sub


Private Sub CmdTAssTable()
  Dim pvPar As PVBranch, pv As PVBranch
  Dim lIdx As Long
  
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
  If pv.Data = Node_Table Then
    RemoveAllTables pv, False
    pv.DataVariant.RemoveSlots
    pv.Clear
    Exit Sub
  End If
      
  If pv.Data = Node_GOST_Ctx Then Set pv = pv.Get(pvtGetParent, 0)
  Set pvPar = pv.Get(pvtGetParent, 0)
  
  pv.Select pvtSelectNode
  
  lIdx = GetSelectedIndex(pvPar)
  If lIdx = -1 Then
    MsgBox "Внутренняя ошибка - слот не найден", vbCritical Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
    
  Dim isci As IStCollItem, ift As IFactorTable
  Set ift = m_cf.Add(pvPar.DataVariant.NComponentName(lIdx))
  Set isci = ift
  
  pvPar.DataVariant.NAssSlotType lIdx, FT_Self, isci.Key
  With pv
    .Clear
    .Text = pvPar.DataVariant.NComponentName(lIdx)
    Set .CustomItemPicture = imgTable.Picture
    Set .DataVariant = ift
    .Data = Node_Table
    .ForeColor = RGB(14, 239, 61)
  End With
    
End Sub

Private Sub MkRefresh(ByVal pv As PVBranch)
  Dim l As Long: l = 0
  Set pv = pv.Get(pvtGetChild, 0)
  Do While pv.IsValid
  
    With pv
      If .Data = Node_GOST Then
        Dim bOpened As Boolean: bOpened = .IsOpen
        .Clear
        LoadGOST pv, .Get(pvtGetParent, 0).DataVariant, l
        If bOpened Then .Open pvtNode Else .Close pvtCloseNode
        
      ElseIf .Data = Node_Table Then
        MkRefresh pv
                
      End If
    End With
  
    Set pv = pv.Get(pvtGetNextSibling, 0)
    l = l + 1
  Loop
End Sub

Private Sub CmdTRefresh()
    
  On Error GoTo ERR_H
  
  tv.Redraw = False
  MkRefresh tv.Branches.Get(pvtGetChild, 0)
  tv.Redraw = True
  
  Exit Sub
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number
End Sub


Private Function GetSelectedIndex(ByVal pvPar As PVBranch) As Long
  
  Dim pvTmp As PVBranch: Set pvTmp = pvPar.Get(pvtGetChild, 0)
    
  GetSelectedIndex = 0
  Do While pvTmp.IsValid
    If pvTmp.IsSelected Then Exit Do
    Set pvTmp = pvTmp.Get(pvtGetNextSibling, 0)
    GetSelectedIndex = GetSelectedIndex + 1
  Loop
  
  If Not pvTmp.IsValid Then GetSelectedIndex = -1
End Function

Private Function GetNewSelOnDel(ByVal pv As PVBranch) As PVBranch
  Set GetNewSelOnDel = pv.Get(pvtGetNextSibling, 0)
  If Not GetNewSelOnDel.IsValid Then
    Set GetNewSelOnDel = pv.Get(pvtGetPrevSibling, 0)
    If Not GetNewSelOnDel.IsValid Then _
      Set GetNewSelOnDel = pv.Get(pvtGetParent, 0)
  End If
End Function

Private Sub CmdTRemove()
  Dim lIdx As Long
  Dim pv As PVBranch, pvPar As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
  If pv.Level = 0 Then
    MsgBox "Нельзя удалить корневую таблицу", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
    
  Dim pvNew As PVBranch
      
  Select Case pv.Data
    Case Node_GOST, Node_Empty
      Set pvPar = ParentTableBr(pv)
      If pvPar Is Nothing Then Exit Sub
      lIdx = GetSelectedIndex(pvPar)
      If lIdx = -1 Then
        MsgBox "Внутренняя ошибка - слот не найден", vbCritical Or vbOKOnly, "Ошибка"
        Exit Sub
      End If
      
      pvPar.DataVariant.RemoveSlots 1, lIdx
      Set pvNew = GetNewSelOnDel(pv)
      pv.Remove
      
    Case Node_Table
      'm_cf.Remove pv.DataVariant
      Set pvPar = pv.Get(pvtGetParent, 0)
      lIdx = GetSelectedIndex(pvPar)
      RemoveAllTables pv, True
      pvPar.DataVariant.RemoveSlots 1, lIdx
      Set pvNew = GetNewSelOnDel(pv)
      pv.Remove
      
    Case Node_GOST_Ctx
  End Select
  
  If Not pvNew Is Nothing Then _
    If pvNew.IsValid Then pvNew.Select pvtSelectNode
  
End Sub

Private Sub CmdTSave()
  On Error GoTo ERR_H
  m_bSavedOK = False
  
  Dim pv As PVBranch, sl As String
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  
  Do While pv.IsValid
    If pv.Data = Node_Table Then
      Dim isci As IStCollItem
      Set isci = pv.DataVariant
      If isci.Status <> OS_None Then
        Dim ift As IFactorTable
        Set ift = pv.DataVariant
        m_cf.Update ift
      End If
    End If
    Set pv = pv.Get(pvtGetNext, 0)
  Loop
  
  m_bSavedOK = True
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Function ParentTableBr(ByVal pv As PVBranch) As PVBranch
  
  If Not pv.IsValid Then
    Set ParentTableBr = Nothing
    Exit Function
  End If
  
  Select Case pv.Data
    Case Node_GOST, Node_Empty
      Set ParentTableBr = pv.Get(pvtGetParent, 0)
    Case Node_Table
      Set ParentTableBr = pv
    Case Node_GOST_Ctx
      Set ParentTableBr = pv.Get(pvtGetParent, 0).Get(pvtGetParent, 0)
  End Select
  
End Function


Private Sub CmdTCheckW()
  Dim pv As PVBranch, cll As New Collection
  Dim ift As IFactorTable, istci As IStCollItem
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  Do While pv.IsValid
    If pv.Data = Node_Table Then
      
      Set ift = pv.DataVariant
      If Not ift.IsWeightsCorrect Then
        Set istci = ift
        cll.Add istci.Name
      End If
      
    End If
    Set pv = pv.Get(pvtGetNext, 0)
  Loop
  
  If cll.Count > 0 Then
    Dim sTmp As String, l As Long, lC As Long
    lC = cll.Count
    For l = 1 To lC
      sTmp = sTmp & cll(l)
      If l <> lC Then sTmp = sTmp & ", "
    Next l
    MsgBox sTmp, vbInformation Or vbOKOnly, "Ошибочные веса у таблиц"
  Else
    MsgBox "Все веса расставлены верно", vbInformation Or vbOKOnly, "Сообщение"
  End If
  
End Sub

Private Sub CmdTAssWeights()
  If tdbgK.Visible Then
    tdbgK.Visible = False
    split1.Panes(1).Control = Nothing
    split1.Panes.Remove "Pane B"
    
    frmShadow.mnuTAssWeights.Checked = False
  Else
    split1.Visible = False
  
    split1.Panes(0).Control = Nothing
    split1.Panes.Add "Pane A", ssBottomOfSplit, "Pane B"
    split1.Panes(0).Control = tv
    split1.Panes(1).Control = tdbgK
    split1.Visible = True
    tdbgK.Visible = True
  
    frmShadow.mnuTAssWeights.Checked = True
        
    Set m_iftCurrentTable = Nothing
    tv_AfterSelChange Nothing
  End If
End Sub

Private Sub tdbgK_UnboundReadDataEx( _
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
    
  Set isci = m_iftCurrentTable
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
      
      'Dim lColIndex As Long, j As Long
      
      If lMaxColumnIdx = 0 Then
      'For j = 0 To lMaxColumnIdx
        'lColIndex = .ColumnIndex(l, j)
        .Value(l, 0) = m_iftCurrentTable.NWeight(lBookmark)
      'Next j
      End If
      
      .Bookmark(l) = lBookmark
      lRowsFetched = lRowsFetched + 1
    Next l
  End With
  
  RowBuf.RowCount = lRowsFetched
  If lStaRow >= 0 Then ApproximatePosition = lStaRow
End Sub

Private Sub tdbgK_UnboundWriteData( _
  ByVal RowBuf As TrueOleDBGrid60.RowBuffer, _
  WriteLocation As Variant)

  Dim Col As Integer
  On Error GoTo NoWrite
  
  With RowBuf
    If m_iftCurrentTable Is Nothing Or .ColumnCount < 1 Then
      .RowCount = 0
    Else
      If Not IsNull(.Value(0, 0)) Then _
        m_iftCurrentTable.NWeight(WriteLocation) = .Value(0, 0)
    End If
  End With
    
  Exit Sub
NoWrite:
    RowBuf.RowCount = 0
    m_s_MsgErr2 = Err.Description
    tdbgK.PostMsg 1
    'MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
    Err.Clear
End Sub

Private Function CurrentTable(Optional ByRef pvTbl As Variant) As IFactorTable
   
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then
    Set CurrentTable = Nothing
    If Not IsMissing(pvTbl) Then Set pvTbl = Nothing
  Else
    Select Case pv.Data
      Case Node_GOST, Node_Empty
        Set pv = pv.Get(pvtGetParent, 0)
      Case Node_Table
        If pv.Level <> 0 Then Set pv = pv.Get(pvtGetParent, 0)
      Case Node_GOST_Ctx
        Set pv = pv.Get(pvtGetParent, 0).Get(pvtGetParent, 0)
    End Select
        
    Set CurrentTable = pv.DataVariant
    If Not IsMissing(pvTbl) Then Set pvTbl = pv
    
  End If
End Function

Private Sub SelectSlot(ByVal pvT As PVBranch)

  If Not pvT Is Nothing Then
    'If pvT.Level <> 0 Then
      Dim lIdx As Long
      lIdx = GetSelectedIndex(pvT)
      If lIdx <> -1 Then tdbgK.Bookmark = lIdx
    'End If
  End If

End Sub

Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  With tdbgK
  
    On Error GoTo ERR_H
  
    If Not .Visible Then Exit Sub
     
    Dim vt As IFactorTable, pvT As PVBranch
    Set vt = CurrentTable(pvT)
    
        
    If vt Is m_iftCurrentTable Then
      SelectSlot pvT
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
    
    SelectSlot pvT
    
  End With
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Информация"
  Err.Clear
End Sub

Private Sub UpdateFooter()
  With tdbgK
    If m_iftCurrentTable Is Nothing Then
      .Columns(0).FooterText = ""
    Else
      Dim fTmp As Single, bFl As Boolean
      bFl = m_iftCurrentTable.IsWeightsCorrect(fTmp)
      .Columns(0).FooterText = FormatNumber(fTmp, 2)
    End If
  End With
End Sub

Private Sub tdbgK_LostFocus()
    
  On Error Resume Next
  'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
  With tdbgK
    If .CurrentCellModified Then
      .Update
      If Err.Number <> 0 Then
        .CurrentCellModified = False
        .EditActive = False
      End If
    End If
  End With
    
End Sub

Private Sub tdbgK_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 3662 Or DataError = 16389 Then
    Response = 0
    'tdbgCompl.PostMsg 1
    MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
    DoEvents
    m_s_MsgErr2 = ""
  End If
End Sub

Private Sub tdbgK_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = vbKeyReturn Then
    tdbgK.Update
  End If
End Sub


Private Sub tdbgK_Validate(Cancel As Boolean)
   On Error Resume Next
   tdbgK.Update
   If Err.Number <> 0 Then Cancel = True
End Sub


Private Sub tdbgK_PostEvent(ByVal MsgID As Integer)
  Select Case MsgID
    Case 1
      MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
      m_s_MsgErr2 = ""
  End Select
End Sub

Private Sub tdbgK_OwnerDrawCell( _
  ByVal hdc As Long, _
  ByVal Bookmark As Variant, _
  ByVal Split As Integer, _
  ByVal Col As Integer, _
  ByVal Left As Integer, _
  ByVal Top As Integer, _
  ByVal Right As Integer, _
  ByVal Bottom As Integer, Done As Integer)

    
  m_lLeftOfCell = Left
  
  If m_iftCurrentTable Is Nothing Then Exit Sub
  
  Dim lOldBkColor As Long, lOldTxtAln As Long, lOldTxtColor As Long, lOldBkMode As Long
  Dim r As RECT, lXLimit As Long, sTmp As String
  Dim sK As Single
  sK = m_iftCurrentTable.NWeight(Bookmark)
  lXLimit = Left + CSng(Right - Left) * sK
  
  lOldBkColor = GetBkColor(hdc)
  
  If sK > 0 Then
    With r
      .Left = Left: .Top = Top
      .Right = lXLimit: .Bottom = Bottom
    End With
    SetBkColor hdc, &H20FFFF
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, r, 0, 0, 0
  End If
  
  If sK < 1 Then
    With r
      .Left = lXLimit: .Top = Top
      .Right = Right: .Bottom = Bottom
    End With
    SetBkColor hdc, &HC0FFFF
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, r, 0, 0, 0
  End If
  
  With r
    .Left = Left: .Top = Top
    .Right = Right: .Bottom = Bottom
  End With
  
  sTmp = FormatNumber(sK, 2)
  'sTmp = CStr(m_iftCurrentTable.NWeight(Bookmark))
  'lOldTxtAln = SetTextAlign(hdc, TA_CENTER Or VTA_CENTER)
  lOldTxtColor = SetTextColor(hdc, &H0)
  lOldBkMode = SetBkMode(hdc, TRANSPARENT)
  'ExtTextOut hdc, Left, Top, 0, r, sTmp, Len(sTmp), 0
  DrawText hdc, sTmp, -1, r, DT_CENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_VCENTER
  
  If Bookmark = tdbgK.Bookmark Then _
    DrawFocusRect hdc, r

  
  SetBkMode hdc, lOldBkMode
  SetTextColor hdc, lOldTxtColor
  'SetTextAlign hdc, lOldTxtAln
  SetBkColor hdc, lOldBkColor
  Done = True
End Sub

Private Sub tdbgK_MouseMove( _
  Button As Integer, _
  Shift As Integer, _
  x As Single, y As Single)

  If Button <> 1 Or m_iftCurrentTable Is Nothing Then Exit Sub

  Dim lRow As Long, lCol As Long, lBookmark As Long
  Dim pt As POINTAPI, fTmp As Single
  
  With tdbgK
  
    If .CellContaining(x, y, lRow, lCol) Then
      lBookmark = lRow + .FirstRow
      If .Columns(0).Width <> 0 Then
        pt.x = ScaleX(x, vbTwips, vbPixels): pt.y = ScaleY(y, vbTwips, vbPixels)
        'ScreenToClient tdbgK.hWnd, pt
        fTmp = pt.x / (ScaleX(.Width, vbTwips, vbPixels) - 0)
        m_iftCurrentTable.NWeight(lBookmark) = IIf(fTmp <= 1, fTmp, 1)
        .RefetchRow lBookmark
        UpdateFooter
      End If
    End If
    
  End With

End Sub

Private Sub CmdTNormalize()
  
  Dim ift As IFactorTable
  Set ift = CurrentTable()
  If ift Is Nothing Then Exit Sub
  
  ift.NormalizeWeights
  
  With tdbgK
    If .Visible Then
      If .EditActive Then _
        .CurrentCellModified = False: .EditActive = False
      
      .RefetchCol 0
      UpdateFooter
    End If
  End With
  
End Sub

Private Sub CmdTUniform()

  Dim ift As IFactorTable
  Set ift = CurrentTable()
  If ift Is Nothing Then Exit Sub
  
  ift.SetWeights
  
  With tdbgK
    If .Visible Then
      If .EditActive Then _
        .CurrentCellModified = False: .EditActive = False
      
      .RefetchCol 0
      UpdateFooter
    End If
  End With
End Sub


