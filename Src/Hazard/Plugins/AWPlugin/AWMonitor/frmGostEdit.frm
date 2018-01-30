VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Begin VB.Form frmGostEdit 
   Caption         =   "Нормативы и ГОСТы"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "frmGostEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   4185
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   5640
      _Version        =   393216
      _ExtentX        =   9948
      _ExtentY        =   7382
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
   Begin VB.TextBox txtSelectedGOST 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "ГОСТ не выбран"
      Top             =   4395
      Width           =   5670
   End
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   1890
      Top             =   5415
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
      WindowText1     =   "frmGostEdit"
   End
   Begin VB.TextBox txtInPlace 
      BackColor       =   &H00FFFBF0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2550
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "Вызов редактора - F2"
      Top             =   5385
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgGOST 
      Height          =   315
      Left            =   960
      Picture         =   "frmGostEdit.frx":0442
      Top             =   5475
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgOFolder 
      Height          =   225
      Left            =   555
      Picture         =   "frmGostEdit.frx":056C
      Top             =   5445
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   135
      Picture         =   "frmGostEdit.frx":06A2
      Top             =   5445
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSCommand ssClose 
      Cancel          =   -1  'True
      Height          =   525
      Left            =   4305
      TabIndex        =   2
      ToolTipText     =   "Закрывает окно. Если выбор ГОСТа не подтверждён, то сохранится исходная связь слота."
      Top             =   4785
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmGostEdit.frx":07D8
      Caption         =   "Закрыть"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssOkGOST 
      Default         =   -1  'True
      Height          =   525
      Left            =   135
      TabIndex        =   1
      ToolTipText     =   "Назначает выбранный ГОСТ слоту таблицы"
      Top             =   4785
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      Enabled         =   0   'False
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
      Picture         =   "frmGostEdit.frx":0942
      Caption         =   "Подтвердить выбор ГОСТа"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmGostEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400

Private Enum PostCmds
  cmdAddGOST
  cmdAddTopic
  cmdCmEdit
  cmdCmRemove
  cmdCmRename
End Enum

Public FirstStart As Integer
Private m_b_Res As Boolean
Public KeyTopic As Variant
Public KeyGOST As Variant
Private m_b_LockResize As Boolean
Private m_collTopics As CollTopics

Private HighlightedNode As PVBranch

Private vEmpty As Variant
Private m_pvEdit As PVBranch
Private m_sPathToGostsStg As String

Private m_myFont As StdFont

Friend Sub ClearModalResult()
  m_b_Res = False
End Sub

Public Property Get ct() As CollTopics
  Set ct = m_collTopics
End Property

Public Property Get PathToGostsStg() As String
  PathToGostsStg = m_sPathToGostsStg
End Property

Friend Function DirectFGOST(ByVal lKT As Long, ByVal lKG As Long) As IGostTable
  Dim pv As PVBranch, gc As ICollGosts
  Set pv = tv.Branches.Get(pvtGetChild, 0)
  Set DirectFGOST = Nothing
  Do While pv.IsValid
    
    If pv.Data = lKT Then
      On Error Resume Next
      Set gc = pv.DataVariant
      If Err.Number <> 0 Then Set gc = Nothing
      Exit Do
    End If
    
    Set pv = pv.Get(pvtGetNextSibling, 0)
  Loop
  
  On Error Resume Next
  If gc Is Nothing Then Set gc = m_collTopics(lKT)
  Set DirectFGOST = gc(lKG)
End Function

Public Property Let PathToGostsStg(ByVal sNam As String)
  m_sPathToGostsStg = sNam
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  
  Dim iStg As IStorage
  Dim sTmp As String
  Set iStg = OpenCF(PathToGostsStg, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Set m_collTopics = New CollTopics
  Dim ips As IPersistStorage
  Dim ci As IStCollItem
  Set ci = m_collTopics: Set ips = m_collTopics
  ci.DefaultCU = False
  m_collTopics.SetUpdateMode Nothing, False
  
  If IsEmptyStg(iStg) Then
    ips.InitNew iStg
    ips.Save Nothing, 1
  Else
    ips.Load iStg
  End If
    
End Property


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
    If Width < 4000 Then Width = 4000
    If Height < 3000 Then Height = 3000
  Else
    m_b_LockResize = False
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  tv.Move L_PAD, T_PAD, ScaleWidth - 2 * L_PAD, ScaleHeight - 3 * T_PAD - ssOkGOST.Height - txtSelectedGOST.Height
  txtSelectedGOST.Move tv.Left, tv.Top + tv.Height + T_PAD / 2, tv.Width
  ssOkGOST.Move tv.Left, txtSelectedGOST.Top + txtSelectedGOST.Height + T_PAD / 2, tv.Width * (2 / 3)
  ssClose.Move ssOkGOST.Left + ssOkGOST.Width + L_PAD, ssOkGOST.Top
  ssClose.Width = tv.Width - ssOkGOST.Width - L_PAD
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssClose_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOkGOST_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  tv.StandardDefaultPicture = pvtNone
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    txtInPlace_LostFocus
    'Me.Hide
    ssClose_Click
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
    Case cmdAddGOST
      CmAddGOST
    Case cmdAddTopic
      CmAddTopic
    Case cmdCmEdit
      CmEdit
    Case cmdCmRemove
      CmRemove
    Case cmdCmRename
      CmRename
  End Select
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssClose_Click()
  'MClose
  'm_b_Res = False
  
  Dim ips As IPersistStorage
  Dim ic As IStCollItem
  Set ic = m_collTopics
  If ic.Status <> OS_None Then
    Set ips = m_collTopics
    ips.Save Nothing, 1
  End If
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

Private Sub ssOkGOST_LostFocus()
  HighLight2 ssOkGOST, False, Me.hWnd
End Sub

Private Sub ssOkGOST_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkGOST, True, Me.hWnd
End Sub

Private Sub ssOkGOST_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkGOST, False, Me.hWnd
End Sub

Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Private Sub ssOkGOST_Click()
  Dim pvb As PVBranch
  Set pvb = tv.Branches.Get(pvtGetNextSelected, 0)
  If pvb.Level = 1 Then
    m_b_Res = True
    KeyGOST = pvb.Data
    Set pvb = pvb.Get(pvtGetParent, 0)
    KeyTopic = pvb.Data
  ElseIf pvb.Level = 2 Then
    m_b_Res = True
    Set pvb = pvb.Get(pvtGetParent, 0)
    KeyGOST = pvb.Data
    Set pvb = pvb.Get(pvtGetParent, 0)
    KeyTopic = pvb.Data
  End If
  UpdateTVHint True
  'SelectSelectedGost
End Sub


Friend Sub AssData()
  FirstStart = 1
  m_b_Res = False
  
  LoadToTV
  
  ssOkGOST.Enabled = False
  Set HighlightedNode = Nothing
    
  SelectSelectedGost
  'UpdateTVHint
  
End Sub


Friend Sub UnassData()
  tv.Branches.Clear
  Set HighlightedNode = Nothing
  Set m_collTopics = Nothing
  Set m_pvEdit = Nothing
End Sub

Friend Function UpdateTVHint(Optional ByVal bInfo As Boolean = False) As PVBranch
  Set UpdateTVHint = Nothing

  Dim pvb As PVBranch
  Dim sGOST As String
  Set pvb = tv.Branches.Get(pvtGetNextSelected, 0)
  If pvb.IsValid Then
    If pvb.Level = 1 Then
      sGOST = pvb.Text
      Set pvb = pvb.Get(pvtGetParent, 0)
      'tv.ToolTipText = "Назначен ГОСТ: '" & pvb.Text & "." & sGOST & "'"
      txtSelectedGOST.Text = "Назначен норматив: '" & pvb.Text & "." & sGOST & "'"
    ElseIf pvb.Level = 2 Then
      Set pvb = pvb.Get(pvtGetParent, 0)
      sGOST = pvb.Text
      Set pvb = pvb.Get(pvtGetParent, 0)
      'tv.ToolTipText = "Назначен ГОСТ: '" & pvb.Text & "." & sGOST & "'"
      txtSelectedGOST.Text = "Назначен норматив: '" & pvb.Text & "." & sGOST & "'"
    End If
  Else
    'tv.ToolTipText = "ГОСТ не назначен"
    txtSelectedGOST.Text = "Норматив не назначен"
  End If
  If bInfo Then
    Set UpdateTVHint = pvb
    'MsgBox ssSelectedGOST.Caption, vbInformation Or vbOKOnly, "Информация"
  End If
End Function

Private Function ClearSelectedGT(Optional ByVal Direct As Boolean = False) As Boolean
  If Direct Or IsEmpty(KeyTopic) Or IsEmpty(KeyGOST) Then
    Dim vE As Variant
    KeyGOST = vE: KeyTopic = vE
    ClearSelectedGT = True
    m_b_Res = True
    txtSelectedGOST.Text = "Норматив не назначен"
    Exit Function
  End If
  ClearSelectedGT = False
End Function

Friend Sub SelectSelectedGost()
  
  If ClearSelectedGT() Then
    tv.Branches.Select pvtSelectUnselect
    Exit Sub
  End If
  
  Dim pvNode As PVBranch, bFlFndTopic As Boolean
  Set pvNode = tv.Branches.Get(pvtGetChild, 0)
  bFlFndTopic = False
  Do While pvNode.IsValid And Not bFlFndTopic
    If pvNode.Data = KeyTopic Then
      bFlFndTopic = True
      'pvNode.Select pvtSelectNode
    Else
      Set pvNode = pvNode.Get(pvtGetNextSibling, 0)
    End If
  Loop
  
  If bFlFndTopic Then
    bFlFndTopic = False
    pvNode.Open pvtNode
    pvNode.EnsureVisible
    tv_AfterExpand pvNode
    Set pvNode = pvNode.Get(pvtGetChild, 0)
    Do While pvNode.IsValid And Not bFlFndTopic
      If pvNode.Data = KeyGOST Then
        bFlFndTopic = True
        pvNode.Select pvtSelectNode
      Else
        Set pvNode = pvNode.Get(pvtGetNextSibling, 0)
      End If
    Loop
    
    If bFlFndTopic Then
      pvNode.Open pvtNode
      pvNode.EnsureVisible
      tv_AfterExpand pvNode
    End If
          
  End If
  
  If Not bFlFndTopic Then
    ClearSelectedGT True
    tv.Branches.Select pvtSelectUnselect
  Else
    UpdateTVHint
  End If
  
End Sub

Private Sub LoadToTV()
  With tv
    If .Count > 0 Then .Branches.Clear
    Dim pvNode As PVBranch
    
    On Error GoTo ERR_H
    .Redraw = False
    
    Dim lNumTopics As Long, l As Long
    lNumTopics = m_collTopics.Count
    For l = 1 To lNumTopics
      Set pvNode = .Branches.Add(pvtPositionInOrder, 0, m_collTopics.NameNth(l))
      pvNode.ForeColor = RGB(237, 232, 83)
      Set pvNode.CustomItemPicture = imgFolder.Picture
      pvNode.Data = m_collTopics.KeyNth(l)
      Set pvNode = pvNode.Add(pvtPositionInOrder, 0, "<?>")
      pvNode.ForeColor = RGB(237, 232, 83)
      'pvNode.DataVariant = 0
    Next l
    
  End With
  tv.Redraw = True
  Exit Sub
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number
End Sub

Private Sub tv_AfterCollapse(ByVal node As PVTreeView3Lib.PVIBranch)
  With node
    If .IsValid Then
      If .Level = 0 Then _
        Set .CustomItemPicture = imgFolder.Picture
    End If
  End With
End Sub

Private Sub LoadBranch(ByVal node As PVBranch)
  Dim bFlEmpty As Boolean: bFlEmpty = False
  Dim pvNodeToRemove As PVBranch
  
  On Error GoTo ERR_H
    
  If node.IsValid Then
    With node
      
      Dim pvNode As PVBranch
      Set pvNode = node.Get(pvtGetChild, 0)
      If pvNode.IsValid Then
      
        On Error Resume Next
        If IsEmpty(.DataVariant) Then _
          Set pvNodeToRemove = pvNode: bFlEmpty = True: Err.Clear
        If Err.Number <> 0 Then
          Set pvNodeToRemove = pvNode
          bFlEmpty = True
        End If
        On Error GoTo ERR_H
          
        If Not bFlEmpty Then Exit Sub
          
        If .Level = 0 Then
                    
          If bFlEmpty Then
            tv.Redraw = False
            
            Dim lNGosts As Long, l As Long
            Dim cllGosts As ICollGosts
            Set cllGosts = m_collTopics(.Data)
            Set .DataVariant = cllGosts
            lNGosts = cllGosts.Count
            For l = 1 To lNGosts
              Set pvNode = .Add(pvtPositionInOrder, 0, cllGosts.NameNth(l))
              pvNode.ForeColor = RGB(108, 245, 250)
              Set pvNode.CustomItemPicture = imgGOST.Picture
              pvNode.Data = cllGosts.KeyNth(l)
              pvNode.DataVariant = vEmpty
              Set pvNode = pvNode.Add(pvtPositionInOrder, 0, "<?>")
              pvNode.ForeColor = RGB(237, 232, 83)
            Next l
          End If
        ElseIf .Level = 1 Then
          If bFlEmpty Then
            tv.Redraw = False
            
            Dim lNSlots As Long
            Dim gTable As IGostTable
            Set pvNode = node.Get(pvtGetParent, 0)
            Set gTable = pvNode.DataVariant.Item(.Data)
            .DataVariant = 0
            lNSlots = gTable.NumberSlots - 1
            For l = 0 To lNSlots
              Set pvNode = .Add(pvtPositionInOrder, 0, gTable.NDescr(l) & " = " & gTable.NValue(l))
              pvNode.ForeColor = RGB(108, 245, 250)
            Next l
          End If
        End If
        
      End If
    End With
  End If
  
  If Not pvNodeToRemove Is Nothing Then pvNodeToRemove.Remove
  If Not tv.Redraw Then tv.Redraw = True
  Exit Sub
  
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number
End Sub

Private Sub tv_AfterExpand(ByVal node As PVTreeView3Lib.PVIBranch)
  On Error GoTo ERR_H
  LoadBranch node
  With node
    If .IsValid Then _
      If .Level = 0 Then Set .CustomItemPicture = imgOFolder.Picture
  End With
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  With node
    If .IsValid Then
      If .Level > 0 Then ssOkGOST.Enabled = True Else ssOkGOST.Enabled = False
    End If
  End With
End Sub


Private Sub tv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      
  If Not HighlightedNode Is Nothing Then _
    Set HighlightedNode.Font = Nothing
    
  With tv
    If Shift = 0 Then
      Set HighlightedNode = .HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
      If HighlightedNode.IsValid Then
          Set HighlightedNode.Font = m_myFont
      End If
    End If
  End With
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim pv As PVBranch
  
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    If pv.IsValid Then
      If pv.Level = 1 Or pv.Level = 2 Then
        ssOkGOST_Click
      End If
    End If
    
  ElseIf KeyCode = vbKeyDelete Then
    KeyCode = 0
    CmRemove
    
  ElseIf KeyCode = vbKeyF2 Then
    KeyCode = 0
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    If pv.IsValid Then
      If pv.Level = 0 Or pv.Level = 1 Then CmRename Else CmEdit
    End If
    
  ElseIf KeyCode = vbKeyF3 Then
    KeyCode = 0
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    If pv.IsValid Then
      If pv.Level <> 0 Then CmEdit
    End If
    
  End If
End Sub


Public Sub CmEdit()
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  
    
  On Error GoTo ERR_H
  If pv.Level = 1 Then
  ElseIf pv.Level = 2 Then
    Set pv = pv.Get(pvtGetParent, 0)
  Else
    Exit Sub
  End If
  
  Dim gt As IGostTable, pvPar As PVBranch
  Set pvPar = pv.Get(pvtGetParent, 0)
  Set gt = pvPar.DataVariant.Item(pv.Data)
  
  With frmGOST
    .Caption = pv.Text
    If .FirstStart = 0 Then CenterForm frmGOST, Me
    .AssData gt
    .Show vbModal, Me
    DoEvents
    .UnassData
    If .ModalResult Then
    
      Dim bm As CBeam: Set bm = New CBeam
      bm.Beam Me
      pvPar.DataVariant.Update gt
      bm.Off
      
      With pv
        .Clear
        .Close pvtCloseNode
        Dim vtEmpty As Variant
        .DataVariant = vtEmpty
        Set pvPar = .Add(pvtPositionInOrder, 0, "<?>")
        pvPar.ForeColor = RGB(237, 232, 83)
        tv_AfterExpand pv
        .Open pvtNode
        .Select pvtSelectNode
        .EnsureVisible
      End With
    End If
  End With
  
  Exit Sub
ERR_H:
  frmGOST.UnassData
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Function GetNtx(ByVal pv As PVBranch) As PVBranch
  Set GetNtx = pv.Get(pvtGetNextSibling, 0)
  If Not GetNtx.IsValid Then
    Set GetNtx = pv.Get(pvtGetPrevSibling, 0)
    If Not GetNtx.IsValid Then
      Set GetNtx = pv.Get(pvtGetParent, 0)
    End If
  End If
End Function

Public Sub CmRemove()
  Dim pv As PVBranch, pv2 As PVBranch
  Dim bFlRemCurr As Boolean
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  
  If Not pv.IsValid Then Exit Sub
  If pv.Level > 1 Then Exit Sub
      
  On Error GoTo ERR_H
  Set pv2 = GetNtx(pv)
  bFlRemCurr = False
  
  If pv.Level = 1 Then
    'UnLoadBrunch pv
    If pv.Data = KeyGOST Then bFlRemCurr = True
    pv.Get(pvtGetParent, 0).DataVariant.Remove pv.Data
    pv.Remove
  Else
    'UnLoadBrunch pv
    Dim vRes As VbMsgBoxResult
    vRes = MsgBox("Удалить раздел '" & pv.Text & "' ?", vbQuestion Or vbYesNo, "Запрос")
    If vRes = vbNo Then
      Set pv2 = Nothing
    Else
      If pv.Data = KeyTopic Then bFlRemCurr = True
      m_collTopics.Remove pv.Data
      pv.Remove
    End If
  End If
  
  If Not pv2 Is Nothing Then _
    If pv2.IsValid Then pv2.Select pvtSelectNode: pv2.EnsureVisible
    
  If bFlRemCurr Then ClearSelectedGT True
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Public Sub CmRename()
 
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  
  If Not pv.IsValid Then Exit Sub
  If pv.Level > 1 Then Exit Sub
  
  'LoadBranch pv
  
  Set m_pvEdit = pv
      
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
      
  KeyPreview = False
  ssClose.Cancel = False
  ssOkGOST.Default = False
    
  With txtInPlace
    .Move x, y, w, h
          
    .Text = pv.Text
    .Visible = True
    .ZOrder 0
    .SetFocus
  End With
  
  Exit Sub
ERR_H:
  Set m_pvEdit = Nothing
  txtInPlace.Visible = False
  KeyPreview = True
  ssClose.Cancel = True
  ssOkGOST.Default = True
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub UnloadBr(ByVal pv As PVBranch)
  On Error Resume Next
  tv.Redraw = False
  Dim vt As Variant
  With pv
    pv.Close pvtCloseNode
    pv.Clear
    pv.DataVariant = vt
    Set pv = pv.Add(pvtPositionInOrder, 0, "<?>")
    pv.ForeColor = RGB(237, 232, 83)
  End With
  tv.Redraw = True
End Sub

Private Function SGCollGosts(ByVal pv As PVBranch) As ICollGosts
  On Error Resume Next
  Set SGCollGosts = pv.DataVariant
  If Err.Number <> 0 Then Set SGCollGosts = Nothing
End Function


Private Sub txtInPlace_KeyPress(KeyAscii As Integer)
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
    
        
    Dim ic As IStCollItem
    If m_pvEdit.Level = 0 Then
      UnloadBr m_pvEdit
      m_collTopics.NameKey(m_pvEdit.Data, SGCollGosts(m_pvEdit)) = szText
      m_pvEdit.Text = szText
      
      GoTo L_UNLOCK
      
    ElseIf m_pvEdit.Level = 1 Then
      Dim pvTmp As PVBranch, icg As ICollGosts
      Set pvTmp = m_pvEdit.Get(pvtGetParent, 0)
      Set icg = pvTmp.DataVariant
      icg.NameKey(m_pvEdit.Data) = szText
      m_pvEdit.Text = szText
      
      GoTo L_UNLOCK
      
    End If
  End If
  
  Exit Sub
  
L_UNLOCK:
  On Error GoTo 0
  Me.KeyPreview = True
  ssClose.Cancel = True
  ssOkGOST.Default = True
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
    Me.KeyPreview = True
    ssClose.Cancel = True
    ssOkGOST.Default = True
    Set m_pvEdit = Nothing
    txtInPlace.Visible = False
    tv.SetFocus
  End If
End Sub


Public Sub CmAddTopic()
  Dim pvFocus As PVBranch, pv As PVBranch
  Dim cItem As IStCollItem
  
  On Error GoTo ERR_H
        
  Set cItem = m_collTopics.Add("<Новый>")
  
  Set pv = tv.Branches.Add(pvtPositionInOrder, 0, "<Новый>")
  Set pvFocus = pv
  pv.ForeColor = RGB(237, 232, 83)
  Set pv.CustomItemPicture = imgFolder.Picture
  pv.Data = cItem.Key
  Set pv = pv.Add(pvtPositionInOrder, 0, "<?>")
  pv.ForeColor = RGB(237, 232, 83)
        
  
  If Not pvFocus Is Nothing Then
    If pvFocus.IsValid Then _
      pvFocus.Select pvtSelectNode: pvFocus.EnsureVisible: CmRename
  End If
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Public Sub CmAddGOST()
  Dim pvFocus As PVBranch, pv As PVBranch
  Dim cItem As IStCollItem
  
  On Error GoTo ERR_H
        
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If Not pv.IsValid Then Exit Sub
  If pv.Level = 1 Then
    Set pv = pv.Get(pvtGetParent, 0)
  ElseIf pv.Level = 2 Then
    Set pv = pv.Get(pvtGetParent, 0).Get(pvtGetParent, 0)
  End If
  
  If Not pv.IsValid Then Exit Sub
  
  LoadBranch pv
        
  Set cItem = pv.DataVariant.Add("<Новый>")
  Set pv = pv.Add(pvtPositionInOrder, 0, "<Новый>")
  Set pvFocus = pv
  pv.ForeColor = RGB(108, 245, 250)
  Set pv.CustomItemPicture = imgGOST.Picture
  pv.Data = cItem.Key
  Set pv = pv.Add(pvtPositionInOrder, 0, "<?>")
  pv.ForeColor = RGB(237, 232, 83)
        
  
  If Not pvFocus Is Nothing Then
    If pvFocus.IsValid Then _
      pvFocus.Select pvtSelectNode: pvFocus.EnsureVisible: CmRename
  End If
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub tv_RButtonUp(ByVal node As PVTreeView3Lib.PVIBranch, ByVal X_ As Single, ByVal Y_ As Single)
  Dim pv As PVBranch
  Dim x As Single, y As Single
  x = ScaleX(X_, vbTwips, vbPixels): y = ScaleX(Y_, vbTwips, vbPixels)
  Set pv = tv.HitTest(x, y)
  
  With frmShadow
  
    If pv.IsValid Then
      pv.Select pvtSelectNode
      
      .mnuItem.Visible = True
      .s1.Visible = True
      
      If pv.Level = 0 Then
        .mnuItem.Caption = "Раздел"
        .mnuAddTopic = True
        .mnuAddGOST = True
        .mnuRename.Enabled = True
        .mnuEdit.Enabled = False
        .mnuRemove.Enabled = True
                
      ElseIf pv.Level = 1 Then
        .mnuItem.Caption = "Норматив"
        .mnuAddTopic = True
        .mnuAddGOST = True
        .mnuRename.Enabled = True
        .mnuEdit.Enabled = True
        .mnuRemove.Enabled = True
        
      Else
        .mnuItem.Caption = "Содержание норматива"
        .mnuAddTopic = True
        .mnuAddGOST = True
        .mnuRename.Enabled = False
        .mnuEdit.Enabled = True
        .mnuRemove.Enabled = False
        
      End If
      PopupMenu .mnuGOSTEdit, vbPopupMenuLeftAlign, X_, Y_, .mnuItem
      
    Else
      .mnuItem.Visible = False
      .s1.Visible = False
      .mnuAddTopic = True
      .mnuAddGOST = False
      .mnuRename.Enabled = False
      .mnuEdit.Enabled = False
      .mnuRemove.Enabled = False
      
      PopupMenu .mnuGOSTEdit, vbPopupMenuLeftAlign, X_, Y_
    End If
        
  End With
End Sub


Friend Sub PostAddGOST()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdAddGOST
End Sub

Friend Sub PostAddTopic()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdAddTopic
End Sub

Friend Sub PostCmEdit()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdCmEdit
End Sub

Friend Sub PostCmRemove()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdCmRemove
End Sub

Friend Sub PostCmRename()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdCmRename
End Sub

