VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Begin VB.Form frmGOST 
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "frmGOST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   2910
      Top             =   3810
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgGOST 
      Height          =   2625
      Left            =   165
      OleObjectBlob   =   "frmGOST.frx":0442
      TabIndex        =   0
      Top             =   150
      Width           =   5640
   End
   Begin Threed.SSCommand ssOk 
      Height          =   525
      Left            =   150
      TabIndex        =   2
      ToolTipText     =   "Вносит изменения и закрывает диалог"
      Top             =   3195
      Width           =   3900
      _ExtentX        =   6879
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
      Picture         =   "frmGOST.frx":3C03
      Caption         =   "Подтвердить"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssCancel 
      Height          =   525
      Left            =   4305
      TabIndex        =   1
      ToolTipText     =   "Закрывает окно и отменяет модификации"
      Top             =   3210
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmGOST.frx":3D6D
      Caption         =   "Отменить"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmGOST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FirstStart As Integer

Private m_b_Res As Boolean
Private m_gt As IGostTable
Private m_b_LockResize As Boolean
Private m_dba_Slots As XArrayDB
Private m_s_MsgErr2 As String

Private m_lBookm As Variant
Private m_lRow As Long

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400

Private Enum PostCmds
  cmdUniformGrad
  cmdColCap
  cmdDelete
  cmdInsBefore
  cmdInsAfter
  cmdUniformGradRevers
End Enum


Private Sub Form_Initialize()
  m_b_LockResize = False
      
  
  With tdbgGOST
       
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
  
     With .Columns(0)
      .Alignment = dbgLeft
      .Locked = False
      .DefaultValue = "<Пусто>"
      .HeadAlignment = dbgCenter
      .WrapText = True
     End With
     
     With .Columns(1)
      .Alignment = dbgLeft
      .Locked = False
      .DefaultValue = 1
      .HeadAlignment = dbgCenter
      .Font.Bold = True
     End With
     
     With .Columns(2)
      .Alignment = dbgLeft
      .Locked = False
      .DefaultValue = "<Пусто>"
      .HeadAlignment = dbgCenter
      .Font.Bold = True
     End With
     
     With .Columns(3)
      .Alignment = dbgLeft
      .Locked = False
      .DefaultValue = Null
      .HeadAlignment = dbgCenter
      .WrapText = True
     End With
     
     .EvenRowStyle.BackColor = &H80FFFF
     .OddRowStyle.BackColor = &HC0FFFF
     .AlternatingRowStyle = True
     
     .RowHeight = 2 * .RowHeight
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
  
  tdbgGOST.Move L_PAD, T_PAD, ScaleWidth - 2 * L_PAD, ScaleHeight - 3 * T_PAD - ssOk.Height
  ssOk.Move tdbgGOST.Left, tdbgGOST.Top + tdbgGOST.Height + T_PAD, tdbgGOST.Width * (2 / 3)
  ssCancel.Move ssOk.Left + ssOk.Width + L_PAD, ssOk.Top
  ssCancel.Width = tdbgGOST.Left + tdbgGOST.Width - ssCancel.Left
  
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
    
  With GTMsgHook1
    .Windows.Item(1).Messages.Add _
      Key:="WM_PENDING_OP", Value:=WM_PENDING_OP, ProcessType:=gtBeforeDefaultProc, ProcessDefaultProc:=False
      
    .Windows.Item(1).Messages.Add _
      Key:="WM_RBUTTONDOWN", Value:=WM_RBUTTONDOWN, ProcessType:=gtBeforeDefaultProc
          
    .Windows.Item(1).SubClassHwnd = tdbgGOST.hWnd
    
    .Subclass = True
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssCancel_Click
  End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Set m_gt = Nothing
  GTMsgHook1.Subclass = False
End Sub

Private Sub GTMsgHook1_Message(ByVal hWnd As Long, ByVal MsgID As Long, ByVal wParam As Long, ByVal lParam As Long)
  With tdbgGOST
    If MsgID = WM_RBUTTONDOWN Then
    
      On Error GoTo ERR_H
    
      Dim bRes As Boolean
      Dim l As Long, lB As Variant
      Dim X_ As Single, Y_ As Single, pt As POINTAPI
              
      If .EditActive = True Or .DataChanged Then
        bRes = True
      Else
        pt.x = GTMsgHook1.LoWord(lParam)
        pt.y = GTMsgHook1.HiWord(lParam)
        'ScreenToClient tdbgGOST.hWnd, pt
        'X_ = pt.X: Y_ = pt.Y
        Y_ = ScaleY(pt.y, vbPixels, vbTwips)
        X_ = ScaleX(pt.x, vbPixels, vbTwips)
        
        l = .RowContaining(Y_)
        m_lRow = l
        
        If l = -1 Then
          bRes = False
          With frmShadow
            .mnuInsBefore.Enabled = False
            .mnuInsAfter.Enabled = False
            .mnuDelete.Enabled = False
            .mnuColCap.Enabled = True
            .mnuUniformGrad.Enabled = IIf(m_dba_Slots.Count(1) > 0, True, False)
          End With
        Else
          lB = .RowBookmark(l)
          m_lBookm = lB
          If IsEmpty(lB) Or IsNull(lB) Or (VarType(lB) = vbString) Then
            bRes = False
            With frmShadow
              .mnuInsBefore.Enabled = True
              .mnuInsAfter.Enabled = True
              .mnuDelete.Enabled = False
              .mnuColCap.Enabled = True
              .mnuUniformGrad.Enabled = False
            End With
          Else
            If lB = 0 And l <> 0 Then
              bRes = True
            Else
              bRes = False
              With frmShadow
                .mnuInsBefore.Enabled = True
                .mnuInsAfter.Enabled = True
                .mnuDelete.Enabled = IIf(m_dba_Slots.Count(1) > 0, True, False)
                .mnuColCap.Enabled = True
                .mnuUniformGrad.Enabled = IIf(m_dba_Slots.Count(1) > 0, True, False)
              End With
            End If
          End If
        End If
      End If
      
      If Not bRes Then _
        PopupMenu frmShadow.mnuTblEdit, vbPopupMenuLeftAlign, X_, Y_
    
      GTMsgHook1.Windows(1).Messages(2).ProcessDefaultProc = bRes
    ElseIf MsgID = WM_PENDING_OP Then
      On Error GoTo ERR_H
      
      Select Case lParam
        Case cmdUniformGrad
          CmUniformGrad
        Case cmdColCap
          CmColCap
        Case cmdDelete
          CmDelete
        Case cmdInsBefore
          CmInsBefore
        Case cmdInsAfter
          CmInsAfter
        Case cmdUniformGradRevers
          CmUniformGradRevers
      End Select
          
    End If
  End With
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Friend Sub PostUniformGrad()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdUniformGrad
End Sub

Friend Sub PostUniformGradRevers()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdUniformGradRevers
End Sub

Friend Sub PostInsAfter()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdInsAfter
End Sub

Friend Sub PostInsBefore()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdInsBefore
End Sub

Friend Sub PostDelete()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdDelete
End Sub

Friend Sub PostColCap()
  PostMessage tdbgGOST.hWnd, WM_PENDING_OP, 0, cmdColCap
End Sub


Private Sub ssCancel_Click()
  m_b_Res = False
  If tdbgGOST.EditActive Then _
    tdbgGOST.CurrentCellModified = False: tdbgGOST.EditActive = False
  Hide
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

Private Sub ssOk_Click()

  On Error Resume Next
  If tdbgGOST.EditActive Or tdbgGOST.CurrentCellModified Then _
    tdbgGOST.Update
    
  If Err.Number <> 0 Then Exit Sub
  On Error GoTo ERR_H
  
  m_gt.Check
  
  m_b_Res = True
  Hide
  Exit Sub
  
ERR_H:
  Dim sErrD As String: sErrD = Err.Description
  If m_gt.LastErrStrZeroIdx <> -1 Then
    On Error Resume Next
    tdbgGOST.Bookmark = m_gt.LastErrStrZeroIdx
    tdbgGOST.SelBookmarks.Add m_gt.LastErrStrZeroIdx
    DoEvents
  End If
  MsgBox sErrD, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub FillGrid()
    
  If m_dba_Slots Is Nothing Then
    Set m_dba_Slots = New XArrayDB
    
    With m_dba_Slots
      .ReDim 0, m_gt.NumberSlots - 1, 0, 3
    
      .DefaultColumnType(0) = XTYPE_STRING
      .DefaultColumnType(1) = XTYPE_SINGLE
      .DefaultColumnType(2) = XTYPE_STRING
      .DefaultColumnType(3) = XTYPE_STRING
    End With
      
    With tdbgGOST
      Set .Array = m_dba_Slots
    End With
  End If
  
  Dim l As Long, l2 As Long
  With m_gt
    l2 = .NumberSlots - 1
    For l = 0 To l2
      m_dba_Slots(l, 0) = .NDescr(l)
      m_dba_Slots(l, 1) = .NValue(l)
      m_dba_Slots(l, 2) = .NLingvoVal(l)
      m_dba_Slots(l, 3) = .NComment(l)
    Next l
  End With
  
  'tdbgGOST.Close False
  
  On Error Resume Next
  tdbgGOST.Bookmark = Null
  tdbgGOST.ReBind
  
End Sub


Private Sub tdbgGOST_BeforeDelete(Cancel As Integer)
  On Error GoTo ERR_H
  If IsNull(tdbgGOST.Bookmark) Then Exit Sub
  If tdbgGOST.SelBookmarks.Count <> 1 Then
    m_s_MsgErr2 = "Чтобы удалить, нужно выделить строку"
    tdbgGOST.PostMsg 1
    Exit Sub
  End If
  
  With m_gt
    .RemoveSlots 1, tdbgGOST.SelBookmarks(0)
  End With
  
  Exit Sub
ERR_H:
  Cancel = True
  m_s_MsgErr2 = Err.Description
  tdbgGOST.PostMsg 1
  Err.Clear
End Sub

Private Sub tdbgGOST_BeforeUpdate(Cancel As Integer)
  If IsNull(tdbgGOST.Columns(1).Value) Or IsEmpty(tdbgGOST.Columns(1).Value) Then _
    tdbgGOST.Columns(1).Value = 1
    
  If IsNull(tdbgGOST.Bookmark) Then _
    If tdbgGOST.AddNewMode <> dbgAddNewPending Then Exit Sub
  
  
  Dim lIdx As Long: lIdx = IIf(IsNull(tdbgGOST.Bookmark), 0, tdbgGOST.Bookmark)
  Dim bFlSlotUndo As Boolean
  bFlSlotUndo = False
  On Error GoTo ERR_H
    
  With m_gt
    If tdbgGOST.AddNewMode = dbgAddNewPending Then
      .InsertSlots 1
      bFlSlotUndo = True
      lIdx = .NumberSlots - 1
    End If
    
    With tdbgGOST
      If .Columns(0).DataChanged Then
        .Columns(0).Value = Trim(.Columns(0).Value)
        m_gt.NDescr(lIdx) = .Columns(0).Value
      End If
      
      If .Columns(1).DataChanged Then _
        m_gt.NValue(lIdx) = .Columns(1).Value
      
      If .Columns(2).DataChanged Then
        .Columns(2).Value = Trim(.Columns(2).Value)
        m_gt.NLingvoVal(lIdx) = .Columns(2).Value
      End If
      
      If .Columns(3).DataChanged Then
        .Columns(3).Value = Trim(.Columns(3).Value)
        m_gt.NComment(lIdx) = .Columns(3).Value
      End If
    End With
              
  End With
  Exit Sub
ERR_H:
  Cancel = True
  m_s_MsgErr2 = Err.Description
  tdbgGOST.PostMsg 1
  Err.Clear
  If bFlSlotUndo Then m_gt.RemoveSlots 1, m_gt.NumberSlots - 1
End Sub

Private Sub tdbgGOST_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 3662 Or DataError = 16389 Then
    Response = 0
    'tdbgCompl.PostMsg 1
    MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
    DoEvents
    m_s_MsgErr2 = ""
  End If
End Sub

Private Sub tdbgGOST_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = vbKeyReturn Then
    tdbgGOST.Update
  End If
End Sub


Private Sub tdbgGOST_Validate(Cancel As Boolean)
   On Error Resume Next
   tdbgGOST.Update
   If Err.Number <> 0 Then Cancel = True
End Sub

Private Sub tdbgGOST_LostFocus()
    
  On Error Resume Next
  'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
  With tdbgGOST
    If .CurrentCellModified Then
      .Update
      If Err.Number <> 0 Then
        .CurrentCellModified = False
        .EditActive = False
      End If
    End If
  End With
    
End Sub

Private Sub tdbgGOST_PostEvent(ByVal MsgID As Integer)
  Select Case MsgID
    Case 1
      MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Ошибка"
      m_s_MsgErr2 = ""
  End Select
End Sub


Friend Sub AssData(ByVal gt As IGostTable)
  FirstStart = 1
  m_b_Res = False
    
  tdbgGOST.Columns(0).Caption = gt.ClmName
  Set m_gt = gt
  
  FillGrid

End Sub

Friend Sub UnassData()
  tdbgGOST.Close True
  Set m_gt = Nothing
  Set m_dba_Slots = Nothing
End Sub

Private Sub CmUniformGradRevers()
  SetUniformGr True
End Sub

Private Sub CmUniformGrad()
  SetUniformGr False
End Sub

Private Sub SetUniformGr(ByVal Revers As Boolean)
  m_gt.SetUniformGrad Revers
  'FillGrid
  Dim l As Long, l2 As Long
  With m_gt
    l2 = .NumberSlots - 1
    For l = 0 To l2
      m_dba_Slots(l, 1) = .NValue(l)
    Next l
  End With
  tdbgGOST.RefetchCol 1
End Sub

Private Sub CmColCap()
  On Error GoTo ERR_H
  Dim pt As POINTAPI
  
  With frmColName
    .AssData
    .Text1.Text = m_gt.ClmName
    pt.x = ScaleX(tdbgGOST.Left, vbTwips, vbPixels)
    pt.y = ScaleY(tdbgGOST.Top, vbTwips, vbPixels)
    ClientToScreen tdbgGOST.Parent.hWnd, pt
    .Move ScaleX(pt.x + 50, vbPixels, vbTwips), ScaleY(pt.y + 20, vbPixels, vbTwips)
    .Show vbModal, Me
    If .ModalResult Then
      m_gt.ClmName = .Text1.Text
      tdbgGOST.Columns(0).Caption = .Text1.Text
    End If
    
    .UnassData
  End With
  
  Exit Sub
  
ERR_H:
  frmColName.UnassData
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub CmDelete()
  'm_lBookm
  'm_lRow
  tdbgGOST.Bookmark = m_lBookm
  tdbgGOST.SelBookmarks.Add m_lBookm
  DoEvents
  tdbgGOST.Delete
End Sub

Private Sub CmInsBefore()
  If IsEmpty(m_lBookm) Or IsNull(m_lBookm) Or (VarType(m_lBookm) = vbString) Then
    InserX 0
  Else
    InserX m_lBookm
  End If
End Sub

Private Sub CmInsAfter()
  If IsEmpty(m_lBookm) Or IsNull(m_lBookm) Or (VarType(m_lBookm) = vbString) Then
    InserX 0
  Else
    InserX m_lBookm + 1
  End If
End Sub

Private Sub InserX(ByVal lInto As Long, Optional ByVal lCnt As Long = 1)
  m_dba_Slots.InsertRows lInto, lCnt
  m_gt.InsertSlots lCnt, lInto
  FillGrid
End Sub
