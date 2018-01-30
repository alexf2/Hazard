VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmComplSP 
   BorderStyle     =   0  'None
   Caption         =   "Редактор комплексов мер"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   2400
      Left            =   1665
      TabIndex        =   0
      Top             =   3375
      Width           =   3690
      _Version        =   393216
      _ExtentX        =   6509
      _ExtentY        =   4233
      _StockProps     =   237
      ForeColor       =   0
      BackColor       =   9450828
      Appearance      =   1
      StandardDefaultPicture=   10
      AllowInPlaceEditing=   0   'False
      BackColor       =   9450828
      ForeColor       =   0
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
   Begin TrueOleDBGrid60.TDBGrid tdbgCompl 
      Height          =   2790
      Left            =   150
      OleObjectBlob   =   "frmComplSP.frx":0000
      TabIndex        =   1
      Top             =   45
      Width           =   3345
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgSP 
      Height          =   2790
      Left            =   3630
      OleObjectBlob   =   "frmComplSP.frx":24BE
      TabIndex        =   2
      Top             =   30
      Width           =   3345
   End
   Begin VB.Image imgFChange 
      Height          =   255
      Left            =   270
      Picture         =   "frmComplSP.frx":5095
      Top             =   5415
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgExcludeClosed 
      Height          =   240
      Left            =   1170
      Picture         =   "frmComplSP.frx":51E3
      Top             =   5490
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgExclude 
      Height          =   240
      Left            =   735
      Picture         =   "frmComplSP.frx":52E5
      Top             =   5295
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMoney 
      Height          =   270
      Left            =   315
      Picture         =   "frmComplSP.frx":53E7
      Top             =   4890
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgC 
      Height          =   270
      Left            =   1005
      Picture         =   "frmComplSP.frx":5541
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgT 
      Height          =   270
      Left            =   480
      Picture         =   "frmComplSP.frx":569B
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   1005
      Picture         =   "frmComplSP.frx":57F5
      Top             =   3975
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgH 
      Height          =   270
      Left            =   555
      Picture         =   "frmComplSP.frx":594F
      Top             =   4065
      Visible         =   0   'False
      Width           =   270
   End
   Begin Threed.SSCommand ssRemov 
      Height          =   375
      Left            =   15
      TabIndex        =   5
      ToolTipText     =   "Удаляет выделенный комплекс мер"
      Top             =   2850
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      Enabled         =   0   'False
      PictureUseMask  =   -1  'True
      Picture         =   "frmComplSP.frx":5AA9
      Caption         =   "Удалить комплекс"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssEditSP 
      Height          =   375
      Left            =   4875
      TabIndex        =   4
      ToolTipText     =   "Редактировать первую выбранную меру"
      Top             =   2850
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      Enabled         =   0   'False
      PictureUseMask  =   -1  'True
      Picture         =   "frmComplSP.frx":5C07
      Caption         =   "Редактировать"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssRemovSP 
      Height          =   375
      Left            =   3630
      TabIndex        =   3
      ToolTipText     =   "Удаляет выделенные в сетке меры"
      Top             =   2865
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      Enabled         =   0   'False
      PictureUseMask  =   -1  'True
      Picture         =   "frmComplSP.frx":5D19
      Caption         =   "Удалить"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin VB.Image imgSF 
      Height          =   270
      Left            =   990
      Picture         =   "frmComplSP.frx":5E77
      Top             =   4935
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmComplSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum PVTNodes
  PVT_None
  PVT_Money
  PVT_Excl
  PVT_FChanges
  PVT_M
  PVT_H
  PVT_C
  PVT_T
End Enum

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6


Private Const m_LPad = 100
Private Const m_UPad = 70
Private Const m_L2Pad = 200

Private m_csf_Coll As CollectionX

Private m_dba_Compl As XArrayDB
Private m_dba_SP As XArrayDB
Private m_s_MsgErr As String
Private m_s_MsgErr2 As String
Private m_csf_SF As collSF
Private m_sf_SafeP As SafetyPrecaution

Private HighlightedNode As PVBranch


Public Property Get MinW() As Single
  MinW = 7000
End Property

Public Property Get MinH() As Single
  MinH = 4500
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
  End If
End Function



Private Sub Form_DblClick()
  DumpXX
End Sub

Private Sub Form_Initialize()
  'Set glUtil = New GlUtils
End Sub

Private Sub Form_Load()
   With tdbgCompl
     .Columns(0).Alignment = dbgLeft
     .Columns(0).Locked = False
     .Columns(0).DefaultValue = "<Пусто>"
     .Columns(0).HeadAlignment = dbgCenter
     .Columns(0).Font.Bold = True
     .EvenRowStyle.BackColor = &H80FFFF
     .OddRowStyle.BackColor = &HC0FFFF
     .AlternatingRowStyle = True
   End With
   
   With tdbgSP
     .Columns(0).Alignment = dbgLeft
     .Columns(0).Locked = False
     .Columns(0).DefaultValue = "<Пусто>"
     .Columns(0).HeadAlignment = dbgCenter
     .Columns(1).Visible = False
     .Columns(1).Locked = True
     .EvenRowStyle.BackColor = &H80FFFF
     .OddRowStyle.BackColor = &HC0FFFF
     .AlternatingRowStyle = True
   End With
End Sub

Private Sub Form_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxSPMonitor.GetExtraWH sWExtr, sHExtr
  
  tdbgCompl.Move m_LPad, m_UPad, (ScaleWidth - 3 * m_LPad - sWExtr) / 2
  tdbgSP.Move tdbgCompl.Left + tdbgCompl.Width + m_LPad, m_UPad, tdbgCompl.Width, tdbgCompl.Height
  
  ssRemov.Move tdbgCompl.Left, tdbgCompl.Top + tdbgCompl.Height + m_UPad / 2
  ssRemovSP.Move tdbgSP.Left, ssRemov.Top
  ssEditSP.Move ssRemovSP.Left + ssRemovSP.Width + m_LPad / 2, ssRemov.Top
  
  tv.Move tdbgCompl.Left, ssRemov.Top + ssRemov.Height + m_UPad / 2, ScaleWidth - sWExtr - 2 * m_LPad, ScaleHeight - sHExtr - tdbgCompl.Height - 3 * m_UPad - ssRemov.Height
End Sub

Private Sub Form_Terminate()
  'Set glUtil = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  HandsOffModel
  Unload frmEditSP
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_csf_Coll Is Nothing)
End Property

Public Sub HandsOffModel()
  If tdbgSP.EditActive Then _
    tdbgSP.CurrentCellModified = False: tdbgSP.EditActive = False
  If tdbgCompl.EditActive Then _
    tdbgCompl.CurrentCellModified = False: tdbgCompl.EditActive = False
    
  tdbgSP.Close
  tdbgCompl.Close
  If Not m_dba_SP Is Nothing Then _
    m_dba_SP.Clear
  If Not m_dba_Compl Is Nothing Then _
    m_dba_Compl.Clear
  
  Set m_dba_SP = Nothing
  Set m_dba_Compl = Nothing
  
  tdbgCompl.Enabled = False
  tdbgSP.Enabled = False
  ssRemov.Enabled = False
  
  Set m_csf_SF = Nothing
  Set m_sf_SafeP = Nothing
  Set m_csf_Coll = Nothing
  Set HighlightedNode = Nothing
End Sub

Public Sub BindModel(ByVal coll As CollectionX)
  Set m_csf_Coll = coll
  
  Set m_dba_Compl = New XArrayDB
  m_dba_Compl.ReDim 0, coll.Count - 1, 0, 1
  m_dba_Compl.DefaultColumnType(0) = XTYPE_STRING
  m_dba_Compl.DefaultColumnType(1) = XTYPE_LONG
    
  Dim l As Long
  For l = 1 To coll.Count
    m_dba_Compl(l - 1, 0) = coll.key(l)
    m_dba_Compl(l - 1, 1) = l
  Next l
  Set tdbgCompl.Array = m_dba_Compl
  
  tdbgCompl.Enabled = True
    
  tdbgCompl.ReBind
  ssRemov.Enabled = IIf(coll.Count > 0, True, False)
End Sub

Private Sub SSCommand1_Click()

End Sub


Private Sub ssEditSP_Click()
  If m_sf_SafeP Is Nothing Then
    MsgBox "Нет выбранной меры", vbExclamation Or vbOKOnly, "Нельзя редактировать"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  frmEditSP.AssData haApp.GertNetMain, m_sf_SafeP, m_csf_SF
  frmEditSP.Show vbModal, frmMain
  DoEvents
  
  If frmEditSP.ModalResult Then
    tv.Redraw = False
    TvClearGl tv, tv.Branches.Get(pvtGetChild, 0).Text
    LoadTreeGl tv, m_sf_SafeP
    ExpandAll tv: tv.Redraw = True
    
    frmSP.LastSP = -1
  End If
  
  frmEditSP.UnassData
  
  Exit Sub
ERR_H:
  If tv.Redraw = False Then ExpandAll tv: tv.Redraw = True
  frmEditSP.UnassData
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ssEditSP_LostFocus()
  HighLight ssEditSP, False
End Sub

Private Sub ssEditSP_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssEditSP, True
End Sub

Private Sub ssEditSP_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssEditSP, False
End Sub

Private Sub ssRemov_Click()
  If Not IsBoundModel Then
    MsgBox "Нет модели", vbExclamation Or vbOKOnly, "Нельзя выполнить"
    Exit Sub
  End If
  If IsNull(tdbgCompl.Bookmark) Then
    MsgBox "Нет комплексов мер. Выберите комплекс в сетке", vbExclamation Or vbOKOnly, "Нельзя выполнить"
    Exit Sub
  End If
  
  tdbgCompl.Delete
  'm_dba_Compl.DeleteRows
  
End Sub

Private Sub DumpXX()
  If Not m_csf_Coll Is Nothing Then
    Dim l As Long, s As String
    For l = 1 To m_csf_Coll.Count
      s = s & CStr(l) & " - [" & m_csf_Coll.key(l) & "]" & vbCrLf
    Next
    MsgBox s, vbOKOnly, "Комплексы"
  End If
  
  If Not m_csf_SF Is Nothing Then
    s = ""
    Dim sfp As SafetyPrecaution
    For Each sfp In m_csf_SF
      Dim key As ILongKey: Set key = sfp
      s = s & CStr(key.Get()) & " - [" & sfp.Name & "]" & vbCrLf
    Next
    MsgBox s, vbOKOnly, "Меры"
  End If
End Sub

Private Sub ssRemov_LostFocus()
  HighLight ssRemov, False
End Sub

Private Sub ssRemov_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssRemov, True
End Sub

Private Sub ssRemov_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssRemov, False
End Sub

Private Sub ssRemovSP_Click()
  If tdbgSP.SelBookmarks.Count < 1 Then
    MsgBox "Чтобы удалить меры, нужно выделить в сетке соответствующие строки", vbInformation Or vbOKOnly, "Сообщение"
    Exit Sub
  End If
  
  If m_csf_SF Is Nothing Then
    MsgBox "Нет текущего набора мер", vbInformation Or vbOKOnly, "Сообщение"
    Exit Sub
  End If
  
  Dim vBm, collTmp As New CollectionX
  For Each vBm In tdbgSP.SelBookmarks
    'm_dba_SP.DeleteRows vBm
    collTmp.Add vBm
  Next
  tdbgSP.Close False
  SortCollX collTmp
  For Each vBm In collTmp
    m_csf_SF.Remove m_dba_SP(vBm, 1)
    m_dba_SP.DeleteRows vBm
  Next
  tdbgSP.ReBind
  
  frmSP.LastSP = -1
End Sub


Private Sub ssRemovSP_LostFocus()
  HighLight ssRemovSP, False
End Sub

Private Sub ssRemovSP_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssRemovSP, True
End Sub

Private Sub ssRemovSP_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssRemovSP, False
End Sub

Private Sub tdbgCompl_AfterInsert()
  'If tdbgCompl.AddNewMode = dbgAddNewPending Then
    Dim ips As IPersistStorage
    Dim csf As New collSF
    Set ips = csf: ips.InitNew Nothing
    
    If m_csf_Coll.Count < 1 Then
      m_csf_Coll.Add csf, m_dba_Compl(m_dba_Compl.Count(1) - 1, 0)
    Else
      m_csf_Coll.Add csf, m_dba_Compl(m_dba_Compl.Count(1) - 1, 0), , m_csf_Coll.Count
    End If
    
    'Dim dd
    'dd = m_dba_Compl(m_dba_Compl.Count(1) - 1, 0)
    m_dba_Compl(m_dba_Compl.Count(1) - 1, 1) = m_csf_Coll.Count
    'm_dba_Compl(m_dba_Compl.Count(1) - 1, 0) = tdbgCompl.Columns(0).Value
  'End If
  
    ssRemov.Enabled = True
    haApp.DirtySFSF = True
End Sub


Private Sub tdbgCompl_BeforeDelete(Cancel As Integer)
'  Dim vBk, cc As Long
'  For Each vBk In tdbgCompl.SelBookmarks
'    cc = vBk
'    m_csf_Coll.Remove CLng(m_dba_Compl(cc, 1))
'  Next vBk
'  Dim cc, ii
'  cc = m_dba_Compl.Count(1)
'  ii = 1
'Dim v1, v2
'v1 = m_dba_Compl(tdbgCompl.Bookmark, 0)
'v2 = m_dba_Compl(tdbgCompl.Bookmark, 1)
  m_csf_Coll.Remove m_dba_Compl(tdbgCompl.Bookmark, 0)
  ssRemov.Enabled = IIf(m_csf_Coll.Count > 0, True, False)
  Set m_csf_SF = Nothing
  haApp.DirtySFSF = True
End Sub

Private Sub tdbgCompl_BeforeUpdate(Cancel As Integer)
  If tdbgCompl.Columns(0).Value = "" Then _
    tdbgCompl.Columns(0).Value = "<Пусто>"
    
  m_csf_Coll.AutoDefaultGetKey = True
  m_csf_Coll.DefaultItem = Nothing
  Dim collSF As GERTNETLib.collSF
  Set collSF = m_csf_Coll.Item(tdbgCompl.Columns(0).Value)
  m_csf_Coll.AutoDefaultGetKey = False
  If Not collSF Is Nothing Then
    Cancel = True
    m_s_MsgErr = "Такой комплекс мер (" & tdbgCompl.Columns(0).Value & ") уже есть"
    tdbgCompl.PostMsg 1
    Exit Sub
  End If
  
  If IsNull(tdbgCompl.Bookmark) Or tdbgCompl.AddNewMode = dbgAddNewPending Then Exit Sub
  
  m_csf_Coll.key(tdbgCompl.Bookmark + 1) = tdbgCompl.Columns(0).Value
  haApp.DirtySFSF = True
  tdbgSP.Caption = "Меры (" & tdbgCompl.Columns(0).Value & ")"
    
End Sub

Private Sub tdbgCompl_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 3662 Or DataError = 16389 Then
    Response = 0
    MsgBox m_s_MsgErr, vbExclamation Or vbOKOnly, "Нельзя назначить имя"
    m_s_MsgErr = ""
  End If
End Sub



Private Sub tdbgCompl_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = vbKeyReturn Then
    'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
      tdbgCompl.Update
  End If
End Sub

Private Sub tdbgCompl_LostFocus()
  On Error Resume Next
  'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
   tdbgCompl.Update
End Sub

Private Sub tdbgCompl_PostEvent(ByVal MsgId As Integer)
  Select Case MsgId
    Case 1
      MsgBox m_s_MsgErr, vbExclamation Or vbOKOnly, "Нельзя назначить имя"
      m_s_MsgErr = ""
  End Select
End Sub

Private Sub tdbgCompl_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If m_dba_Compl Is Nothing Then
    haApp.CurrSFColl = -1
    ssRemov.Enabled = False
    tdbgSP.Caption = "Меры"
    UpdateSPGrid True
    Exit Sub
  End If
  If tdbgCompl.AddNewMode <> dbgNoAddNew Or IsNull(tdbgCompl.Bookmark) Then
    If IsNull(tdbgCompl.Bookmark) Then _
      haApp.CurrSFColl = -1
      
    ssRemov.Enabled = False
    tdbgSP.Caption = "Меры"
    UpdateSPGrid True
  Else
    haApp.CurrSFColl = tdbgCompl.Bookmark + 1
    tdbgSP.Caption = "Меры (" & tdbgCompl.Columns(0).Value & ")"
    UpdateSPGrid False
  End If
  
  If tdbgCompl.AddNewMode = dbgNoAddNew Then _
    ssRemov.Enabled = True
    
  'frmSP.LastSP = -1
  
End Sub

Private Sub UpdateSPGrid(ByVal bOff As Boolean)
  If bOff Then
    tdbgSP.Close
    Set m_csf_SF = Nothing
    tdbgSP.Enabled = False
    ssRemovSP.Enabled = False
    ssEditSP.Enabled = False
    
    If m_dba_SP Is Nothing Then Exit Sub
    
    If m_dba_SP.Count(1) > 0 Then _
      m_dba_SP.DeleteRows 0, m_dba_SP.Count(1)
  Else
    Dim csf As collSF
    Set csf = m_csf_Coll(tdbgCompl.Bookmark + 1)
    Set m_csf_SF = csf
    
    If m_dba_SP Is Nothing Then _
      Set m_dba_SP = New XArrayDB: Set tdbgSP.Array = m_dba_SP
      
    m_dba_SP.ReDim 0, csf.Count - 1, 0, 1
    m_dba_SP.DefaultColumnType(0) = XTYPE_STRING
    m_dba_SP.DefaultColumnType(1) = XTYPE_LONG
      
    Dim sp As SafetyPrecaution, l As Long
    l = 0
    For Each sp In csf
      Dim ilk As ILongKey
      Set ilk = sp
      m_dba_SP(l, 0) = sp.Name
      m_dba_SP(l, 1) = ilk.Get()
      l = l + 1
    Next sp
    
    tdbgSP.ReBind
    tdbgSP.Enabled = True
    If csf.Count > 0 Then
      ssRemovSP.Enabled = True
      ssEditSP.Enabled = True
    Else
      ssRemovSP.Enabled = False
      ssEditSP.Enabled = False
    End If
  End If
End Sub

Private Sub tdbgSP_AfterInsert()
  If m_csf_SF Is Nothing Then
    MsgBox "Чтобы добавить меру нужно выбрать комплекс", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  Dim spSaf As New SafetyPrecaution
      
  m_csf_SF.Add spSaf
  Dim ilKey As ILongKey: Set ilKey = spSaf
  spSaf.Name = m_dba_SP(m_dba_SP.Count(1) - 1, 0)
  m_dba_SP(m_dba_SP.Count(1) - 1, 1) = ilKey.Get()
  tdbgSP.RefetchRow tdbgSP.RowBookmark(tdbgSP.VisibleRows - 1)
      
  ssRemovSP.Enabled = True
  ssEditSP.Enabled = True
  
  frmSP.LastSP = -1
End Sub


Private Sub tdbgSP_BeforeUpdate(Cancel As Integer)
  If tdbgSP.Columns(0).Value = "" Then _
    tdbgSP.Columns(0).Value = "<Пусто>"
    
  Dim spSaf As SafetyPrecaution, sName As String
  sName = UCase(tdbgSP.Columns(0).Value)
  For Each spSaf In m_csf_SF
    If UCase(spSaf.Name) = sName Then
      sName = tdbgSP.Columns(0).Value
      Cancel = True
      m_s_MsgErr2 = "Такая  мера (" & sName & ") уже есть"
      tdbgSP.PostMsg 1
      Exit Sub
    End If
  Next
      
  If IsNull(tdbgSP.Bookmark) Or tdbgSP.AddNewMode = dbgAddNewPending Then Exit Sub
  
  m_csf_SF.Item(CLng(m_dba_SP(tdbgSP.Bookmark, 1))).Name = tdbgSP.Columns(0).Value
  
  frmSP.LastSP = -1
End Sub

Private Sub tdbgSP_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 3662 Or DataError = 16389 Then
    Response = 0
    MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Нельзя назначить имя"
    m_s_MsgErr2 = ""
  End If
End Sub

Private Sub tdbgSP_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = vbKeyReturn Then
    'If tdbgSP.EditActive Or tdbgSP.CurrentCellModified Then
      tdbgSP.Update
  End If
End Sub

Private Sub tdbgSP_LostFocus()
  On Error Resume Next
  'If tdbgSP.EditActive Or tdbgSP.CurrentCellModified Then
   tdbgSP.Update
End Sub

Private Sub tdbgSP_PostEvent(ByVal MsgId As Integer)
  Select Case MsgId
    Case 1
      MsgBox m_s_MsgErr2, vbExclamation Or vbOKOnly, "Нельзя назначить имя"
      m_s_MsgErr2 = ""
  End Select
End Sub

Private Sub tdbgSP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If m_dba_SP Is Nothing Then
    ssRemovSP.Enabled = False
    ssEditSP.Enabled = False
    UpdateTree True, "Мера"
    Exit Sub
  End If
  If m_dba_SP.Count(1) < 1 Then
    ssRemovSP.Enabled = False
    ssEditSP.Enabled = False
    UpdateTree True, "Мера"
    Exit Sub
  End If
  
  If tdbgSP.AddNewMode <> dbgNoAddNew Or IsNull(tdbgSP.Bookmark) Then
    ssRemovSP.Enabled = False
    ssEditSP.Enabled = False
    UpdateTree True, "Мера"
  Else
    UpdateTree False, "Мера '" & tdbgSP.Columns(0).Value & "'"
  End If
  
  If tdbgSP.AddNewMode = dbgNoAddNew Then _
    ssRemovSP.Enabled = True: ssEditSP.Enabled = True
End Sub

Private Sub UpdateTree(ByVal bFlOff As Boolean, ByVal sCapt As String)
  On Error GoTo ERR_H
'Exit Sub
  Set HighlightedNode = Nothing
  If bFlOff Then
    tv.Redraw = False
    TvClearGl tv
    
    
    Set m_sf_SafeP = Nothing
  Else
'Exit Sub
    Set m_sf_SafeP = m_csf_SF(m_dba_SP(tdbgSP.Bookmark, 1))
    
    tv.Redraw = False
    TvClearGl tv, sCapt
    LoadTreeGl tv, m_sf_SafeP
    
  End If
  ExpandAll tv
  tv.Redraw = True
  Exit Sub
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub tv_AfterCollapse(ByVal node As PVTreeView3Lib.PVIBranch)
  If node.Data = PVT_Excl Then _
    Set node.CustomItemPicture = imgExcludeClosed.Picture
End Sub

Private Sub tv_AfterExpand(ByVal node As PVTreeView3Lib.PVIBranch)
  If node.Data = PVT_Excl Then _
    Set node.CustomItemPicture = imgExclude.Picture
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim myFont As StdFont
    
    If Not HighlightedNode Is Nothing Then
        Set HighlightedNode.Font = Nothing
    End If
    
    If Shift = 0 Then
        Set HighlightedNode = tv.HitTest(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
        If HighlightedNode.IsValid Then
            Set myFont = tv.CreateFont
            myFont.Bold = True
            Set HighlightedNode.Font = myFont
        End If
    End If
End Sub


Friend Function FindNodeXXGl(ByVal pv As PVTreeViewX, ByVal pvtTyp As PVTNodes) As PVBranch
  Dim pvNode As PVBranch
  
  Set pvNode = pv.Branches.Get(pvtGetChild, 0).Get(pvtGetChild, 0)
  While pvNode.IsValid
      If pvNode.Data = pvtTyp Then
        Set FindNodeXXGl = pvNode
        Exit Function
      End If
      Set pvNode = pvNode.Get(pvtGetNextSibling, 0)
  Wend
End Function

Friend Sub CreBranchGl(ByVal pvParent As PVBranch, ByRef pvNode As PVBranch, ByVal sName As String, ByVal pic As StdPicture)
  Set pvNode = pvParent.Add(pvtPositionInOrder, 0, sName)
  pvNode.ForeColor = RGB(14, 239, 61)
  Set pvNode.CustomItemPicture = pic
End Sub

Friend Sub LoadTreeGl(ByVal pv As PVTreeViewX, ByVal sf As SafetyPrecaution)

  Set HighlightedNode = Nothing

  Dim pvNodeMain As PVBranch, pvNodeM As PVBranch, pvNodeH As PVBranch, pvNodeT As PVBranch, pvNodeC As PVBranch
  Set pvNodeMain = pv.Branches.Get(pvtGetChild, 0)
  Set pvNodeM = FindNodeXXGl(tv, PVT_Money)
  pvNodeM.Text = "Стоимость = " & FormatCurrency(sf.Cost, 2, vbTrue, vbFalse, vbTrue)
  Set pvNodeM = Nothing

  Dim pvNodeFCh As PVBranch
  Set pvNodeFCh = FindNodeXXGl(tv, PVT_FChanges)
  
  Dim vvt

  For Each vvt In sf.FChanges
    Dim pvB As PVBranch
    Dim fc As FChange
    Set fc = vvt
    
    Select Case UCase(Left$(fc.NameFactor, 1))
      Case "M"
        If pvNodeM Is Nothing Then CreBranchGl pvNodeFCh, pvNodeM, "Машина", imgM.Picture
        pvNodeM.Data = PVT_M
        Set pvB = pvNodeM
      Case "H"
        If pvNodeH Is Nothing Then CreBranchGl pvNodeFCh, pvNodeH, "Человек", imgH.Picture
        pvNodeH.Data = PVT_H
        Set pvB = pvNodeH
      Case "C"
        If pvNodeC Is Nothing Then CreBranchGl pvNodeFCh, pvNodeC, "Среда", imgC.Picture
        pvNodeC.Data = PVT_C
        Set pvB = pvNodeC
      Case "T"
        If pvNodeT Is Nothing Then CreBranchGl pvNodeFCh, pvNodeT, "Технология", imgT.Picture
        pvNodeT.Data = PVT_T
        Set pvB = pvNodeT
    End Select


   Dim sFName As String, sTxt As String
   sFName = haApp.GertNetMain.Factors(fc.NameFactor).Name
   sTxt = fc.NameFactor & "  (" & TSign(fc.Delta) & CStr(Abs(fc.Delta)) & ")  '" & sFName & "'"
   Set pvB = pvB.Add(pvtPositionInOrder, 0, sTxt)
   pvB.ForeColor = RGB(108, 245, 250)
   pvB.Data = PVT_None

  Next
    
  Set pvNodeFCh = FindNodeXXGl(tv, PVT_Excl)
  Dim l As Long, larrComp() As Long
  larrComp = sf.NonCompatibles
  For l = LBound(larrComp) To UBound(larrComp)
    Dim csf As SafetyPrecaution
    Set csf = m_csf_SF(larrComp(l))
    sTxt = csf.Name
    Set pvB = pvNodeFCh.Add(pvtPositionInOrder, 0, sTxt)
    pvB.ForeColor = RGB(108, 245, 250)
    pvB.Data = PVT_None
  Next l
End Sub

Friend Sub TvClearGl(ByVal pv As PVTreeViewX, Optional ByVal sName As String = "Мера")
  
  Set HighlightedNode = Nothing
  
  If pv.Count < 1 Then
    Dim pvNodeMain As PVBranch, pvNodeMoney As PVBranch, pvNodeExcl As PVBranch
    Set pvNodeMain = pv.Branches.Add(pvtPositionInOrder, 0, sName)
    pvNodeMain.ForeColor = RGB(237, 232, 83)
    pvNodeMain.StandardItemPicture = pvtpicFolders
    Set pvNodeMain.CustomItemPicture = imgSF.Picture
    pvNodeMain.Data = PVT_None
    
    Set pvNodeMoney = pvNodeMain.Add(pvtPositionInOrder, 0, "Стоимость = ")
    pvNodeMoney.ForeColor = RGB(14, 239, 61)
    pvNodeMoney.Data = PVT_Money
    Set pvNodeMoney.CustomItemPicture = imgMoney.Picture
    
    Set pvNodeExcl = pvNodeMain.Add(pvtPositionInOrder, 0, "Несовместность")
    pvNodeExcl.ForeColor = RGB(14, 239, 61)
    pvNodeExcl.Data = PVT_Excl
    Set pvNodeExcl.CustomItemPicture = imgExcludeClosed.Picture
    
    Set pvNodeExcl = pvNodeMain.Add(pvtPositionInOrder, 0, "Воздействие на факторы опасности")
    pvNodeExcl.ForeColor = RGB(14, 239, 61)
    pvNodeExcl.Data = PVT_FChanges
    Set pvNodeExcl.CustomItemPicture = imgFChange.Picture
          
    ExpandAll tv
  Else
  
    pv.Branches.Get(pvtGetChild, 0).Text = sName
    Set pvNodeMain = pv.Branches.Get(pvtGetChild, 0).Get(pvtGetChild, 0)
    Do While pvNodeMain.IsValid
      Set pvNodeMoney = pvNodeMain
      Set pvNodeMain = pvNodeMain.Get(pvtGetNextSibling, 0)
      If pvNodeMoney.Data <> PVT_Money Then
        If pvNodeMoney.Data = PVT_FChanges Or pvNodeMoney.Data = PVT_Excl Then
          pvNodeMoney.Clear
        Else
          pvNodeMoney.Remove
        End If
      ElseIf pvNodeMoney.Data = PVT_Money Then
        pvNodeMoney.Text = "Стоимость = "
      End If
    Loop
  End If
  
End Sub



