VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmEnumAss 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "frmEnumAss.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   4320
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   7620
      _Version        =   131074
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
         Height          =   2790
         Left            =   1267
         OleObjectBlob   =   "frmEnumAss.frx":000C
         TabIndex        =   2
         Top             =   120
         Width           =   4935
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFBF0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2790
         IntegralHeight  =   0   'False
         ItemData        =   "frmEnumAss.frx":2991
         Left            =   120
         List            =   "frmEnumAss.frx":2993
         TabIndex        =   1
         ToolTipText     =   "Енумераторы"
         Top             =   120
         Width           =   1005
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   525
         Left            =   1890
         TabIndex        =   5
         ToolTipText     =   "Выполняет модификации и закрывает диалог"
         Top             =   3210
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmEnumAss.frx":2995
         Caption         =   "Подтверждение"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4290
         TabIndex        =   6
         ToolTipText     =   "Отменяет любые модификации и закрывает диалог"
         Top             =   3210
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmEnumAss.frx":2AFF
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssRemov 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Удаляет текущий енумератор"
         Top             =   3060
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmEnumAss.frx":2C69
         Caption         =   "Удалить"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssAdd 
         Height          =   375
         Left            =   105
         TabIndex        =   4
         ToolTipText     =   "Добавляет новый енумератор"
         Top             =   3615
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmEnumAss.frx":2DC7
         Caption         =   "Добавить"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmEnumAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400
Private Const LST_WID = (1# / 5#)

Private m_f As Factor
Private m_gn As MGertNet
Private m_len_CollEnum As CollLingvo
Private m_dba_x As XArrayDB
Private m_b_Res As Boolean
Private m_l_CurrEnumID As Long

Private m_b_LockResize As Boolean

Private Sub Form_Initialize()
  m_b_LockResize = False
  m_l_CurrEnumID = -1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssCancel_Click
  End If
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
  
  Dim lBW As Long, lBH As Long
  lBW = ScaleX(2 * GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips)
  lBH = ScaleY(2 * GetSystemMetrics(SM_CYBORDER), vbPixels, vbTwips)
  
  SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
  
  List1.Move L_PAD, T_PAD, ScaleWidth * LST_WID, ScaleHeight - T_PAD * 4 - 2 * ssRemov.Height - lBH
  TDBGrid1.Move List1.Left + List1.Width + L_PAD, List1.Top, ScaleWidth - 3 * L_PAD - List1.Width + lBW, List1.Height
  ssRemov.Move L_PAD, List1.Top + List1.Height + T_PAD
  ssAdd.Move L_PAD, ssRemov.Top + ssRemov.Height + T_PAD
  
  ssOk.Move ssAdd.Left + ssAdd.Width + L_PAD2, List1.Top + List1.Height + (ScaleHeight - List1.Height - ssOk.Height) / 2, (ScaleWidth - ssAdd.Width - L_PAD * 3 - L_PAD2) / 2
  ssCancel.Move ssOk.Left + ssOk.Width + L_PAD, ssOk.Top, ScaleWidth - (ssOk.Left + ssOk.Width) - 2 * L_PAD
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnassData
End Sub

Private Sub List1_Click()
  FillGrid
End Sub

Private Sub ssAdd_Click()
  Dim lEnum As New LingvoEnum, l As Long
  m_len_CollEnum.Add lEnum
  For l = 0 To lEnum.Count - 1
    lEnum.Mark(l) = "<Пусто>"
  Next l
  List1.AddItem CStr(lEnum.ID)
  List1.ListIndex = List1.NewIndex
  FillGrid
End Sub

Private Sub ssAdd_LostFocus()
  HighLight2 ssAdd, False, Me.hWnd
End Sub

Private Sub ssCancel_Click()
  m_b_Res = False
  
  If TDBGrid1.EditActive Then _
    TDBGrid1.CurrentCellModified = False: TDBGrid1.EditActive = False
    
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

Private Sub ssOk_Click()
    
  Dim bm As New CBeam
  bm.Beam Me
  
    
  On Error Resume Next
  If TDBGrid1.EditActive Or TDBGrid1.CurrentCellModified Then _
    TDBGrid1.Update
  
  On Error GoTo 0
  
  SynchronizeData
  
  m_b_Res = True
  Hide
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


Friend Function ExistRefTo(ByVal lID As Long, ByRef sRef As String) As Boolean
  Dim f As Factor
  sRef = ""
  For Each f In m_gn.Factors
    If f.IDEnum = lID Then
      Dim ibs As IBSTRKey
      Set ibs = f
      sRef = sRef & ibs.Get() & ","
    End If
  Next f
  If Right(sRef, 1) = "," Then sRef = Left(sRef, Len(sRef) - 1)
  ExistRefTo = (sRef <> "")
End Function

Private Sub ssRemov_Click()
  If List1.ListIndex <> -1 Then
    Dim l As Long
    l = List1.ListIndex
    
    Dim sRefs As String
    If ExistRefTo(CLng(List1.List(l)), sRefs) Then
      MsgBox "На этот енумератор (" & List1.List(l) & ") имеются ссылки: '" & sRefs & "'", vbExclamation Or vbOKOnly, "Нельзя удалить"
      Exit Sub
    End If
    
    m_len_CollEnum.Remove CLng(List1.List(l))
    List1.RemoveItem l
    l = IIf(l > 0, l - 1, 0)
    If List1.ListCount > 0 Then
      If l >= List1.ListCount Then l = List1.ListCount - 1
      List1.ListIndex = l
    End If
    
    FillGrid
  End If
End Sub

Private Sub ssRemov_LostFocus()
  HighLight2 ssRemov, False, Me.hWnd
End Sub

Private Sub ssRemov_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemov, True, Me.hWnd
End Sub

Private Sub ssRemov_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemov, False, Me.hWnd
End Sub

Private Sub ssRemovAll_Click()

End Sub

Private Sub ssAdd_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAdd, True, Me.hWnd
End Sub

Private Sub ssAdd_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAdd, False, Me.hWnd
End Sub

Public Sub AssData(ByVal f As Factor, ByVal m As MGertNet)
    
  Set m_f = f
  Set m_gn = m
  
  m_b_Res = False
  Dim ibs As IBSTRKey
  Set ibs = f
  Caption = "Назначить енумератор для " & ibs.Get
  
  Dim iClon As IClonable
  Set iClon = m.Enumerators
  Set m_len_CollEnum = iClon.Clone
    
  
  Dim en As LingvoEnum, l As Long, lCnt As Long
  lCnt = 0
  List1.Clear
  For Each en In m_len_CollEnum
    List1.AddItem CStr(en.ID)
    If f.IDEnum = en.ID Then l = lCnt
    lCnt = lCnt + 1
  Next en
  
  List1.ListIndex = l
  FillGrid
End Sub

Private Sub FillGrid()
  'If TDBGrid1.EditActive Then

  TDBGrid1.Caption = "Енумератор " & List1.List(List1.ListIndex)
  
  Dim lenm As LingvoEnum
  For Each lenm In m_len_CollEnum
    Exit For
  Next lenm
  
  If m_dba_x Is Nothing Then
    Set m_dba_x = New XArrayDB
    
    m_dba_x.ReDim 0, lenm.Count - 1, 0, 2
    
    m_dba_x.DefaultColumnType(0) = XTYPE_DOUBLE
    m_dba_x.DefaultColumnType(1) = XTYPE_STRING
  
    'Set TDBGrid1.Array = m_dba_x
    With TDBGrid1
      .Columns(0).Alignment = dbgRight
      .Columns(1).Alignment = dbgLeft
      .Columns(0).Locked = True
      .Columns(1).Locked = False
      .Columns(1).DefaultValue = "<Пусто>"
      
      .Columns(0).HeadAlignment = dbgCenter
      .Columns(1).HeadAlignment = dbgCenter
      .CaptionStyle.ForeColor = RGB(0, 0, &HFF)
      .Columns(0).HeadFont.Bold = False
      .Columns(1).HeadFont.Bold = True
      .Columns(1).Font.Bold = True
      .Columns(0).ForeColor = RGB(&H80, &H80, &H80)
      
      .EvenRowStyle.BackColor = &H80FFFF
      .OddRowStyle.BackColor = &HC0FFFF
      .AlternatingRowStyle = True
       
      Set .Array = m_dba_x
    End With
  End If
  
  TDBGrid1.Close False
  
  Dim l As Long
  Dim lenm2 As LingvoEnum
  If List1.ListIndex = -1 Then
    m_dba_x.DeleteRows 0, m_dba_x.Count(1)
    TDBGrid1.ReBind
    m_l_CurrEnumID = -1
    Exit Sub
  End If
  
  If m_dba_x.Count(1) < 1 Then _
    m_dba_x.AppendRows lenm.Count
  
  m_l_CurrEnumID = CLng(List1.List(List1.ListIndex))
  Set lenm2 = m_len_CollEnum(m_l_CurrEnumID)
  For l = 0 To lenm2.Count - 1
    m_dba_x(l, 0) = lenm2.Quality(l)
    m_dba_x(l, 1) = lenm2.Mark(l)
  Next l
    
  TDBGrid1.ReBind
  
End Sub

Public Sub UnassData()
  List1.Clear
  
  TDBGrid1.Close True
  
  Set m_f = Nothing
  Set m_gn = Nothing
  Set m_len_CollEnum = Nothing
  
  'Set TDBGrid1.Array = Nothing
  Set m_dba_x = Nothing
End Sub

Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 1 Then
    'Dim lInd As Long
    'lInd = List1.ListIndex
    If m_l_CurrEnumID = -1 Then Exit Sub
    m_len_CollEnum(m_l_CurrEnumID).Mark(TDBGrid1.Bookmark) = TDBGrid1.Columns(1).Value
  End If
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If TDBGrid1.Columns(ColIndex).Value = "" Then _
    TDBGrid1.Columns(ColIndex).Value = "<Пусто>"
End Sub

Private Sub SynchronizeData()
  If m_f.IDEnum <> m_l_CurrEnumID And m_l_CurrEnumID <> -1 Then _
    m_f.IDEnum = m_l_CurrEnumID
    
  m_gn.Enumerators.UpdateFrom m_len_CollEnum
End Sub
