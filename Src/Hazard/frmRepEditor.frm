VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmRepEditor 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   ClipControls    =   0   'False
   Icon            =   "frmRepEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5953
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFBF0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         IntegralHeight  =   0   'False
         ItemData        =   "frmRepEditor.frx":000C
         Left            =   90
         List            =   "frmRepEditor.frx":000E
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   120
         Width           =   3555
      End
      Begin Threed.SSCommand ssRemovAll 
         Height          =   525
         Left            =   3855
         TabIndex        =   5
         ToolTipText     =   "Удаляет все отчёты"
         Top             =   1245
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepEditor.frx":0010
         Caption         =   "Удалить всё"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssRemov 
         Height          =   525
         Left            =   3855
         TabIndex        =   4
         ToolTipText     =   "Удаляет выделенные отчёты"
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepEditor.frx":0186
         Caption         =   "Удалить"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   615
         TabIndex        =   3
         ToolTipText     =   "Отменяет любые модификации и закрывает диалог"
         Top             =   2805
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepEditor.frx":02E4
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   525
         Left            =   3150
         TabIndex        =   2
         ToolTipText     =   "Выполняет модификации и закрывает диалог"
         Top             =   2805
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepEditor.frx":044E
         Caption         =   "Подтверждение"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmRepEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private m_vs As VSPrinter
Private m_c As CollectionX
Private m_cRemove As CollectionX
Private m_b_LockResize As Boolean
Private m_b_Res As Boolean

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400

Public Sub Assign(ByVal vs As VSPrinter, ByVal coll As CollectionX)
  Set m_vs = vs
  Set m_c = coll
  m_b_Res = False
  Set m_cRemove = New CollectionX
  With List1
    .Clear
    
    Dim iRpt As IRepItem, l As Long
    l = 1
    For Each iRpt In m_c
      .AddItem CStr(l) & ". " & iRpt.Title
      .ItemData(.NewIndex) = l
      l = l + 1
    Next iRpt
    
    If .ListCount > 0 Then .Selected(0) = True
  End With
End Sub

Public Sub Unassign()
  List1.Clear
  Set m_vs = Nothing
  Set m_c = Nothing
  Set m_cRemove = Nothing
End Sub


Private Sub Form_Initialize()
  m_b_LockResize = False
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
  If UnloadMode = vbFormControlMenu Then _
    Cancel = 1: ssCancel_Click
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
  List1.Move L_PAD, T_PAD, ScaleWidth - ssRemov.Width - 3 * L_PAD, ScaleHeight - ssRemov.Height - 3 * T_PAD
  ssRemov.Move lBW + List1.Left + List1.Width + L_PAD, List1.Top + (List1.Height - 2 * ssRemov.Height - T_PAD2) / 2
  ssRemovAll.Move ssRemov.Left, ssRemov.Top + ssRemov.Height + T_PAD2
  ssCancel.Move (ScaleWidth - 2 * ssCancel.Width - L_PAD2) / 2, List1.Top + List1.Height + T_PAD + (ScaleHeight - List1.Height - ssCancel.Height - 2 * T_PAD) / 2 - lBH
  ssOk.Move ssCancel.Left + ssCancel.Width + L_PAD2, ssCancel.Top
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_vs = Nothing
  Set m_c = Nothing
  Set m_cRemove = Nothing
End Sub



Private Sub SSCommand2_Click()

End Sub


Private Sub ssCancel_Click()
  m_b_Res = False
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
  
  Dim l
  SortCollX m_cRemove
  For Each l In m_cRemove
    m_c.Remove CLng(l)
  Next l
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

Private Sub ssRemov_Click()
  Dim l As Long
  Dim cSel As New Collection
  For l = List1.ListCount - 1 To 0 Step -1
    If List1.Selected(l) Then _
      m_cRemove.Add List1.ItemData(l): cSel.Add l
  Next l
  Dim lMin As Long, lv
  lMin = 65536
  For Each lv In cSel
      List1.RemoveItem CLng(lv)
      If lMin > lv Then lMin = lv
  Next lv
  lMin = lMin - 1
  If lMin < 0 Then lMin = 0
  If List1.ListCount > 0 Then List1.Selected(lMin) = True
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
  Dim l As Long
  For l = List1.ListCount - 1 To 0 Step -1
    m_cRemove.Add List1.ItemData(l)
  Next l
  List1.Clear
End Sub

Private Sub ssRemovAll_LostFocus()
  HighLight2 ssRemovAll, False, Me.hWnd
End Sub

Private Sub ssRemovAll_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemovAll, True, Me.hWnd
End Sub

Private Sub ssRemovAll_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemovAll, False, Me.hWnd
End Sub

Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

