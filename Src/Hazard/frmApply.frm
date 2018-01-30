VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmApply 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Применить меры"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmApply.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   5609
      _Version        =   131074
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFBF0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2085
         IntegralHeight  =   0   'False
         ItemData        =   "frmApply.frx":000C
         Left            =   247
         List            =   "frmApply.frx":000E
         TabIndex        =   1
         Top             =   345
         Width           =   4230
      End
      Begin Threed.SSPanel sspAllSP 
         Height          =   210
         Left            =   247
         TabIndex        =   4
         Top             =   105
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Выбранные меры"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSCommand ssNo 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2730
         TabIndex        =   3
         Top             =   2610
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmApply.frx":0010
         Caption         =   "Нет"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssYes 
         Default         =   -1  'True
         Height          =   375
         Left            =   1230
         TabIndex        =   2
         ToolTipText     =   "Вносит изменения в оценки факторов опасности модели"
         Top             =   2610
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmApply.frx":017A
         Caption         =   "Да"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private m_b_Res As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssNo_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssYes_Click
  End If
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssNo_Click
  End If
End Sub


Private Sub ssNo_Click()
  m_b_Res = False
  Hide
End Sub

Private Sub ssNo_LostFocus()
  HighLight2 ssNo, False, Me.hWnd
End Sub

Private Sub ssNo_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssNo, True, Me.hWnd
End Sub

Private Sub ssNo_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssNo, False, Me.hWnd
End Sub

Private Sub ssYes_LostFocus()
  HighLight2 ssYes, False, Me.hWnd
End Sub

Private Sub ssYes_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssYes, True, Me.hWnd
End Sub

Private Sub ssYes_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssYes, False, Me.hWnd
End Sub


Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Private Sub ssYes_Click()
  m_b_Res = True
  Hide
End Sub

Friend Sub AssData(ByVal coll As collSF)
  List1.Clear
  m_b_Res = False
  
'  Dim node As PVBranch
'
'  Set node = pvTree.Branches.Get(pvtGetChild, 0).Get(pvtGetChild, 0)
'  While node.IsValid
'      If node.CheckBoxState = pvtChkBoxStateChecked Then
'        List1.AddItem node.DataVariant
'      End If
'      Set node = node.Get(pvtGetNextSibling, 0)
'  Wend

  Dim sf As SafetyPrecaution
  For Each sf In coll
    List1.AddItem sf.Name
  Next sf
End Sub

Friend Sub UnassData()
  List1.Clear
End Sub
