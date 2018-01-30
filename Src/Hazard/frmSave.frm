VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmSave 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Сохранить"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmSave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   4948
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      ClipControls    =   0   'False
      BorderWidth     =   0
      BevelOuter      =   0
      Begin Threed.SSFrame SSFrame1 
         Height          =   1290
         Left            =   840
         TabIndex        =   6
         Top             =   180
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   2275
         _Version        =   131074
         BackStyle       =   1
         Caption         =   "Данные"
         ClipControls    =   0   'False
         Begin Threed.SSCheck ssSP 
            Height          =   255
            Left            =   315
            TabIndex        =   2
            Top             =   810
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   450
            _Version        =   131074
            BackStyle       =   1
            Caption         =   "Комплексы мер"
            Value           =   1
         End
         Begin Threed.SSCheck ssModel 
            Height          =   255
            Left            =   315
            TabIndex        =   1
            Top             =   360
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   450
            _Version        =   131074
            BackStyle       =   1
            Caption         =   "Модель"
            Value           =   1
         End
      End
      Begin Threed.SSCommand ssNo 
         Height          =   420
         Left            =   2775
         TabIndex        =   4
         ToolTipText     =   "Ничего не сохранять"
         Top             =   1665
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   741
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmSave.frx":000C
         Caption         =   "Не сохранять"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   5
         ToolTipText     =   "Отменить операцию"
         Top             =   2205
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   741
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmSave.frx":0176
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   420
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "Сохранить указанные данные"
         Top             =   1665
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   741
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmSave.frx":02E0
         Caption         =   "Сохранить"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_l_Res As Long



Public Property Let SaveModel(ByVal bFl As Boolean)
  'ssModel.Value = IIf(bFl, ssCBChecked, ssCBUnchecked)
  ssModel.Value = bFl
End Property
Public Property Get SaveModel() As Boolean
  'SaveModel = IIf(ssModel.Value = ssCBChecked, True, False)
  SaveModel = ssModel.Value
End Property

Public Property Let SaveSP(ByVal bFl As Boolean)
  'ssSP.Value = IIf(bFl, ssCBChecked, ssCBUnchecked)
  ssSP.Value = bFl
End Property
Public Property Get SaveSP() As Boolean
  'SaveSP = IIf(ssSP.Value = ssCBChecked, True, False)
  SaveSP = ssSP.Value
End Property


Public Property Get ModalResult() As Long
  ModalResult = m_l_Res
End Property

Public Property Let ModalResult(ByVal lRes As Long)
  m_l_Res = lRes
End Property

Private Sub Form_Initialize()
  m_l_Res = vbCancel
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    Me.Hide
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk.SetFocus
    ssOk_Click
  End If
End Sub

Private Sub ssCancel_Click()
  m_l_Res = vbCancel
  MClose
End Sub

Private Sub ssCancel_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, True, Me.hWnd
End Sub

Private Sub ssCancel_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, False, Me.hWnd
End Sub

Private Sub ssNo_Click()
  m_l_Res = vbNo
  MClose
End Sub

Private Sub ssNo_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssNo, True, Me.hWnd
End Sub

Private Sub ssNo_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssNo, False, Me.hWnd
End Sub

Private Sub ssOk_Click()
  m_l_Res = vbOK
  MClose
End Sub

Friend Sub MClose()
  PostMessage hWnd, WM_CLOSE, 0, 0
End Sub


Private Sub ssOk_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, True, Me.hWnd
End Sub

Private Sub ssOk_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, False, Me.hWnd
End Sub

