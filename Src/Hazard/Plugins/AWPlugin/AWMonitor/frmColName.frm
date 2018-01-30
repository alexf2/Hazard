VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmColName 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmColName.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFBF0&
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   420
      Width           =   4410
   End
   Begin Threed.SSPanel sspI1 
      Height          =   210
      Left            =   150
      TabIndex        =   3
      Top             =   165
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   370
      _Version        =   131074
      Font3D          =   3
      ForeColor       =   0
      BackStyle       =   1
      Caption         =   "Название колонки"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   2
      Alignment       =   0
   End
   Begin Threed.SSCommand ssOk 
      Default         =   -1  'True
      Height          =   390
      Left            =   2730
      TabIndex        =   1
      Top             =   975
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   688
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmColName.frx":000C
      Caption         =   "Подтверждение"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssCancel 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   390
      Left            =   165
      TabIndex        =   2
      Top             =   975
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   688
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmColName.frx":0176
      Caption         =   "Отмена"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmColName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_Res As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssCancel_Click
  End If
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
  Text1.Text = Trim(Text1.Text)
  If Len(Text1.Text) < 1 Then
    MsgBox "Недопустимо пустое имя", vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If

  m_b_Res = True
  Hide
End Sub


Friend Sub AssData()
  m_b_Res = False
End Sub

Friend Sub UnassData()
  
End Sub


