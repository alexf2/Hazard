VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmScalePage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ClipControls    =   0   'False
   Icon            =   "frmScalePage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   3149
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      ClipControls    =   0   'False
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVNumericLib.PVNumeric PVNumeric1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   1
         EndProperty
         Height          =   345
         Left            =   1260
         TabIndex        =   1
         Top             =   345
         Width           =   1635
         _Version        =   393216
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         FreeFormEntry   =   -1  'True
         SpinButtons     =   0
         DecimalSeparator=   ","
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   525
         Left            =   2160
         TabIndex        =   3
         Top             =   990
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmScalePage.frx":000C
         Caption         =   "Подтверждение"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   525
         Left            =   135
         TabIndex        =   2
         Top             =   1020
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmScalePage.frx":0176
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmScalePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_Res As Boolean




Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Public Property Let ModalResult(ByVal bFl As Boolean)
  m_b_Res = bFl
End Property

Private Sub Form_Initialize()
  m_b_Res = False
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

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk.SetFocus
    ssOk_Click
  End If
End Sub


Private Sub PVNumeric1_Validate(Cancel As Boolean)
  If PVNumeric1.ValueReal > PVNumeric1.ValueMax Or PVNumeric1.ValueReal < PVNumeric1.ValueMin Then
    Cancel = True
    MsgBox "Ошибочное значение", vbExclamation Or vbOKOnly, "Ошибка"
  End If
End Sub

Private Sub ssCancel_Click()
  m_b_Res = False
  Hide
End Sub

Private Sub ssCancel_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, True, Me.hWnd
End Sub

Private Sub ssCancel_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, False, Me.hWnd
End Sub

Private Sub ssOk_Click()
  Dim bFl As Boolean
  PVNumeric1_Validate bFl
  If bFl Then Exit Sub
  
  m_b_Res = True
  Hide
End Sub


Private Sub ssOk_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, True, Me.hWnd
End Sub

Private Sub ssOk_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOk, False, Me.hWnd
End Sub

