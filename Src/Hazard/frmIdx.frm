VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmIdx 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Назначить индексы опасности"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmIdx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   3598
      _Version        =   131074
      BackColor       =   12632256
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVNumericLib.PVNumeric pvnI4 
         Height          =   345
         Left            =   3075
         TabIndex        =   4
         Top             =   630
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         FreeFormEntry   =   -1  'True
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -3.4E+38
         ValueMax        =   3.4E+38
         LimitValueByType=   1
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnI3 
         Height          =   345
         Left            =   2105
         TabIndex        =   3
         Top             =   630
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         FreeFormEntry   =   -1  'True
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -3.4E+38
         ValueMax        =   3.4E+38
         LimitValueByType=   1
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnI2 
         Height          =   345
         Left            =   1135
         TabIndex        =   2
         Top             =   630
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         FreeFormEntry   =   -1  'True
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -3.4E+38
         ValueMax        =   3.4E+38
         LimitValueByType=   1
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnI1 
         Height          =   345
         Left            =   165
         TabIndex        =   1
         Top             =   630
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         FreeFormEntry   =   -1  'True
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -3.4E+38
         ValueMax        =   3.4E+38
         LimitValueByType=   1
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin Threed.SSPanel sspI1 
         Height          =   210
         Left            =   165
         TabIndex        =   7
         Top             =   405
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Индекс 1"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSPanel sspI2 
         Height          =   210
         Left            =   1125
         TabIndex        =   8
         Top             =   405
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Индекс 2"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSPanel sspI3 
         Height          =   210
         Left            =   2085
         TabIndex        =   9
         Top             =   405
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Индекс 3"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSPanel sspI4 
         Height          =   210
         Left            =   3075
         TabIndex        =   10
         Top             =   405
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Индекс 4"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSCommand ssYes 
         Default         =   -1  'True
         Height          =   375
         Left            =   442
         TabIndex        =   5
         ToolTipText     =   "Применить новые индексы"
         Top             =   1350
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmIdx.frx":000C
         Caption         =   "Назначить"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssNo 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2452
         TabIndex        =   6
         ToolTipText     =   "Отказаться от модификации"
         Top             =   1350
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmIdx.frx":0176
         Caption         =   "Отменить"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmIdx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_b_Res As Boolean
Private m_f As Factor

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssNo_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssYes_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssNo_Click
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set m_f = Nothing
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
    
  Dim vArr, l As Long
  vArr = Array(pvnI1, pvnI2, pvnI3, pvnI4)
        
      
  For l = 0 To 3
    If vArr(l).Visible Then
      vArr(l).SetFocus
      If vArr(l).ValueReal < 0 Then
        MsgBox "Индекс должен быть >= 0", vbExclamation Or vbOKOnly, "Ошибка"
        Exit Sub
      End If
    End If
  Next l
  
  For l = 0 To 3
    If vArr(l).Visible Then m_f.Idx(l) = vArr(l).ValueReal
  Next l

  m_b_Res = True
  Hide
End Sub


Friend Sub AssData(ByVal f As Factor)
  
  m_b_Res = False
  
  Dim vArr, l As Long, l2 As Long
  vArr = Array(sspI1, pvnI1, sspI2, pvnI2, sspI3, pvnI3, sspI4, pvnI4)
 
  Dim sLeft As Single, sW As Single, bVis As Boolean
  sW = f.NIdx * pvnI1.Width + 100 * (f.NIdx - 1)
  sLeft = (ScaleWidth - sW) / 2#
 
  l2 = 2 * f.NIdx - 1
  For l = 0 To 7
    bVis = IIf(l <= l2, True, False)
    vArr(l).Visible = bVis
    If bVis Then
      vArr(l).Move sLeft
      If l Mod 2 <> 0 Then
        sLeft = sLeft + pvnI1.Width + 100
      End If
    End If
    If l Mod 2 <> 0 Then vArr(l).ValueReal = f.Idx(l \ 2)
  Next l

  Set m_f = f

End Sub

Friend Sub UnassData()
  Set m_f = Nothing
End Sub

