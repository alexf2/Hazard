VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmAlhoOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки алгоритмов"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmAlhoOpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VsOcxLib.VideoSoftIndexTab vsTabOpt 
      Align           =   1  'Align Top
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _Version        =   327680
      _ExtentX        =   8731
      _ExtentY        =   4128
      _StockProps     =   102
      Caption         =   "Аналитический устойчивый|Ан. уст. приближенный"
      ConvInfo        =   1418783674
      ForeColor       =   0
      Style           =   3
      AutoSwitch      =   -1  'True
      ShowFocusRect   =   0   'False
      TabsPerPage     =   1
      BackSheets      =   1
      BoldCurrent     =   -1  'True
      FrontTabForeColor=   16711680
      MultiRow        =   -1  'True
      New3D           =   0   'False
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
         Height          =   1680
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   615
         Width           =   4665
         _Version        =   327680
         _ExtentX        =   8229
         _ExtentY        =   2963
         _StockProps     =   70
         ConvInfo        =   1418783674
         BevelOuter      =   0
         Begin PVNumericLib.PVNumeric pvnPoints 
            Height          =   345
            Left            =   2730
            TabIndex        =   1
            ToolTipText     =   "Число точек, которыми заменяются наклонные участки"
            Top             =   420
            Width           =   825
            _Version        =   393216
            _ExtentX        =   1455
            _ExtentY        =   609
            _StockProps     =   253
            Text            =   "7"
            ForeColor       =   0
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
            ForeColor       =   0
            SpinButtons     =   2
            SuppressThousand=   0   'False
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            ValueMin        =   -3.4E+38
            ValueMax        =   3.4E+38
            ValueReal       =   7
            LimitValueByType=   1
            ValidateLimit   =   1
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnTh 
            Height          =   315
            Left            =   2730
            TabIndex        =   2
            ToolTipText     =   "Абсолютная величина разности индексов, необходимая, чтобы они считались различными"
            Top             =   1140
            Width           =   1785
            _Version        =   393216
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   0
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
            ForeColor       =   0
            SpinButtons     =   0
            SuppressThousand=   0   'False
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            ValueMin        =   0
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            DecimalMax      =   9
            ValidateLimit   =   1
            ChangeColor     =   -1  'True
         End
         Begin Threed.SSPanel sspI4 
            Height          =   210
            Left            =   180
            TabIndex        =   7
            Top             =   480
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   370
            _Version        =   131074
            Font3D          =   3
            ForeColor       =   0
            BackStyle       =   1
            Caption         =   "Число заменяющих точек"
            ClipControls    =   0   'False
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   210
            Left            =   180
            TabIndex        =   8
            Top             =   1185
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   370
            _Version        =   131074
            Font3D          =   3
            ForeColor       =   0
            BackStyle       =   1
            Caption         =   "Порог различимости индексов"
            ClipControls    =   0   'False
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
         End
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
         Height          =   1680
         Left            =   75045
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   4665
         _Version        =   327680
         _ExtentX        =   8229
         _ExtentY        =   2963
         _StockProps     =   70
         ConvInfo        =   1418783674
         BevelOuter      =   0
         Begin PVNumericLib.PVNumeric pvnStepInt 
            Height          =   315
            Left            =   2295
            TabIndex        =   5
            Top             =   855
            Width           =   1785
            _Version        =   393216
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,1"
            ForeColor       =   0
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
            ForeColor       =   0
            SpinButtons     =   0
            SuppressThousand=   0   'False
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            ValueMin        =   0
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            ValueReal       =   0.1
            DecimalMax      =   16
            ValidateLimit   =   1
            ChangeColor     =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   210
            Left            =   570
            TabIndex        =   10
            Top             =   900
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   370
            _Version        =   131074
            Font3D          =   3
            ForeColor       =   0
            BackStyle       =   1
            Caption         =   "Шаг интегрирования"
            ClipControls    =   0   'False
            BevelOuter      =   0
            AutoSize        =   2
            Alignment       =   0
         End
      End
   End
   Begin Threed.SSCommand ssYes 
      Default         =   -1  'True
      Height          =   375
      Left            =   765
      TabIndex        =   3
      ToolTipText     =   "Применить опции"
      Top             =   2625
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmAlhoOpt.frx":000C
      Caption         =   "Подтверждение"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand ssNo 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3165
      TabIndex        =   4
      ToolTipText     =   "Отказаться от модификации"
      Top             =   2625
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmAlhoOpt.frx":0176
      Caption         =   "Отмена"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmAlhoOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_Res As Boolean
Private m_g As GERTNETLib.MGertNet

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssNo_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssYes_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssNo_Click
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set m_g = Nothing
End Sub

Private Sub pvnPoints_Validate(Cancel As Boolean)
  If pvnPoints.ValueInteger < 1 Or pvnPoints.ValueInteger > 100 Then
    Cancel = True
    MsgBox "Ошибочное значение числа заменяющих точек. Требуется 0 < Val <= 100", vbExclamation Or vbOKOnly, "Ошибка"
  End If
End Sub

Private Sub pvnStepInt_Validate(Cancel As Boolean)
  If pvnStepInt.ValueReal <= 0 Then
    Cancel = True
    MsgBox "Ошибочное значение шага интегрирования. Требуется Val > 0", vbExclamation Or vbOKOnly, "Ошибка"
  End If
End Sub

Private Sub pvnTh_Validate(Cancel As Boolean)
  If pvnTh.ValueReal < 0 Then
    Cancel = True
    MsgBox "Ошибочное значение порога различимости. Требуется 0 <= Val", vbExclamation Or vbOKOnly, "Ошибка"
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
      
  'ssYes.SetFocus
    
  Dim bFl As Boolean: bFl = False
      
  pvnTh_Validate bFl
  If bFl Then vsTabOpt.CurrTab = 0: Exit Sub
  
  pvnPoints_Validate bFl
  If bFl Then vsTabOpt.CurrTab = 0: Exit Sub
  
  pvnStepInt_Validate bFl
  If bFl Then vsTabOpt.CurrTab = 1: Exit Sub
    
  m_g.NDiv = pvnPoints.ValueInteger
  m_g.UnionThreshold = pvnTh.ValueReal
  m_g.IntegralAccuracy = pvnStepInt.ValueReal

  m_b_Res = True
  Hide
End Sub


Friend Sub AssData(ByVal g As MGertNet)
  
  m_b_Res = False
  
  pvnPoints.ValueInteger = g.NDiv
  pvnTh.ValueReal = g.UnionThreshold
  pvnStepInt.ValueReal = g.IntegralAccuracy
  
  Set m_g = g

End Sub

Friend Sub UnassData()
  Set m_g = Nothing
End Sub



