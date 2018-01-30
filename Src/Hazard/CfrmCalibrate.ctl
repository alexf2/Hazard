VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.UserControl CfrmCalibrate 
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
   ScaleHeight     =   6045
   ScaleWidth      =   9630
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5625
      Top             =   360
   End
   Begin Threed.SSFrame ssfrmParams 
      Height          =   2910
      Left            =   90
      TabIndex        =   0
      Top             =   1380
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   5133
      _Version        =   131074
      Font3D          =   1
      Caption         =   "��������� ����������"
      ShadowStyle     =   1
      Begin PVNumericLib.PVNumeric pvn4 
         Height          =   315
         Left            =   2430
         TabIndex        =   4
         Top             =   1620
         Width           =   1785
         _Version        =   393216
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,000003063411157"
         ForeColor       =   -2147483640
         BackColor       =   14811135
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColor       =   14811135
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0
         ValueMax        =   1
         ValueSpinDelta  =   0
         ValueReal       =   0.000003063411157
         DecimalMax      =   16
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
         NegativeColor   =   0
      End
      Begin PVNumericLib.PVNumeric pvn3 
         Height          =   315
         Left            =   2415
         TabIndex        =   3
         Top             =   705
         Width           =   1785
         _Version        =   393216
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,000024000324679"
         ForeColor       =   0
         BackColor       =   14811135
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColor       =   14811135
         ForeColor       =   0
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0
         ValueMax        =   1
         ValueSpinDelta  =   0
         ValueReal       =   0.000024000324679
         DecimalMax      =   16
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
         NegativeColor   =   0
      End
      Begin PVNumericLib.PVNumeric pvn2 
         Height          =   315
         Left            =   195
         TabIndex        =   2
         Top             =   1620
         Width           =   1785
         _Version        =   393216
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,000229008535665"
         ForeColor       =   0
         BackColor       =   14811135
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColor       =   14811135
         ForeColor       =   0
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0
         ValueMax        =   1
         ValueSpinDelta  =   0
         ValueReal       =   0.000229008535665
         DecimalMax      =   16
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
         NegativeColor   =   0
      End
      Begin PVNumericLib.PVNumeric pvn1 
         Height          =   315
         Left            =   195
         TabIndex        =   1
         Top             =   705
         Width           =   1785
         _Version        =   393216
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,207566991414"
         ForeColor       =   0
         BackColor       =   14811135
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColor       =   14811135
         ForeColor       =   0
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0
         ValueMax        =   1
         ValueSpinDelta  =   0
         ValueReal       =   0.207566991414
         DecimalMax      =   16
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
         NegativeColor   =   0
      End
      Begin VB.ComboBox cbnPrior 
         BackColor       =   &H00FFFBF0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "CfrmCalibrate.ctx":0000
         Left            =   975
         List            =   "CfrmCalibrate.ctx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   2490
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��������� �������������� ����"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   960
         TabIndex        =   20
         Top             =   2145
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1) ����������� ��������� ����������"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   195
         TabIndex        =   19
         Top             =   270
         Width           =   2070
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2) ����������� ������� ��������"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   195
         TabIndex        =   18
         Top             =   1185
         Width           =   1965
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "3) ����������� ����������� ��������"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2415
         TabIndex        =   17
         Top             =   270
         Width           =   1785
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "4) ����������� ������������"
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2415
         TabIndex        =   16
         Top             =   1185
         Width           =   1965
         WordWrap        =   -1  'True
      End
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   570
      Left            =   765
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   270
      Width           =   4650
      _Version        =   327680
      _ExtentX        =   8202
      _ExtentY        =   1005
      _StockProps     =   70
      ConvInfo        =   1418783674
      BackColor       =   0
      BevelOuter      =   5
      Begin PVProgressBarLib.PVProgressBar pbRun 
         Height          =   375
         Left            =   60
         TabIndex        =   22
         Top             =   105
         Width           =   4530
         _Version        =   393216
         _ExtentX        =   7990
         _ExtentY        =   661
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   0
         Appearance      =   0
         Value           =   0
         BackColor       =   0
         FillColor       =   49152
         ForeColor       =   0
      End
   End
   Begin Threed.SSFrame ssfrmRes 
      Height          =   4350
      Left            =   4710
      TabIndex        =   23
      Top             =   870
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   7673
      _Version        =   131074
      Font3D          =   1
      Caption         =   "��������� (�������� ������)"
      ShadowStyle     =   1
      Begin Threed.SSFrame SSFrame1 
         Height          =   3270
         Left            =   195
         TabIndex        =   24
         Top             =   255
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   5768
         _Version        =   131074
         Caption         =   "����������"
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnK3 
            Height          =   315
            Left            =   195
            TabIndex        =   9
            Top             =   1995
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
            ValueMax        =   4294967295
            ValueSpinDelta  =   0
            LimitValueByType=   4
            DecimalMax      =   0
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnK4 
            Height          =   315
            Left            =   180
            TabIndex        =   10
            Top             =   2745
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
            ValueMax        =   4294967295
            ValueSpinDelta  =   0
            LimitValueByType=   4
            DecimalMax      =   0
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnK1 
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Top             =   495
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
            ValueMax        =   4294967295
            ValueSpinDelta  =   0
            LimitValueByType=   4
            DecimalMax      =   0
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnK2 
            Height          =   315
            Left            =   165
            TabIndex        =   8
            Top             =   1245
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
            ValueMax        =   4294967295
            ValueSpinDelta  =   0
            LimitValueByType=   4
            DecimalMax      =   0
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "��. I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   165
            TabIndex        =   28
            Top             =   225
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "��. II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   165
            TabIndex        =   27
            Top             =   965
            Width           =   525
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "��. III"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   26
            Top             =   1705
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "��. IV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   25
            Top             =   2460
            Width           =   615
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   3270
         Left            =   2430
         TabIndex        =   29
         Top             =   255
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   5768
         _Version        =   131074
         Caption         =   "����������"
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnKD2 
            Height          =   315
            Left            =   165
            TabIndex        =   12
            Top             =   1245
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
            ValueMin        =   -1
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnKD1 
            Height          =   315
            Left            =   165
            TabIndex        =   11
            Top             =   495
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
            ValueMin        =   -1
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnKD4 
            Height          =   315
            Left            =   180
            TabIndex        =   14
            Top             =   2745
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
            ValueMin        =   -1
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin PVNumericLib.PVNumeric pvnKD3 
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   1995
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
            ValueMin        =   -1
            ValueMax        =   214748364
            ValueSpinDelta  =   0
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "��. IV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   180
            TabIndex        =   33
            Top             =   2460
            Width           =   615
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "��. III"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   32
            Top             =   1710
            Width           =   585
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "��. II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   165
            TabIndex        =   31
            Top             =   975
            Width           =   525
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "��. I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   165
            TabIndex        =   30
            Top             =   240
            Width           =   465
         End
      End
      Begin Threed.SSCommand ssApply 
         Height          =   510
         Left            =   795
         TabIndex        =   15
         Top             =   3675
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   900
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         PictureFrames   =   1
         Enabled         =   0   'False
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
         Picture         =   "CfrmCalibrate.ctx":0047
         Caption         =   "��������� � ������"
         ButtonStyle     =   2
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "CfrmCalibrate.ctx":0159
      End
   End
   Begin Threed.SSCommand ssStart 
      Height          =   510
      Left            =   975
      TabIndex        =   6
      Top             =   5295
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   900
      _Version        =   131074
      Font3D          =   3
      ForeColor       =   0
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
      Picture         =   "CfrmCalibrate.ctx":026B
      Caption         =   "������ ���������� ������"
      ButtonStyle     =   2
      PictureAlignment=   9
      PictureDisabledFrames=   1
      PictureDisabled =   "CfrmCalibrate.ctx":037D
   End
End
Attribute VB_Name = "CfrmCalibrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICallBack

Private m_mg_GertNet As MGertNet

Private Const m_def_LPad = 100
Private Const m_def_UPad = 70
Private m_i_Phaze As Integer
Private m_i_Percent As Integer

Public Caption As String

Public Property Get PssStart() As SSCommand
  Set PssStart = ssStart
End Property

Public Property Get PssApply() As SSCommand
  Set PssApply = ssApply
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get MousePointer() As Long
  MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal lP As Long)
  UserControl.MousePointer = lP
End Property


Public Property Get MinW() As Single
  MinW = 9750
End Property

Public Property Get MinH() As Single
  MinH = 5835
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
  End If
End Function


Private Sub cbnPrior_Click()
  If Not haApp Is Nothing Then
    If Not haApp.GN_Run Is Nothing Then _
      haApp.GN_Run.CalcSpeed = GetCurPrior()
  End If
End Sub


Private Function GetCurPrior() As GERTNETLib.ThrdPriority
  If cbnPrior.ListIndex < 0 Then
    GetCurPrior = TP_NORMAL
  Else
    Select Case cbnPrior.ListIndex
      Case 0
        GetCurPrior = TP_NORMAL
      Case 1
        GetCurPrior = TP_ABOVE_NORMAL
      Case 2
        GetCurPrior = TP_HIGHEST
      Case 3
        GetCurPrior = TP_BELOW_NORMAL
      Case 4
        GetCurPrior = TP_LOWEST
    End Select
  End If
End Function


Private Sub UserControl_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxModelMon.GetExtraWH sWExtr, sHExtr
  
  VideoSoftElastic1.Move m_def_LPad, m_def_UPad, ScaleWidth - 2 * m_def_LPad - sWExtr
  ssfrmParams.Move VideoSoftElastic1.Left + (VideoSoftElastic1.Width - ssfrmParams.Width - ssfrmRes.Width - m_def_LPad) / 2, VideoSoftElastic1.Top + VideoSoftElastic1.Height + m_def_UPad
  ssfrmRes.Move ssfrmParams.Left + ssfrmParams.Width + m_def_LPad, ssfrmParams.Top
  Dim lH As Long
  If ssfrmRes.Height > ssfrmParams.Height Then
    lH = ssfrmRes.Height
  Else
    lH = ssfrmParams.Height
  End If
  ssStart.Move VideoSoftElastic1.Left, ssfrmRes.Top + lH + m_def_UPad, VideoSoftElastic1.Width
  
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  Dim sH As Single: sH = ScaleY(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
    
  pbRun.Move 50 + sW, 100, VideoSoftElastic1.Width - 2 * (60 + sW), VideoSoftElastic1.Height - 200
  
End Sub


Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_mg_GertNet = Nothing
    
  haApp.CloseCalc
  pbRun.Value = 0
End Sub

Public Sub SetupNumbers()
 If Not haApp Is Nothing Then
   If haApp.HaveDoc Then
     SetupPVNumeric haApp.GertNetMain.NFormatPr, haApp.GertNetMain.NDigitsPr, pvn1, pvn2, pvn3, pvn4
   End If
 End If
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
  SetupNumbers
    
  With mgn
    pvnK1.ValueReal = .VParam(1)
    pvnK2.ValueReal = .VParam(2)
    pvnK3.ValueReal = .VParam(3)
    pvnK4.ValueReal = .VParam(4)
    
    pvnKD1.ValueReal = .VParamIndistinct(1)
    pvnKD2.ValueReal = .VParamIndistinct(2)
    pvnKD3.ValueReal = .VParamIndistinct(3)
    pvnKD4.ValueReal = .VParamIndistinct(4)
    
    pvn1.ValueReal = .VProbability(1)
    pvn2.ValueReal = .VProbability(2)
    pvn3.ValueReal = .VProbability(3)
    pvn4.ValueReal = .VProbability(4)
  End With
  
  ssApply.Enabled = False
End Sub

Private Sub pvnK1_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnK2_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnK3_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnK4_Change()
  ssApply.Enabled = True
End Sub

Private Sub pvnKD1_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnKD2_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnKD3_Change()
  ssApply.Enabled = True
End Sub
Private Sub pvnKD4_Change()
  ssApply.Enabled = True
End Sub

Private Sub ssApply_Click()
  If Me.IsBoundModel Then
            
    With m_mg_GertNet
      .VParam(1) = pvnK1.ValueInteger
      .VParam(2) = pvnK2.ValueInteger
      .VParam(3) = pvnK3.ValueInteger
      .VParam(4) = pvnK4.ValueInteger
      
      .VParamIndistinct(1) = pvnKD1.ValueReal
      .VParamIndistinct(2) = pvnKD2.ValueReal
      .VParamIndistinct(3) = pvnKD3.ValueReal
      .VParamIndistinct(4) = pvnKD4.ValueReal
      
      .VProbability(1) = pvn1.ValueReal
      .VProbability(2) = pvn2.ValueReal
      .VProbability(3) = pvn3.ValueReal
      .VProbability(4) = pvn4.ValueReal
    End With
    
    ssApply.Enabled = False
  Else
    MsgBox "��� ������", vbExclamation Or vbOKOnly, "������"
  End If
End Sub

Private Sub ssStart_Click()
  If Not IsBoundModel Then
    MsgBox "��� ������", vbOKOnly, "������"
    Exit Sub
  End If
  
  Dim dSumm As Double
  dSumm = pvn1.ValueReal + pvn2.ValueReal + pvn3.ValueReal + pvn4.ValueReal
  If dSumm = 0 Or dSumm >= 1 Then
    MsgBox "��������� �����������: ������ ���� � ����� > 0 � < 1", vbExclamation Or vbOKOnly, "������"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  pbRun.Value = 0
  m_i_Phaze = 1
  m_i_Percent = 0
  pvn1.Enabled = False: pvn2.Enabled = False: pvn3.Enabled = False: pvn4.Enabled = False
  haApp.StartCalibrate 1, pvn1.ValueReal, pvn2.ValueReal, pvn3.ValueReal, pvn4.ValueReal, GetCurPrior(), Me
  ssStart.Enabled = False
  frmMon.PssStart.Enabled = False
  frmCancel.blbRunOnce.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "������"
  Err.Clear
End Sub

Public Sub ICallBack_EndOfWork(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc True, dt, gn
End Sub
Public Sub ICallBack_EndOfWork2(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub ICallBack_EndOfWork3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub


Public Sub ICallBack_StepOfWork(ByVal iPercent As Integer)
  If m_i_Phaze = 1 Then
    pbRun.Value = iPercent / 2
  Else
    pbRun.Value = 50 + iPercent / 2
  End If
End Sub
Public Sub ICallBack_StepOfWork2(ByVal iPercent As Integer)
End Sub
Public Sub ICallBack_StepOfWork3(ByVal iPercent As Integer)
End Sub


Public Sub ICallBack_Cancel(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc False, dt, gn
End Sub
Public Sub ICallBack_Cancel3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub

Public Sub ICallBack_ErrorCalc(ByVal sMsg As String, ByVal gn As MGertNet)
  Dim dt As Date
  dt = gn.TimeWork
  TermCalc False, dt, gn
  MsgBox sMsg, vbExclamation Or vbOKOnly, "������"
End Sub

Private Sub TermCalc(ByVal bOk As Boolean, ByVal dt As Date, Optional ByVal gn As MGertNet = Nothing)
  If Not bOk Then
    ssApply.Enabled = False
    Timer1.Enabled = False
    m_i_Phaze = 3
    
    frmCancel.blbRunOnce.PulseStop
    ssStart.Enabled = True
    frmMon.PssStart.Enabled = True
    pvn1.Enabled = True: pvn2.Enabled = True: pvn3.Enabled = True: pvn4.Enabled = True
    
    MsgBox "���������� ��������", vbInformation Or vbOKOnly, "����������"
    Exit Sub
  End If
  If m_i_Phaze = 1 Then
    m_i_Phaze = 2
    With gn
      pvnK1.ValueInteger = .VParam(1)
      pvnK2.ValueInteger = .VParam(2)
      pvnK3.ValueInteger = .VParam(3)
      pvnK4.ValueInteger = .VParam(4)
    End With
    ssApply.Enabled = True
    Timer1.Interval = 200
    Timer1.Enabled = True
  ElseIf m_i_Phaze = 3 Then
    With gn
      pvnKD1.ValueReal = .VParamIndistinct(1)
      pvnKD2.ValueReal = .VParamIndistinct(2)
      pvnKD3.ValueReal = .VParamIndistinct(3)
      pvnKD4.ValueReal = .VParamIndistinct(4)
    End With
    ssApply.Enabled = True
    
    frmCancel.blbRunOnce.PulseStop
    ssStart.Enabled = True
    frmMon.PssStart.Enabled = True
    pvn1.Enabled = True: pvn2.Enabled = True: pvn3.Enabled = True: pvn4.Enabled = True
  End If
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  If m_i_Phaze = 2 Then
    m_i_Phaze = 3
    haApp.StartCalibrate 2, pvn1.ValueReal, pvn2.ValueReal, pvn3.ValueReal, pvn4.ValueReal, GetCurPrior(), Me
  End If
End Sub


Private Sub UserControl_Initialize()
  cbnPrior.ListIndex = 3
  m_i_Phaze = 1
  m_i_Percent = 0
  
  Caption = "������� ���������� ������"
End Sub

Private Sub UserControl_Terminate()
  HandsOffModel
End Sub
