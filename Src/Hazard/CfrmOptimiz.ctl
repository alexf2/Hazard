VERSION 5.00
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#105.0#0"; "AlexOCX.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{8F25C900-8346-11CF-AACD-444553540000}#6.0#0"; "pvtime.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#6.0#0"; "pvcurr.ocx"
Begin VB.UserControl CfrmOptimiz 
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ScaleHeight     =   7065
   ScaleWidth      =   8490
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   3990
      Left            =   5310
      TabIndex        =   14
      Top             =   2250
      Width           =   2625
      _Version        =   393216
      _ExtentX        =   4630
      _ExtentY        =   7038
      _StockProps     =   237
      ForeColor       =   0
      BackColor       =   9450828
      Appearance      =   1
      StandardDefaultPicture=   10
      AllowInPlaceEditing=   0   'False
      BackColor       =   9450828
      ForeColor       =   0
      EnableToolTips  =   0   'False
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
   Begin VB.OptionButton optK 
      Caption         =   "Кэ"
      Height          =   225
      Left            =   6870
      TabIndex        =   12
      ToolTipText     =   "Сортировка по убыванию коэффициента эффективности"
      Top             =   1965
      Width           =   495
   End
   Begin VB.OptionButton optCost 
      Caption         =   "Затраты"
      Height          =   225
      Left            =   5895
      TabIndex        =   11
      ToolTipText     =   "Сортировка по возрастанию стоимости"
      Top             =   1965
      Width           =   945
   End
   Begin VB.OptionButton optDQ 
      Caption         =   "dQ"
      Height          =   225
      Left            =   5325
      TabIndex        =   10
      ToolTipText     =   "Сортировать по убыванию дельты Q"
      Top             =   1965
      Value           =   -1  'True
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   915
      Left            =   780
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   210
      Width           =   4620
      _Version        =   327680
      _ExtentX        =   8149
      _ExtentY        =   1614
      _StockProps     =   70
      ConvInfo        =   1418783674
      BackColor       =   0
      BevelOuter      =   5
      Begin PVProgressBarLib.PVProgressBar pbRun 
         Height          =   435
         Left            =   60
         TabIndex        =   17
         Top             =   420
         Width           =   4530
         _Version        =   393216
         _ExtentX        =   7990
         _ExtentY        =   767
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   0
         Appearance      =   0
         Value           =   0
         BackColor       =   0
         FillColor       =   49152
         ForeColor       =   0
      End
      Begin PVTIMELib.PVTime PVTime1 
         Height          =   345
         Left            =   660
         TabIndex        =   18
         ToolTipText     =   "Фактическое время оптимизации"
         Top             =   75
         Width           =   1365
         _Version        =   393216
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   237
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         DisplayStyle    =   1
         ShowSeconds     =   -1  'True
         ReadOnly        =   -1  'True
         LeadZero        =   -1  'True
         TwentyFourHourClock=   -1  'True
         BackColor       =   0
         ForeColor       =   65280
         Time            =   "0:00:00"
      End
      Begin PVTIMELib.PVTime PVTimeEst 
         Height          =   345
         Left            =   2700
         TabIndex        =   32
         ToolTipText     =   "Ожидаемое время оптимизации"
         Top             =   75
         Width           =   1365
         _Version        =   393216
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   237
         ForeColor       =   65280
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         DisplayStyle    =   1
         ShowSeconds     =   -1  'True
         ReadOnly        =   -1  'True
         LeadZero        =   -1  'True
         TwentyFourHourClock=   -1  'True
         BackColor       =   0
         ForeColor       =   65280
         Time            =   "0:00:00"
      End
      Begin Threed.SSPanel ssp1 
         Height          =   240
         Left            =   90
         TabIndex        =   19
         Top             =   45
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   423
         _Version        =   131074
         ForeColor       =   16776960
         BackStyle       =   1
         Caption         =   "0"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel ssp2 
         Height          =   240
         Left            =   3435
         TabIndex        =   20
         Top             =   30
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   423
         _Version        =   131074
         ForeColor       =   16776960
         BackStyle       =   1
         Caption         =   "0"
         BevelOuter      =   0
         Alignment       =   4
      End
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
      Height          =   555
      Left            =   780
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4650
      _Version        =   327680
      _ExtentX        =   8202
      _ExtentY        =   979
      _StockProps     =   70
      ConvInfo        =   1418783674
      BackColor       =   0
      BevelOuter      =   5
      Begin PVProgressBarLib.PVProgressBar pbRun2 
         Height          =   375
         Left            =   60
         TabIndex        =   22
         Top             =   90
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
   Begin Threed.SSFrame ssfrmOpt 
      Height          =   5085
      Left            =   75
      TabIndex        =   23
      Top             =   1845
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8969
      _Version        =   131074
      Font3D          =   1
      Caption         =   "Оптимизация"
      ShadowStyle     =   1
      Begin PVNumericLib.PVNumeric pvnN 
         Height          =   330
         Left            =   165
         TabIndex        =   33
         ToolTipText     =   "Максимальное число возвращаемых вариантов"
         Top             =   1890
         Width           =   1440
         _Version        =   393216
         _ExtentX        =   2540
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "10"
         ForeColor       =   0
         BackColor       =   16776176
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
         BackColor       =   16776176
         ForeColor       =   0
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0
         ValueMax        =   4294967295
         ValueReal       =   10
         LimitValueByType=   4
         DecimalMax      =   0
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnBaseP 
         Height          =   315
         Left            =   165
         TabIndex        =   36
         ToolTipText     =   "Базовая вероятность происшествия"
         Top             =   1185
         Width           =   1860
         _Version        =   393216
         _ExtentX        =   3281
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,0"
         ForeColor       =   0
         BackColor       =   14811135
         Appearance      =   1
         BackColor       =   14811135
         ForeColor       =   0
         ReadOnly        =   -1  'True
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         DecimalMin      =   1
         DecimalMax      =   16
         ValidateLimit   =   -16364
         ChangeColor     =   -1  'True
      End
      Begin VB.OptionButton ssoMoney 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   75
         TabIndex        =   5
         ToolTipText     =   "Обеспечить max дельты Q при ограниченных фин. средствах"
         Top             =   2940
         Value           =   -1  'True
         Width           =   300
      End
      Begin VB.OptionButton ssoP 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   75
         TabIndex        =   7
         ToolTipText     =   "Обеспечить заданную надёжность системы ЧМС минимальными фин. затратами"
         Top             =   4208
         Width           =   300
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic3 
         Height          =   1665
         Left            =   2145
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   195
         Width           =   2880
         _Version        =   327680
         _ExtentX        =   5080
         _ExtentY        =   2937
         _StockProps     =   70
         ConvInfo        =   1418783674
         BevelOuter      =   0
         Begin AlexOCX.AngularLabel AngularLabel2 
            Height          =   1185
            Left            =   2505
            TabIndex        =   34
            Top             =   255
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   2090
            FontCtx         =   "CfrmOptimiz.ctx":0000
            Caption         =   "Быстрее-->"
         End
         Begin Threed.SSOption ssoInt 
            Height          =   285
            Left            =   360
            TabIndex        =   0
            Top             =   30
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            Caption         =   "Назначение единиц"
         End
         Begin Threed.SSOption ssoIntAdv 
            Height          =   285
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            Caption         =   "Локальная оптимизация"
         End
         Begin Threed.SSOption ssoFull 
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   1335
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            Caption         =   "Полная оценка"
         End
         Begin Threed.SSOption ssoRestrict 
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   1008
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            Caption         =   "Рестриктивная оценка"
         End
         Begin Threed.SSOption ssoLine 
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Top             =   682
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            Caption         =   "Линейная выборка"
            Value           =   -1
         End
         Begin AlexOCX.AngularLabel AngularLabel1 
            Height          =   1185
            Left            =   90
            TabIndex        =   25
            Top             =   255
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   2090
            FontCtx         =   "CfrmOptimiz.ctx":0058
            Caption         =   "<--Точнее"
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1200
         Left            =   450
         TabIndex        =   26
         ToolTipText     =   "Обеспечить max дельты Q при ограниченных фин. средствах"
         Top             =   2445
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   2117
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         Caption         =   "Затраты"
         ShadowStyle     =   1
         Begin PVCurrencyLib.PVCurrency pvnMoney 
            Height          =   315
            Left            =   150
            TabIndex        =   6
            ToolTipText     =   "Средства на внедрение мероприятий"
            Top             =   585
            Width           =   4095
            _Version        =   393216
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "1.500,00р."
            ForeColor       =   0
            BackColor       =   8421504
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
            DisplayFormat   =   3
            BackColor       =   8421504
            ForeColor       =   0
            HighlightColor  =   12632256
            FormatNegative  =   5
            FormatPositive  =   1
            Symbol          =   "р."
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            Value           =   1500
            ValueMax        =   910000000000000
            ValidateLimit   =   1
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Финансовые средства"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            TabIndex        =   27
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   360
            Width           =   1965
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1200
         Left            =   465
         TabIndex        =   28
         ToolTipText     =   "Обеспечить заданную надёжность системы ЧМС минимальными фин. затратами"
         Top             =   3720
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   2117
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         Caption         =   "Вероятность происшествия <="
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnP 
            Height          =   315
            Left            =   150
            TabIndex        =   8
            ToolTipText     =   "Определяет минимальный уровень надёжности системы"
            Top             =   600
            Width           =   4065
            _Version        =   393216
            _ExtentX        =   7170
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,002075669914148"
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
            Enabled         =   0   'False
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
            ValueReal       =   0.002075669914148
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Требуемая вероятность"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            TabIndex        =   29
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   360
            Width           =   2070
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSCommand ssSetBase 
         Height          =   570
         Left            =   165
         TabIndex        =   35
         ToolTipText     =   "Копирует из монитора модели текущую просчитанную модель"
         Top             =   300
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1005
         _Version        =   131074
         PictureFrames   =   1
         PictureUseMask  =   -1  'True
         Picture         =   "CfrmOptimiz.ctx":00B0
         Caption         =   "Установить 'базу'"
         Alignment       =   6
         ButtonStyle     =   3
         PictureAlignment=   8
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Базовая P"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   31
         ToolTipText     =   "Базовая вероятность происшествия"
         Top             =   975
         Width           =   1860
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Возвращать не более:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   1680
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin Threed.SSCommand ssOptimize 
         Height          =   510
         Left            =   1995
         TabIndex        =   9
         ToolTipText     =   "Выполнить ранжировку текущего комплекса мер по эффективности"
         Top             =   1905
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   900
         _Version        =   131074
         Font3D          =   3
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
         Picture         =   "CfrmOptimiz.ctx":0226
         Caption         =   "Выполнить"
         ButtonStyle     =   2
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "CfrmOptimiz.ctx":0338
      End
   End
   Begin Threed.SSCommand ssImport 
      Height          =   375
      Left            =   5415
      TabIndex        =   15
      ToolTipText     =   "Отмечает в мониторе мер текущую выборку"
      Top             =   6450
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "CfrmOptimiz.ctx":044A
      Caption         =   "Импортировать"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin VB.Image imgDQ 
      Height          =   270
      Left            =   165
      Picture         =   "CfrmOptimiz.ctx":05B4
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgComplSP 
      Height          =   240
      Left            =   135
      Picture         =   "CfrmOptimiz.ctx":070E
      Top             =   555
      Visible         =   0   'False
      Width           =   240
   End
   Begin Threed.SSCheck sscheckHisto 
      Height          =   315
      Left            =   7095
      TabIndex        =   16
      Top             =   6420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      Caption         =   " Гистограмма"
   End
End
Attribute VB_Name = "CfrmOptimiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICallBack



Private Const m_LPad = 100
Private Const m_UPad = 70
Private Const m_L2Pad = 200

Private m_csf_Coll As CollectionX
Private m_mg_Base As MGertNet
Private m_coll_Opt As CollectionX
Private m_csp_Tmp As collSF
Private m_s_NameCompl As String
Private m_pvb_LastSel As PVBranch

Private m_d_P0 As Double
Private m_c_AveDamage As Currency
Private m_ot_Task As OptTask

Private m_bSSP2Enbl As Boolean

Private HighlightedNode As PVBranch
Private m_myFont As StdFont

Public Caption As String

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
  MinW = 7800
End Property

Public Property Get MinH() As Single
  MinH = 6800
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
    UpdateControls
  End If
End Function


Private Sub UserControl_InitProperties()
  tv.StandardDefaultPicture = pvtNone
  Caption = "Монитор оптимизации"
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
End Sub

Private Sub UserControl_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxSPMonitor.GetExtraWH sWExtr, sHExtr
      
  VideoSoftElastic1.Move m_LPad, m_UPad, ssfrmOpt.Width
  VideoSoftElastic2.Move m_LPad, VideoSoftElastic1.Top + VideoSoftElastic1.Height + m_UPad / 2, VideoSoftElastic1.Width
  
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  'PVTime1.Move (VideoSoftElastic1.Width - PVTime1.Width) / 2, 50
  
  PVTime1.Move (VideoSoftElastic1.Width - 1.2 * PVTime1.Width - PVTimeEst.Width) / 2, 50
  PVTimeEst.Move PVTime1.Left + 1.2 * PVTime1.Width, PVTime1.Top
  
  pbRun.Move 60 + sW, PVTime1.Top + PVTime1.Height + 10, VideoSoftElastic1.Width - 2 * (70 + sW)
  pbRun2.Move 50 + sW, 100, VideoSoftElastic2.Width - 2 * (60 + sW), VideoSoftElastic2.Height - 200
  ssp1.Move 2 * sW, 0
  ssp2.Move VideoSoftElastic1.Width - ssp2.Width - 2 * sW, 0
  
  ssfrmOpt.Move m_LPad, VideoSoftElastic2.Top + VideoSoftElastic2.Height + m_UPad
  
  optDQ.Move ssfrmOpt.Left + ssfrmOpt.Width + m_LPad, VideoSoftElastic1.Top
  optCost.Move optDQ.Left + optDQ.Width, optDQ.Top
  optK.Move optCost.Left + optCost.Width, optCost.Top
  
  tv.Move optDQ.Left, optDQ.Top + optDQ.Height, ScaleWidth - 3 * m_LPad - ssfrmOpt.Width - sWExtr, ScaleHeight - (optDQ.Top + optDQ.Height) - 2 * m_UPad - sHExtr - ssImport.Height
  ssImport.Move tv.Left, tv.Top + tv.Height + m_UPad
  sscheckHisto.Move ssImport.Left + ssImport.Width + m_LPad / 2, ssImport.Top + (ssImport.Height - sscheckHisto.Height) / 2
End Sub


Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_csf_Coll Is Nothing)
End Property

Public Sub HandsOffModel()
  ClearTree
  Set m_csf_Coll = Nothing
  Set m_mg_Base = Nothing
  Set m_coll_Opt = Nothing
  Set m_csp_Tmp = Nothing
  Set m_pvb_LastSel = Nothing
  UpdateControls

  Set HighlightedNode = Nothing
  PVTime1.Time = "0:00:00"
  PVTimeEst.Time = "0:00:00"
  haApp.CloseOpt
  pbRun.Value = 0
  pbRun2.Value = 0
  ssp1.Caption = 0: ssp2.Caption = 0
End Sub

Public Sub SetupNumbers()

End Sub


Public Sub BindModel(ByVal coll As CollectionX)
  Set m_csf_Coll = coll
  SetupNumbers
  
  pvnMoney.Value = haApp.GertNetMain.MoneyForSF
  UpdateControls
  
  Select Case haApp.GertNetMain.OptimizationType
    Case OT_Full
      ssoFull.Value = True
    Case OT_Integer
      ssoInt.Value = True
    Case OT_IntegerAdv
      ssoIntAdv.Value = True
    Case OT_Quick
      ssoLine.Value = True
    Case OT_Quick2
      ssoRestrict.Value = True
  End Select
    
End Sub



Private Sub optCost_Click()
  FillTree
End Sub

Private Sub optDQ_Click()
  FillTree
End Sub

Private Sub optK_Click()
  FillTree
End Sub

Private Sub pvnMoney_LostFocusEvent()
  If Not haApp.GertNetMain Is Nothing Then
    If haApp.GertNetMain.MoneyForSF <> pvnMoney.Value Then
      haApp.GertNetMain.MoneyForSF = pvnMoney.Value
      If Not m_mg_Base Is Nothing Then _
        m_mg_Base.MoneyForSF = pvnMoney.Value
    End If
  End If
End Sub


Private Sub sscheckHisto_Click(Value As Integer)
  If sscheckHisto.Value = True Then
    If m_coll_Opt Is Nothing Then
      sscheckHisto.Value = False
      MsgBox "Нет результатов", vbOKOnly Or vbExclamation, "Ошибка"
      Exit Sub
    End If
    frmHistoOpt.LoadColl m_coll_Opt, CurrSorting, m_s_NameCompl
    frmHistoOpt.Visible = True
  Else
    frmHistoOpt.Visible = False
    frmHistoOpt.Clear
  End If
End Sub

Private Sub ssImport_Click()
  Dim bFl As Boolean
  bFl = (haApp.CurrSFColl = -1)
  If Not bFl Then
    bFl = Not (m_csf_Coll.Key(haApp.CurrSFColl) = m_s_NameCompl)
  End If
  
  If bFl Then
    MsgBox "Чтобы импортировать выборку, требуется, чтобы текущим был комплекс мер, для которого выполнена оптимизация", vbExclamation Or vbOKOnly, "Нельзя импортировать"
    Exit Sub
  End If
  
  bFl = (m_pvb_LastSel Is Nothing)
  If Not bFl Then
    bFl = Not m_pvb_LastSel.IsValid
  End If
  If bFl Then
    MsgBox "Чтобы импортировать выборку, надо её выделить", vbExclamation Or vbOKOnly, "Нельзя импортировать"
    Exit Sub
  End If
  
  Dim pvb As PVBranch
  Set pvb = m_pvb_LastSel
  Do While pvb.Level > 1
    Set pvb = pvb.Get(pvtGetParent, 0)
  Loop
  If pvb.Level <> 1 Then
    MsgBox "Чтобы импортировать выборку, надо её выделить", vbExclamation Or vbOKOnly, "Нельзя импортировать"
    Exit Sub
  End If
  'jjjjjjjj
  Dim collSel As collSF
  Set collSel = m_coll_Opt(pvb.DataVariant)
  Set pvb = frmSP.Ptv.Branches.Get(pvtGetChild, 0)
  If pvb.IsValid Then
    Set pvb = pvb.Get(pvtGetChild, 0)
    Do While pvb.IsValid
          
      pvb.CheckBoxState = pvtChkBoxStateNotChecked
      
      Dim sf As SafetyPrecaution
      For Each sf In collSel
        Dim ilk As ILongKey: Set ilk = sf
        If ilk.Get() = pvb.Data Then
          pvb.CheckBoxState = pvtChkBoxStateChecked
          Exit For
        End If
      Next sf
     
      Set pvb = pvb.Get(pvtGetNextSibling, 0)
    Loop
    
    frxSPMonitor.vsTabModel.CurrTab = 1
  Else
    MsgBox "В мониторе комплексов мер не загружен комплекс", vbExclamation Or vbOKOnly, "Нельзя импортировать"
  End If
End Sub

Private Sub ssoMoney_Click()
  If ssoMoney.Value = True Then _
    pvnP.Enabled = False: pvnMoney.Enabled = True
End Sub

Private Sub ssoP_Click()
  If ssoP.Value = True Then _
    pvnP.Enabled = True: pvnMoney.Enabled = False
End Sub

Private Sub ssOptimize_Click()
  If Not IsBoundModel Then
    MsgBox "Нет модели", vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If m_mg_Base Is Nothing Then
    MsgBox "Нет базовой модели", vbOKOnly, "Нельзя ранжировать"
    Exit Sub
  End If
    
  If haApp.CurrSFColl = -1 Then
    MsgBox "Нет текущего набора мер", vbExclamation Or vbOKOnly, "Нельзя начать операцию"
    Exit Sub
  End If
  
  'If MsgBox("Текущий набор мер: '" & m_csf_Coll.Key(haApp.CurrSFColl) & "'." & vbCrLf & "Решить для него оптимизационную задачу ?", vbQuestion Or vbYesNo, "Запрос") <> vbYes Then _
   ' Exit Sub
    
  If ssoP.Value = True And m_d_P0 - pvnP.ValueReal <= 0 Then
    MsgBox "Требуемый уровень безопасности уже есть", vbExclamation Or vbOKOnly, "Не требуется начинать операцию"
    Exit Sub
  End If
        
  pvnMoney_LostFocusEvent
  
  On Error GoTo ERR_H
  pbRun.Value = 0: pbRun2.Value = 0
  ssp1.Caption = 0: ssp2.Caption = 0
  m_bSSP2Enbl = False
  PVTime1.Time = "0:00:00"
  PVTimeEst.Time = "0:00:00"
  
  Dim vDta
  vDta = IIf(ssoMoney.Value = True, pvnMoney.Value, m_d_P0 - pvnP.ValueReal)
  
  Dim icln As IClonable
  Set icln = m_csf_Coll(haApp.CurrSFColl)
  Set m_csp_Tmp = icln.Clone()
  m_s_NameCompl = m_csf_Coll.Key(haApp.CurrSFColl)
  
  m_ot_Task = IIf(ssoMoney.Value = True, OT_FixMoney_MinQ, OT_FixQ_MinMoney)

  haApp.StartOpt m_d_P0, m_mg_Base, m_csp_Tmp, m_ot_Task, vDta, CurrAlho(), pvnN.ValueInteger, frmMon.GetCurPrior(), Me
  
  Timer1.Interval = 1000
  Timer1.Enabled = True
  ssSetBase.Enabled = False
  ssOptimize.Enabled = False
  ssoMoney.Enabled = False
  ssoP.Enabled = False
  ssoLine.Enabled = False
  ssoInt.Enabled = False: ssoIntAdv.Enabled = False
  ssoRestrict.Enabled = False
  ssoFull.Enabled = False
  pvnN.Enabled = False
  frmCancel.blbOpt.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssSetBase_LostFocus()
  HighLight ssSetBase, False
End Sub

Private Sub ssSetBase_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssSetBase, True
End Sub

Private Sub ssSetBase_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssSetBase, False
End Sub

Private Sub ssImport_LostFocus()
  HighLight ssImport, False
End Sub

Private Sub ssImport_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssImport, True
End Sub

Private Sub ssImport_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssImport, False
End Sub

Private Sub ssSetBase_Click()
  Set m_mg_Base = frmMon.LastRun
  
  Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double
  m_mg_Base.GetInfSampleK 73, lMin, lMax, dMx, dDx
  m_d_P0 = dMx
  pvnBaseP.ValueReal = dMx
  
  UpdateControls
End Sub


Friend Sub UpdateControls()
  If frmMon Is Nothing Then
    ssSetBase.Enabled = False
  Else
    ssSetBase.Enabled = Not haApp.IsCalcOpt And Not frmMon.LastRun Is Nothing
  End If
  
  'ssApply.Enabled = haApp.CurrSFColl <> -1
  'ssRange.Enabled = Not haApp.IsCalcRang And Not m_mg_Base Is Nothing
  ssOptimize.Enabled = Not haApp.IsCalcOpt And Not m_mg_Base Is Nothing And haApp.CurrSFColl <> -1
  ssoMoney.Enabled = ssOptimize.Enabled
  ssoP.Enabled = ssOptimize.Enabled
  ssoLine.Enabled = ssOptimize.Enabled
  ssoInt.Enabled = ssOptimize.Enabled
  ssoIntAdv.Enabled = ssOptimize.Enabled
  ssoRestrict.Enabled = ssOptimize.Enabled
  ssoFull.Enabled = ssOptimize.Enabled
  pvnN.Enabled = ssOptimize.Enabled
End Sub

Friend Function CurrSorting() As GERTNETLib.CollSFSorting
  If optDQ.Value = True Then
    CurrSorting = CSFS_DeltaQDescending
  ElseIf optCost.Value = True Then
    CurrSorting = CSFS_Price
  Else
    CurrSorting = CSFS_KDescending
  End If
End Function

Friend Function CurrAlho() As GERTNETLib.OptType
  If ssoLine.Value = True Then
    CurrAlho = OT_Quick
  ElseIf ssoRestrict.Value = True Then
    CurrAlho = OT_Quick2
  ElseIf ssoFull.Value = True Then
    CurrAlho = OT_Full
  ElseIf ssoInt.Value = True Then
    CurrAlho = OT_Integer
  ElseIf ssoIntAdv.Value = True Then
    CurrAlho = OT_IntegerAdv
  End If
End Function

Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  Set m_pvb_LastSel = node
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Dim myFont As StdFont
    
    If Not HighlightedNode Is Nothing Then
        Set HighlightedNode.Font = Nothing
    End If
    
    If Shift = 0 Then
        Set HighlightedNode = tv.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
        If HighlightedNode.IsValid Then
            'Set myFont = tv.CreateFont
            'myFont.Bold = True
            Set HighlightedNode.Font = m_myFont
        Else
          Set HighlightedNode = Nothing
        End If
    End If
End Sub

Private Sub Timer1_Timer()
  With PVTime1
    If .Second Mod 3 = 0 Then _
      PVTimeEst.Time = FormatDateTime(haApp.GN_Opt.OptCount, vbLongTime)
      
    .Second = .Second + 1
    If .Second >= 60 Then .Minute = .Minute + 1: .Second = 0
    If .Minute >= 60 Then .Hour = .Hour + 1: .Minute = 0
  End With
End Sub

Public Sub ICallBack_EndOfWork(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub ICallBack_EndOfWork2(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub ICallBack_EndOfWork3(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc True, dt, gn
End Sub


Public Sub ICallBack_StepOfWork(ByVal iPercent As Integer)
  pbRun2.Value = iPercent
End Sub
Public Sub ICallBack_StepOfWork2(ByVal iPercent As Integer)
End Sub
Public Sub ICallBack_StepOfWork3(ByVal iPercent As Integer)
  pbRun.Value = iPercent
  ssp1.Caption = haApp.GN_Opt.CurrIterStr
  If Not m_bSSP2Enbl Then
    ssp2.Caption = haApp.GN_Opt.TotalIterStr
    m_bSSP2Enbl = True
  End If
End Sub


Public Sub ICallBack_Cancel(ByVal dt As Date, ByVal gn As MGertNet)
  'TermCalc False, dt, gn
End Sub
Public Sub ICallBack_Cancel3(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc False, dt, gn
End Sub


Public Sub ICallBack_ErrorCalc(ByVal sMsg As String, ByVal gn As MGertNet)
  Dim dt As Date
  dt = gn.TimeWork
  TermCalc False, dt, gn
  MsgBox sMsg, vbExclamation Or vbOKOnly, "Ошибка"
End Sub

Private Sub TermCalc(ByVal bOk As Boolean, ByVal dt As Date, Optional ByVal gn As MGertNet = Nothing)
  Timer1.Enabled = False
  frmCancel.blbOpt.PulseStop
  PVTime1.Time = Format(dt, "h:mm:ss")
  m_bSSP2Enbl = True
  
  If bOk Then
    Select Case m_ot_Task
      Case OT_FixMoney_MinQ
        optDQ.Value = True
      Case OT_FixQ_MinMoney
        optCost.Value = True
    End Select
  End If
  
  CreateReport bOk, gn
  
  If Not bOk Then _
    MsgBox "Оптимизация прервана", vbInformation Or vbOKOnly, "Информация"
  
  Dim bEn As Boolean
  bEn = haApp.CurrSFColl <> -1
  
  ssSetBase.Enabled = bEn
  ssOptimize.Enabled = bEn
  ssoMoney.Enabled = bEn
  ssoP.Enabled = bEn
  ssoLine.Enabled = bEn
  ssoInt.Enabled = bEn: ssoIntAdv.Enabled = bEn
  ssoRestrict.Enabled = bEn
  ssoFull.Enabled = bEn
  pvnN.Enabled = bEn
End Sub

Friend Sub CreateReport(ByVal bOk As Boolean, ByVal gn As MGertNet)
  On Error GoTo ERR_H
  
  Dim bm As New CBeam
  bm.Beam Me
  
  m_c_AveDamage = gn.AverageDamage
  
  ClearTree
  
  If bOk Then
    Dim arrSelects() As Object, l As Long, lCnt As Long
    arrSelects = gn.OptimizResultsGetAndClear
    
    Set m_coll_Opt = New CollectionX
    lCnt = 1
    Dim cAveDamage As Currency: cAveDamage = haApp.GertNetMain.AverageDamage
    For l = LBound(arrSelects) To UBound(arrSelects)
      Dim cllTmp As collSF
      Set cllTmp = arrSelects(l)
      m_coll_Opt.Add cllTmp, "Выборка " & CStr(lCnt)
      
      With cllTmp
        .Profit = .DeltaQ * cAveDamage
        If .Cost <> 0 Then
          .Ke = .DeltaQ * cAveDamage / .Cost
        Else
          .Ke = 0
        End If
      End With
      
      lCnt = lCnt + 1
    Next l
    
    FillTree

    If haApp.CreRepOpt Then
      Dim vDta
      vDta = IIf(m_ot_Task = OT_FixMoney_MinQ, pvnMoney.Value, pvnP.ValueReal)
      Dim rep As New CRepOpt
      rep.InitReport haApp, m_coll_Opt, m_s_NameCompl, m_d_P0, m_c_AveDamage, m_ot_Task, vDta, gn.UserProp
      Dim ir As IRepItem
      Set ir = rep
      haApp.Rep2.Add ir
      
      UpdateGUI_Rep2 haApp
    End If
    
  End If
    
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub ClearTree()
  Dim node As PVBranch, node2 As PVBranch

  Set node = tv.Branches.Get(pvtGetChild, 0)
  While node.IsValid
    Set node2 = node.Get(pvtGetNextSibling, 0)
    node.Remove
    Set node = node2
  Wend
End Sub

Private Sub FillTree()
  tv.Redraw = False
  On Error GoTo ERR_H
  
  If tv.Count > 0 Then ClearTree
      
  Dim pvNodeMain As PVBranch
  Set pvNodeMain = tv.Branches.Add(pvtPositionInOrder, 0, "Средний ущерб от происшествия = " & FormatCurrency(m_c_AveDamage, 2, vbTrue, vbFalse, vbTrue))
  pvNodeMain.ForeColor = RGB(237, 232, 83)
  'pvNodeMain.StandardItemPicture = pvtpicFolders
  Set pvNodeMain.CustomItemPicture = frmComplSP.PimgMoney.Picture
      
  If m_coll_Opt Is Nothing Then
    Set pvNodeMain = tv.Branches.Add(pvtPositionInOrder, 0, "Пусто")
  Else
    Set pvNodeMain = tv.Branches.Add(pvtPositionInOrder, 0, "Оптимальные варианты для: '" & m_s_NameCompl & "'")
  End If
  pvNodeMain.ForeColor = RGB(237, 232, 83)
  pvNodeMain.StandardItemPicture = pvtpicFolders
  
  If m_coll_Opt Is Nothing Then
    tv.Redraw = True
    Exit Sub
  End If
  
  SortSelects m_coll_Opt, CurrSorting
  
  Dim l As Long:
  For l = 1 To m_coll_Opt.Count
    AddToTree m_coll_Opt(l), m_coll_Opt.Key(l)
  Next l
  
  ExpandAll tv
  tv.Redraw = True
  
  If frmHistoOpt.Visible Then
    If sscheckHisto.Value = True Then
      frmHistoOpt.LoadColl m_coll_Opt, CurrSorting, m_s_NameCompl
    Else
      'frmHistoOpt.Clear
    End If
  End If
  
  Exit Sub
  
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub SortSelects(ByVal coll As CollectionX, ByVal srt As GERTNETLib.CollSFSorting)
  Dim i As Long, j As Long, i2 As Long, bCmp As Boolean
  i2 = coll.Count
  For i = 1 To i2 - 1
    For j = i + 1 To i2
      Select Case srt
        Case CSFS_DeltaQDescending
          bCmp = (coll(i).DeltaQ < coll(j).DeltaQ)
        Case CSFS_PriceDescending
          bCmp = (coll(i).Cost < coll(j).Cost)
        Case CSFS_Price
          bCmp = (coll(i).Cost > coll(j).Cost)
        Case CSFS_KDescending
          bCmp = (coll(i).Ke < coll(j).Ke)
      End Select
      If bCmp Then
        Dim objTmp As Object, sKi As String, sKj As String
        Set objTmp = coll(i): sKi = coll.Key(i): sKj = coll.Key(j)
        coll.Key(j) = ""
        coll.Replace i, coll(j): coll.Key(i) = sKj
        'Set coll(i) = coll(j)
        coll.Replace j, objTmp: coll.Key(j) = sKi
        'Set coll(j) = objTmp
      End If
    Next j
  Next i
End Sub

Private Sub AddToTree(ByVal cll As collSF, ByVal sNam As String)
  Dim pvb As PVBranch, pv2 As PVBranch
  Set pvb = tv.Branches.Get(pvtGetChild, 1)
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, sNam)
  pv2.ForeColor = RGB(14, 239, 61)
  pv2.DataVariant = sNam
  Set pv2.CustomItemPicture = imgComplSP.Picture
  
  Set pvb = pv2
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Затраты на внедрение = " & FormatCurrency(cll.Cost, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.PimgMoney.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Дельта Q = " & Format(cll.DeltaQ, haApp.FmtDbl))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = imgDQ.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Выгода = " & FormatCurrency(cll.Profit, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.PimgMoney.Picture
  
  Dim sStr As String
  
  If cll.Cost = 0 Then
    sStr = "не доступно (стоимость = 0)"
  Else
    sStr = Format(cll.Ke, haApp.FmtDbl2)
  End If
  
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Кэ (выгода/затраты) = " & sStr)
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.PimgMoney.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Меры")
  pv2.ForeColor = RGB(14, 239, 61)
  pv2.StandardItemPicture = pvtpicFolders
  
  Set pvb = pv2
  Dim sf As SafetyPrecaution
  For Each sf In cll
    Set pv2 = pvb.Add(pvtPositionInOrder, 0, sf.Name)
    pv2.ForeColor = RGB(108, 245, 250)
    Set pv2.CustomItemPicture = frmComplSP.PimgSF.Picture
  Next sf
End Sub



Private Sub UserControl_Terminate()
  HandsOffModel
End Sub
