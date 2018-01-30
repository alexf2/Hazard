VERSION 5.00
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#104.0#0"; "AlexOCX.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{8F25C900-8346-11CF-AACD-444553540000}#6.0#0"; "pvtime.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#6.0#0"; "pvcurr.ocx"
Begin VB.Form frmOptimiz 
   BorderStyle     =   0  'None
   Caption         =   "Монитор оптимизации"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   3990
      Left            =   5310
      TabIndex        =   15
      Top             =   2235
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton optDQ 
      Caption         =   "dQ"
      Height          =   225
      Left            =   5325
      TabIndex        =   12
      ToolTipText     =   "Сортировать по убыванию дельты Q"
      Top             =   1965
      Value           =   -1  'True
      Width           =   540
   End
   Begin VB.OptionButton optCost 
      Caption         =   "Затраты"
      Height          =   225
      Left            =   5895
      TabIndex        =   13
      ToolTipText     =   "Сортировка по возрастанию стоимости"
      Top             =   1965
      Width           =   945
   End
   Begin VB.OptionButton optK 
      Caption         =   "Кэ"
      Height          =   225
      Left            =   6870
      TabIndex        =   14
      ToolTipText     =   "Сортировка по убыванию коэффициента эффективности"
      Top             =   1965
      Width           =   495
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   915
      Left            =   780
      TabIndex        =   0
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
      Begin PVTIMELib.PVTime PVTime1 
         Height          =   345
         Left            =   1905
         TabIndex        =   17
         ToolTipText     =   "Время моделирования"
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
      Begin PVProgressBarLib.PVProgressBar pbRun 
         Height          =   435
         Left            =   60
         TabIndex        =   18
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
      Begin Threed.SSPanel ssp1 
         Height          =   240
         Left            =   90
         TabIndex        =   30
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
         TabIndex        =   31
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
      TabIndex        =   19
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
         TabIndex        =   20
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
      Height          =   5040
      Left            =   75
      TabIndex        =   21
      Top             =   1845
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8890
      _Version        =   131074
      Font3D          =   1
      BackColor       =   12632256
      Caption         =   "Оптимизация"
      ShadowStyle     =   1
      Begin PVNumericLib.PVNumeric pvnBaseP 
         Height          =   315
         Left            =   165
         TabIndex        =   2
         ToolTipText     =   "Базовая вероятность происшествия"
         Top             =   1140
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
      Begin PVNumericLib.PVNumeric pvnN 
         Height          =   345
         Left            =   3525
         TabIndex        =   6
         ToolTipText     =   "Максимальное число возвращаемых вариантов"
         Top             =   1500
         Width           =   1440
         _Version        =   393216
         _ExtentX        =   2540
         _ExtentY        =   609
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
         ValueMin        =   1
         ValueMax        =   100000000
         ValueReal       =   10
         DecimalMax      =   13
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin VB.OptionButton ssoP 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   300
         TabIndex        =   9
         ToolTipText     =   "Обеспечить заданную надёжность системы ЧМС минимальными фин. затратами"
         Top             =   3623
         Width           =   300
      End
      Begin VB.OptionButton ssoMoney 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   300
         TabIndex        =   7
         ToolTipText     =   "Обеспечить max дельты Q при ограниченных фин. средствах"
         Top             =   2333
         Value           =   -1  'True
         Width           =   300
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic3 
         Height          =   1260
         Left            =   2190
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   195
         Width           =   2805
         _Version        =   327680
         _ExtentX        =   4948
         _ExtentY        =   2222
         _StockProps     =   70
         ConvInfo        =   1418783674
         BackColor       =   12632256
         BevelOuter      =   0
         Begin AlexOCX.AngularLabel AngularLabel2 
            Height          =   1185
            Left            =   2400
            TabIndex        =   33
            Top             =   15
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   2090
            FontCtx         =   "frmOptimiz.frx":0000
            Caption         =   "Быстрее-->"
         End
         Begin AlexOCX.AngularLabel AngularLabel1 
            Height          =   1185
            Left            =   90
            TabIndex        =   32
            Top             =   30
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   2090
            FontCtx         =   "frmOptimiz.frx":0058
            Caption         =   "<--Точнее"
         End
         Begin Threed.SSOption ssoLine 
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   135
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            Caption         =   "Линейная выборка"
            Value           =   -1
         End
         Begin Threed.SSOption ssoRestrict 
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   465
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            Caption         =   "Рестриктивная оценка"
         End
         Begin Threed.SSOption ssoFull 
            Height          =   285
            Left            =   360
            TabIndex        =   5
            Top             =   795
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            Caption         =   "Полная оценка"
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1200
         Left            =   675
         TabIndex        =   22
         ToolTipText     =   "Обеспечить max дельты Q при ограниченных фин. средствах"
         Top             =   1845
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   2117
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Затраты"
         ShadowStyle     =   1
         Begin PVCurrencyLib.PVCurrency pvnMoney 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Средства на внедрение мероприятий"
            Top             =   585
            Width           =   4035
            _Version        =   393216
            _ExtentX        =   7117
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
            BackColor       =   &H00C0C0C0&
            Caption         =   "Имеющиеся финансовые средства"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   105
            TabIndex        =   23
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   360
            Width           =   4020
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1200
         Left            =   675
         TabIndex        =   24
         ToolTipText     =   "Обеспечить заданную надёжность системы ЧМС минимальными фин. затратами"
         Top             =   3135
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   2117
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Вероятность происшествия <="
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnP 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Определяет минимальный уровень надёжности системы"
            Top             =   600
            Width           =   4035
            _Version        =   393216
            _ExtentX        =   7117
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
            BackColor       =   &H00C0C0C0&
            Caption         =   "Требуемая величина вероятности"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   105
            TabIndex        =   25
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   360
            Width           =   4020
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSCommand ssOptimize 
         Default         =   -1  'True
         Height          =   510
         Left            =   1620
         TabIndex        =   11
         ToolTipText     =   "Выполнить ранжировку текущего комплекса мер по эффективности"
         Top             =   4425
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   900
         _Version        =   131074
         Font3D          =   3
         BackColor       =   12632256
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
         Picture         =   "frmOptimiz.frx":00B0
         Caption         =   "Выполнить"
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "frmOptimiz.frx":01C2
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Возвращать не более:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1500
         TabIndex        =   27
         Top             =   1590
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Базовая P"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   26
         ToolTipText     =   "Базовая вероятность происшествия"
         Top             =   930
         Width           =   1860
         WordWrap        =   -1  'True
      End
      Begin Threed.SSCommand ssSetBase 
         Height          =   570
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Копирует из монитора модели текущую просчитанную модель"
         Top             =   330
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1005
         _Version        =   131074
         PictureFrames   =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmOptimiz.frx":02D4
         Caption         =   "Установить 'базу'"
         Alignment       =   6
         ButtonStyle     =   3
         PictureAlignment=   8
      End
   End
   Begin Threed.SSCheck sscheckHisto 
      Height          =   315
      Left            =   7095
      TabIndex        =   29
      Top             =   6420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      BackColor       =   12632256
      Caption         =   " Гистограмма"
   End
   Begin VB.Image imgComplSP 
      Height          =   240
      Left            =   135
      Picture         =   "frmOptimiz.frx":044A
      Top             =   555
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDQ 
      Height          =   270
      Left            =   165
      Picture         =   "frmOptimiz.frx":054C
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
   End
   Begin Threed.SSCommand ssImport 
      Height          =   375
      Left            =   5415
      TabIndex        =   16
      ToolTipText     =   "Отмечает в мониторе мер текущую выборку"
      Top             =   6450
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmOptimiz.frx":06A6
      Caption         =   "Импортировать"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmOptimiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ICallBack


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6


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

Private HighlightedNode As PVBranch

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


Private Sub Form_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxSPMonitor.GetExtraWH sWExtr, sHExtr
      
  VideoSoftElastic1.Move m_LPad, m_UPad, ssfrmOpt.Width
  VideoSoftElastic2.Move m_LPad, VideoSoftElastic1.Top + VideoSoftElastic1.Height + m_UPad / 2, VideoSoftElastic1.Width
  
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  PVTime1.Move (VideoSoftElastic1.Width - PVTime1.Width) / 2, 50
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

Private Sub Form_Unload(Cancel As Integer)
  HandsOffModel
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_csf_Coll Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_csf_Coll = Nothing
  Set m_mg_Base = Nothing
  Set m_coll_Opt = Nothing
  Set m_csp_Tmp = Nothing
  Set m_pvb_LastSel = Nothing
  UpdateControls

  Set HighlightedNode = Nothing
  PVTime1.Time = "0:00:00"
  haApp.CloseOpt
  pbRun.Value = 0
  pbRun2.Value = 0
  ssp1.Caption = 0: ssp2.Caption = 0
End Sub

Public Sub BindModel(ByVal coll As CollectionX)
  Set m_csf_Coll = coll
  pvnMoney.Value = haApp.GertNetMain.MoneyForSF
  UpdateControls
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
    bFl = Not (m_csf_Coll.key(haApp.CurrSFColl) = m_s_NameCompl)
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
  
  Dim pvB As PVBranch
  Set pvB = m_pvb_LastSel
  Do While pvB.Level > 1
    Set pvB = pvB.Get(pvtGetParent, 0)
  Loop
  If pvB.Level <> 1 Then
    MsgBox "Чтобы импортировать выборку, надо её выделить", vbExclamation Or vbOKOnly, "Нельзя импортировать"
    Exit Sub
  End If
  'jjjjjjjj
  Dim collSel As collSF
  Set collSel = m_coll_Opt(pvB.DataVariant)
  Set pvB = frmSP.tv.Branches.Get(pvtGetChild, 0)
  If pvB.IsValid Then
    Set pvB = pvB.Get(pvtGetChild, 0)
    Do While pvB.IsValid
          
      pvB.CheckBoxState = pvtChkBoxStateNotChecked
      
      Dim sf As SafetyPrecaution
      For Each sf In collSel
        Dim ilk As ILongKey: Set ilk = sf
        If ilk.Get() = pvB.Data Then
          pvB.CheckBoxState = pvtChkBoxStateChecked
          Exit For
        End If
      Next sf
     
      Set pvB = pvB.Get(pvtGetNextSibling, 0)
    Loop
    
    frxSPMonitor.vsTabModel.CurrTab = 0
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
  
  If MsgBox("Текущий набор мер: '" & m_csf_Coll.key(haApp.CurrSFColl) & "'." & vbCrLf & "Решить для него оптимизационную задачу ?", vbQuestion Or vbYesNo, "Запрос") <> vbYes Then _
    Exit Sub
        
  pvnMoney_LostFocusEvent
  
  On Error GoTo ERR_H
  pbRun.Value = 0: pbRun2.Value = 0
  ssp1.Caption = 0: ssp2.Caption = 0
  ssp2.Enabled = False
  PVTime1.Time = "0:00:00"
  
  Dim vDta
  vDta = IIf(ssoMoney.Value = True, pvnMoney.Value, m_d_P0 - pvnP.ValueReal)
  
  Dim iCln As IClonable
  Set iCln = m_csf_Coll(haApp.CurrSFColl)
  Set m_csp_Tmp = iCln.Clone()
  m_s_NameCompl = m_csf_Coll.key(haApp.CurrSFColl)
  
  m_ot_Task = IIf(ssoMoney.Value = True, OT_FixMoney_MinQ, OT_FixQ_MinMoney)

  haApp.StartOpt m_d_P0, m_mg_Base, m_csp_Tmp, m_ot_Task, vDta, CurrAlho(), pvnN.ValueInteger, frmMon.GetCurPrior(), Me
  
  Timer1.Interval = 1000
  Timer1.Enabled = True
  ssSetBase.Enabled = False
  ssOptimize.Enabled = False
  ssoMoney.Enabled = False
  ssoP.Enabled = False
  ssoLine.Enabled = False
  ssoRestrict.Enabled = False
  ssoFull.Enabled = False
  pvnN.Enabled = False
  frmCancel.blbOpt.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
End Sub

Private Sub ssSetBase_LostFocus()
  HighLight ssSetBase, False
End Sub

Private Sub ssSetBase_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssSetBase, True
End Sub

Private Sub ssSetBase_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssSetBase, False
End Sub

Private Sub ssImport_LostFocus()
  HighLight ssImport, False
End Sub

Private Sub ssImport_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  HighLight ssImport, True
End Sub

Private Sub ssImport_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
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
  ssSetBase.Enabled = Not haApp.IsCalcOpt And Not frmMon.LastRun Is Nothing
  'ssApply.Enabled = haApp.CurrSFColl <> -1
  'ssRange.Enabled = Not haApp.IsCalcRang And Not m_mg_Base Is Nothing
  ssOptimize.Enabled = Not haApp.IsCalcOpt And Not m_mg_Base Is Nothing And haApp.CurrSFColl <> -1
  ssoMoney.Enabled = ssOptimize.Enabled
  ssoP.Enabled = ssOptimize.Enabled
  ssoLine.Enabled = ssOptimize.Enabled
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
  Else
    CurrAlho = OT_Full
  End If
End Function

Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  Set m_pvb_LastSel = node
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim myFont As StdFont
    
    If Not HighlightedNode Is Nothing Then
        Set HighlightedNode.Font = Nothing
    End If
    
    If Shift = 0 Then
        Set HighlightedNode = tv.HitTest(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
        If HighlightedNode.IsValid Then
            Set myFont = tv.CreateFont
            myFont.Bold = True
            Set HighlightedNode.Font = myFont
        Else
          Set HighlightedNode = Nothing
        End If
    End If
End Sub

Private Sub Timer1_Timer()
  With PVTime1
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
  If ssp2.Enabled = False Then
    ssp2.Caption = haApp.GN_Opt.TotalIterStr
    ssp2.Enabled = True
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
  ssp2.Enabled = True
  
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
  Set pvNodeMain.CustomItemPicture = frmComplSP.imgMoney.Picture
      
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
    AddToTree m_coll_Opt(l), m_coll_Opt.key(l)
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
        Set objTmp = coll(i): sKi = coll.key(i): sKj = coll.key(j)
        coll.key(j) = ""
        coll.Replace i, coll(j): coll.key(i) = sKj
        'Set coll(i) = coll(j)
        coll.Replace j, objTmp: coll.key(j) = sKi
        'Set coll(j) = objTmp
      End If
    Next j
  Next i
End Sub

Private Sub AddToTree(ByVal cll As collSF, ByVal sNam As String)
  Dim pvB As PVBranch, pv2 As PVBranch
  Set pvB = tv.Branches.Get(pvtGetChild, 1)
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, sNam)
  pv2.ForeColor = RGB(14, 239, 61)
  pv2.DataVariant = sNam
  Set pv2.CustomItemPicture = imgComplSP.Picture
  
  Set pvB = pv2
  
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, "Затраты на внедрение = " & FormatCurrency(cll.Cost, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.imgMoney.Picture
  
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, "Дельта Q = " & Format(cll.DeltaQ, sFmt))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = imgDQ.Picture
  
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, "Выгода = " & FormatCurrency(cll.Profit, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.imgMoney.Picture
  
  Dim sStr As String
  
  If cll.Cost = 0 Then
    sStr = "не доступно (стоимость = 0)"
  Else
    sStr = Format(cll.Ke, sFmt)
  End If
  
  
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, "Кэ (выгода/затраты) = " & sStr)
  pv2.ForeColor = RGB(14, 239, 61)
  Set pv2.CustomItemPicture = frmComplSP.imgMoney.Picture
  
  Set pv2 = pvB.Add(pvtPositionInOrder, 0, "Меры")
  pv2.ForeColor = RGB(14, 239, 61)
  pv2.StandardItemPicture = pvtpicFolders
  
  Set pvB = pv2
  Dim sf As SafetyPrecaution
  For Each sf In cll
    Set pv2 = pvB.Add(pvtPositionInOrder, 0, sf.Name)
    pv2.ForeColor = RGB(108, 245, 250)
    Set pv2.CustomItemPicture = frmComplSP.imgSF.Picture
  Next sf
End Sub

