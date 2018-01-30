VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{8F25C900-8346-11CF-AACD-444553540000}#6.0#0"; "pvtime.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#6.0#0"; "pvcurr.ocx"
Begin VB.UserControl CfrmSP 
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   ScaleHeight     =   6465
   ScaleWidth      =   8880
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   3480
      Left            =   4635
      TabIndex        =   16
      Top             =   2460
      Width           =   3495
      _Version        =   393216
      _ExtentX        =   6165
      _ExtentY        =   6138
      _StockProps     =   237
      ForeColor       =   0
      BackColor       =   14811135
      Appearance      =   1
      StandardDefaultPicture=   10
      AllowInPlaceEditing=   0   'False
      BackColor       =   14811135
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
   Begin VB.OptionButton optDQ 
      Caption         =   "dQ"
      Height          =   225
      Left            =   4590
      TabIndex        =   11
      ToolTipText     =   "Сортировать по убыванию дельты Q"
      Top             =   2160
      Value           =   -1  'True
      Width           =   540
   End
   Begin VB.OptionButton optCost 
      Caption         =   "Затраты"
      Height          =   225
      Left            =   5152
      TabIndex        =   12
      ToolTipText     =   "Сортировка по возрастанию стоимости"
      Top             =   2160
      Width           =   945
   End
   Begin VB.OptionButton optK 
      Caption         =   "Кэ"
      Height          =   225
      Left            =   6150
      TabIndex        =   13
      ToolTipText     =   "Сортировка по убыванию коэффициента эффективности"
      Top             =   2160
      Width           =   495
   End
   Begin Threed.SSFrame ssfrmRang 
      Height          =   1920
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   3387
      _Version        =   131074
      Font3D          =   1
      Caption         =   "Ранжировка мер по эффективности"
      ShadowStyle     =   1
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   0
         Top             =   0
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
         Height          =   915
         Left            =   1800
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   4620
         _Version        =   327680
         _ExtentX        =   8149
         _ExtentY        =   1614
         _StockProps     =   70
         ConvInfo        =   1418783674
         BackColor       =   0
         BevelOuter      =   5
         Begin PVTIMELib.PVTime PVTimeEst 
            Height          =   345
            Left            =   2655
            TabIndex        =   31
            ToolTipText     =   "Ожидаемое время ранжировки"
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
         Begin PVTIMELib.PVTime PVTime1 
            Height          =   345
            Left            =   990
            TabIndex        =   19
            ToolTipText     =   "Фактическое время рпнжировки"
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
            TabIndex        =   20
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
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
         Height          =   555
         Left            =   1800
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1215
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
      Begin Threed.SSCheck sscheckHisto 
         Height          =   540
         Left            =   180
         TabIndex        =   1
         Top             =   1095
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   953
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   0
         Caption         =   " Гистограмма эффективности"
      End
      Begin Threed.SSCommand ssRange 
         Height          =   510
         Left            =   150
         TabIndex        =   0
         ToolTipText     =   "Выполнить ранжировку текущего комплекса мер по эффективности"
         Top             =   450
         Width           =   1575
         _ExtentX        =   2778
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
         Picture         =   "CfrmSP.ctx":0000
         Caption         =   "Выполнить"
         ButtonStyle     =   2
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "CfrmSP.ctx":0112
      End
   End
   Begin Threed.SSFrame ssfrmParams 
      Height          =   3900
      Left            =   60
      TabIndex        =   23
      Top             =   2010
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   6879
      _Version        =   131074
      Font3D          =   1
      Caption         =   "Параметры моделирования"
      ShadowStyle     =   1
      Begin PVCurrencyLib.PVCurrency pvnCost 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   510
         Width           =   4035
         _Version        =   393216
         _ExtentX        =   7117
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "10.000.000,00р."
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
         Value           =   10000000
         ValueMax        =   910000000000000
         ValidateLimit   =   1
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2235
         Left            =   150
         TabIndex        =   24
         Top             =   855
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3942
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   255
         Caption         =   "Вероятность происшествия"
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnRealP 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
            ToolTipText     =   "Получена прогоном модели после применения выбранных мер"
            Top             =   495
            Width           =   1860
            _Version        =   393216
            _ExtentX        =   3281
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,0"
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
            ReadOnly        =   -1  'True
            SpinButtons     =   0
            SuppressThousand=   0   'False
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            DecimalMin      =   1
            DecimalMax      =   16
            ValidateLimit   =   -16364
            ChangeColor     =   -1  'True
            NegativeColor   =   0
         End
         Begin PVCurrencyLib.PVCurrency pvnProfit 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Реальная выгода"
            Top             =   1785
            Width           =   1860
            _Version        =   393216
            _ExtentX        =   3281
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,00р."
            ForeColor       =   -2147483640
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
            FormatNegative  =   5
            FormatPositive  =   1
            Symbol          =   "р."
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            Value           =   0
            ValueMax        =   910000000000000
            ValidateLimit   =   1
         End
         Begin PVNumericLib.PVNumeric pvnBaseP 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Базовая вероятность происшествия"
            Top             =   1110
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
            NegativeColor   =   0
         End
         Begin PVNumericLib.PVNumeric pvnCalcedP 
            Height          =   315
            Left            =   2055
            TabIndex        =   7
            ToolTipText     =   "Получена суммированием дельт Q - результатов прогонов модели после применения отдельных мер"
            Top             =   1110
            Width           =   1860
            _Version        =   393216
            _ExtentX        =   3281
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,0"
            ForeColor       =   0
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
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
            NegativeColor   =   0
         End
         Begin PVNumericLib.PVNumeric pvnProfitCost 
            Height          =   315
            Left            =   2055
            TabIndex        =   8
            ToolTipText     =   "Реальный Кэ"
            Top             =   1785
            Width           =   1860
            _Version        =   393216
            _ExtentX        =   3281
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0,0"
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
            ReadOnly        =   -1  'True
            SpinButtons     =   0
            SuppressThousand=   0   'False
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            DecimalMin      =   1
            DecimalMax      =   16
            ValidateLimit   =   -16364
            ChangeColor     =   -1  'True
            NegativeColor   =   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Базовая P"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Базовая вероятность происшествия"
            Top             =   900
            Width           =   1860
            WordWrap        =   -1  'True
         End
         Begin Threed.SSCommand ssSetBase 
            Height          =   570
            Left            =   315
            TabIndex        =   3
            ToolTipText     =   "Копирует из монитора модели текущую просчитанную модель"
            Top             =   300
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   1005
            _Version        =   131074
            PictureFrames   =   1
            PictureUseMask  =   -1  'True
            Picture         =   "CfrmSP.ctx":0224
            Caption         =   "Установить 'базу'"
            Alignment       =   6
            ButtonStyle     =   3
            PictureAlignment=   8
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Расчётная дельта Q"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2055
            TabIndex        =   28
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   885
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Кэ (выгода/затраты)"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2055
            TabIndex        =   27
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   1560
            Width           =   1800
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Выгода от внедрения"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   1560
            Width           =   1740
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Реальная дельта Q"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2055
            TabIndex        =   25
            ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
            Top             =   270
            Width           =   1830
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Величина среднего ущерба от происшествия"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   30
         ToolTipText     =   "В качестве базовой из редактора моделей берётся текущая просчитанная модель"
         Top             =   255
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin Threed.SSCommand ssApply 
         Height          =   540
         Left            =   2955
         TabIndex        =   10
         ToolTipText     =   "Применить выбранные меры к модели (внести в неё изменения)"
         Top             =   3240
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   953
         _Version        =   131074
         PictureFrames   =   1
         PictureUseMask  =   -1  'True
         Picture         =   "CfrmSP.ctx":039A
         Caption         =   "Применить >>"
         Alignment       =   6
         ButtonStyle     =   3
         PictureAlignment=   8
      End
      Begin Threed.SSCommand ssCalc 
         Height          =   510
         Left            =   150
         TabIndex        =   9
         ToolTipText     =   "Расчитать вероятность происшествия после внедрения выбранных мер и оценить выгоду"
         Top             =   3240
         Width           =   2685
         _ExtentX        =   4736
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
         Picture         =   "CfrmSP.ctx":0510
         Caption         =   "Оценить эффективность"
         ButtonStyle     =   2
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "CfrmSP.ctx":0622
      End
   End
   Begin Threed.SSCommand ssCopy 
      Height          =   285
      Left            =   7335
      TabIndex        =   15
      ToolTipText     =   "Копировать меры в буфер обмена Windows"
      Top             =   2145
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   503
      _Version        =   131074
      PictureUseMask  =   -1  'True
      Caption         =   "Коп."
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCheck ssPercent 
      Height          =   210
      Left            =   6840
      TabIndex        =   14
      ToolTipText     =   "Показать соотношение на гистограмме"
      Top             =   2167
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   370
      _Version        =   131074
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "%"
   End
   Begin VB.Image imgExclamation 
      Height          =   270
      Left            =   1905
      Picture         =   "CfrmSP.ctx":0734
      Top             =   6135
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgDQ 
      Height          =   270
      Left            =   360
      Picture         =   "CfrmSP.ctx":088E
      Top             =   6180
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgFChange 
      Height          =   255
      Left            =   0
      Picture         =   "CfrmSP.ctx":09E8
      Top             =   6090
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgMoney 
      Height          =   270
      Left            =   735
      Picture         =   "CfrmSP.ctx":0B36
      Top             =   6150
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgC 
      Height          =   270
      Left            =   1455
      Picture         =   "CfrmSP.ctx":0C90
      Top             =   6135
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgT 
      Height          =   270
      Left            =   3060
      Picture         =   "CfrmSP.ctx":0DEA
      Top             =   6060
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   2655
      Picture         =   "CfrmSP.ctx":0F44
      Top             =   6150
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgH 
      Height          =   270
      Left            =   2265
      Picture         =   "CfrmSP.ctx":109E
      Top             =   6120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgSF 
      Height          =   270
      Left            =   1140
      Picture         =   "CfrmSP.ctx":11F8
      Top             =   6120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgKey 
      Height          =   270
      Left            =   3630
      Picture         =   "CfrmSP.ctx":1352
      Top             =   6090
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgComplSP 
      Height          =   240
      Left            =   4350
      Picture         =   "CfrmSP.ctx":14AC
      Top             =   6090
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "CfrmSP"
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
Private m_l_LastIdx As Long
Private m_s_Currname As String

Private m_mg_Base As MGertNet
Private m_csp_Tmp As collSF
Private m_csp_Tmp2 As collSF
Private m_d_P0 As Double

Private HighlightedNode As PVBranch
Private m_myFont As StdFont

Public Caption As String

Public Property Get P0() As Double
  P0 = m_d_P0
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

Public Property Get Ptv() As PVTreeViewX
  Set Ptv = tv
End Property


Public Property Get LastSP() As Long
  LastSP = m_l_LastIdx
End Property

Public Property Let LastSP(ByVal lSP As Long)
  m_l_LastIdx = lSP
End Property


Public Property Get MinW() As Single
  MinW = 7100
End Property

Public Property Get MinH() As Single
  MinH = 6100
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
    UpdateControls
    LoadCurrSP
  End If
End Function


Private Sub ssCopy_Click()

  If haApp.CurrSFColl = -1 Or m_l_LastIdx = -1 Then
    MsgBox "Нет текущего комплекса мер", vbOKOnly Or vbExclamation, "Нельзя скопировать"
    Exit Sub
  End If
      
  Dim bm As CBeam
  Set bm = New CBeam
  
  bm.Beam Me
  
  'mmmmmmmmmmmmm
  'haApp.CurrSFColl
  
  Dim cllSF As collSF
  Set cllSF = m_csf_Coll(m_l_LastIdx)
  
  Dim srtOld As CollSFSorting
  srtOld = cllSF.Sorting
  cllSF.Sorting = CurrSorting
  On Error GoTo ERR_H
  
  Dim cllSFS As CollSFSorting: cllSFS = CurrSorting
  
  Dim vSumm
  Dim sp As SafetyPrecaution, sTmp As String
  Dim l1 As Long: l1 = 0: vSumm = 1
  
  If ssPercent.Value Then
    For Each sp In cllSF
      vSumm = Abs(frmHisto.GetDta(sp, cllSFS))
      Exit For
    Next sp
    For Each sp In cllSF
      If l1 > 0 Then
        vSumm = vSumm + Abs(frmHisto.GetDta(sp, cllSFS))
      End If
      l1 = l1 + 1
    Next sp
    
    If vSumm <> 0 Then
      vSumm = 100# / vSumm
    Else
      vSumm = 1
    End If
  End If
  
  
  sTmp = ""
  For Each sp In cllSF
    Select Case cllSFS
    
      Case CSFS_DeltaQDescending
        If ssPercent.Value Then
          sTmp = sTmp & Sprintf("%25s %10.2f", , sp.Name, sp.DeltaQ * vSumm) & vbCrLf
        Else
          sTmp = sTmp & Sprintf("%25s %22.15f", , sp.Name, sp.DeltaQ) & vbCrLf
        End If
        
      Case CSFS_KDescending
        If ssPercent.Value Then
          sTmp = sTmp & Sprintf("%25s %10.2f", , sp.Name, sp.Ke * vSumm) & vbCrLf
        Else
          sTmp = sTmp & Sprintf("%25s %22.15f", , sp.Name, sp.Ke) & vbCrLf
        End If
        
      Case CSFS_PriceDescending, CSFS_Price
        'GetDta = sf.Cost
        If ssPercent.Value Then
          sTmp = sTmp & Sprintf("%25s %10.2f", , sp.Name, CDbl(sp.Cost * vSumm)) & vbCrLf
        Else
          sTmp = sTmp & Sprintf("%25s %10s", , sp.Name, FormatCurrency(sp.Cost, 2, vbTrue, vbFalse, vbTrue)) & vbCrLf
        End If
        
    End Select
    
  Next sp
  
  
  Clipboard.Clear
  Clipboard.SetText sTmp, vbCFText
  
  cllSF.Sorting = srtOld
  Exit Sub
ERR_H:
  cllSF.Sorting = srtOld
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ssPercent_Click(Value As Integer)
  LoadCurrSP False
End Sub

Private Sub UserControl_InitProperties()
  tv.StandardDefaultPicture = pvtNone
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
End Sub

Private Sub UserControl_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxSPMonitor.GetExtraWH sWExtr, sHExtr
  
  
  ssfrmRang.Move m_LPad, m_UPad, ScaleWidth - 2 * m_LPad - sWExtr
  VideoSoftElastic1.Width = ssfrmRang.Width - VideoSoftElastic1.Left - ssRange.Left
  VideoSoftElastic2.Width = VideoSoftElastic1.Width
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  'PVTime1.Move (VideoSoftElastic1.Width - PVTime1.Width) / 2, 50
  
  PVTime1.Move (VideoSoftElastic1.Width - 1.2 * PVTime1.Width - PVTimeEst.Width) / 2, 50
  PVTimeEst.Move PVTime1.Left + 1.2 * PVTime1.Width, PVTime1.Top
  
  pbRun.Move 60 + sW, PVTime1.Top + PVTime1.Height + 10, VideoSoftElastic1.Width - 2 * (70 + sW)
  pbRun2.Move 50 + sW, 100, VideoSoftElastic2.Width - 2 * (60 + sW), VideoSoftElastic2.Height - 200
  
  ssfrmParams.Move m_LPad, ssfrmRang.Top + ssfrmRang.Height + m_UPad
  optDQ.Move ssfrmParams.Left + ssfrmParams.Width + m_LPad, ssfrmParams.Top
  optCost.Move optDQ.Left + optDQ.Width + m_LPad / 2, optDQ.Top
  optK.Move optCost.Left + optCost.Width + m_LPad / 2, optCost.Top
  ssPercent.Move optK.Left + optK.Width + m_LPad / 2, optK.Top + (optK.Height - ssPercent.Height) / 2
  ssCopy.Move ssPercent.Left + ssPercent.Width + m_LPad / 2, ssPercent.Top + (ssPercent.Height - ssCopy.Height) / 2
  tv.Move optDQ.Left, optDQ.Top + optDQ.Height + m_UPad / 2, ScaleWidth - (ssfrmParams.Left + ssfrmParams.Width + m_LPad) - m_LPad - sWExtr, ScaleHeight - (optDQ.Top + optDQ.Height) - m_UPad - sHExtr - m_UPad / 2
End Sub


Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_csf_Coll Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_csf_Coll = Nothing
  Set m_csp_Tmp = Nothing
  Set m_csp_Tmp2 = Nothing
  Set m_mg_Base = Nothing
  UpdateControls
  Set HighlightedNode = Nothing
  PVTime1.Time = "0:00:00"
  PVTimeEst.Time = "0:00:00"
  If Not haApp Is Nothing Then haApp.CloseRang
  pbRun.Value = 0
  pbRun2.Value = 0
End Sub

Public Sub SetupNumbers()
  If Not haApp Is Nothing Then
   If haApp.HaveDoc Then
     SetupPVNumeric haApp.GertNetMain.NFormatPr, haApp.GertNetMain.NDigitsPr, pvnBaseP, pvnRealP, pvnCalcedP, pvnProfitCost
   End If
 End If
End Sub


Public Sub BindModel(ByVal coll As CollectionX)
  Set m_csf_Coll = coll
  SetupNumbers
  
  pvnCost.Value = haApp.GertNetMain.AverageDamage
  UpdateControls
  LoadCurrSP False
    
End Sub

Private Sub optCost_Click()
  LoadCurrSP False
End Sub

Private Sub optDQ_Click()
  LoadCurrSP False
End Sub

Private Sub optK_Click()
  LoadCurrSP False
End Sub

Private Function GetSFColl() As collSF
  If tv.Count < 1 Then
    Set GetSFColl = Nothing
    Exit Function
  End If

  Dim node As PVBranch
  Set node = tv.Branches.Get(pvtGetChild, 0).Get(pvtGetChild, 0)
  
  Dim collTmp As New collSF, collCurr As collSF, ips As IPersistStorage
  Set ips = collTmp: ips.InitNew Nothing
  
  m_s_Currname = tv.Branches.Get(pvtGetChild, 0).DataVariant
  Set collCurr = m_csf_Coll(m_s_Currname)
  
  While node.IsValid
    If node.CheckBoxState = pvtChkBoxStateChecked Then
      Dim ilk As ILongKey: Set ilk = collCurr(node.Data)
      collTmp.Add collCurr(node.Data), ilk.Get()
    End If
    Set node = node.Get(pvtGetNextSibling, 0)
  Wend
  
  collTmp.Sorting = Me.CurrSorting
  Set GetSFColl = collTmp
  
End Function

Private Sub pvnCost_LostFocusEvent()
  If Not haApp.GertNetMain Is Nothing Then
    If haApp.GertNetMain.AverageDamage <> pvnCost.Value Then
      haApp.GertNetMain.AverageDamage = pvnCost.Value
      If Not m_mg_Base Is Nothing Then _
        m_mg_Base.AverageDamage = pvnCost.Value
    End If
  End If
End Sub

Private Sub ssApply_Click()
      
  Dim collTmp As collSF
  Set collTmp = GetSFColl()
  
  If collTmp Is Nothing Then
    MsgBox "Нет выбранных мер", vbExclamation Or vbOKOnly, "Нельзя применить"
    Exit Sub
  End If
  If collTmp.Count = 0 Then
    MsgBox "Нет выбранных мер", vbExclamation Or vbOKOnly, "Нельзя применить"
    Exit Sub
  End If
  
  Dim collNC As collSF
  If Not collTmp.CheckCompatible(collNC) Then
    Dim lRes As Long, sNC As String, spTmp As SafetyPrecaution
    For Each spTmp In collNC
      sNC = sNC & spTmp.Name & "; "
    Next spTmp
    sNC = Left(sNC, Len(sNC) - 2)
    lRes = MsgBox("Меры [" & sNC & "] не совместимы." & vbCrLf & "Игнорировать это ?", vbExclamation Or vbYesNo, "Такая комбинация мер не рекомендуется")
    If lRes = vbNo Then Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  frmApply.AssData collTmp
  frmApply.Show vbModal, frmMain
  DoEvents
  
  If frmApply.ModalResult Then
     Dim bFlOver As Boolean
     haApp.GertNetMain.ApplySF collTmp, bFlOver
     If bFlOver Then
       MsgBox "В ходе применения мер произошёл выход значений некоторых факторов за допустимые пределы", vbInformation Or vbOKOnly, "информация"
     End If
     Dim bm As New CBeam
     bm.Beam Me
     DoEvents
     frmMEditor.UpdateValuesAll
     LoadCurrSP False
  End If
  
  frmApply.UnassData
  
  Exit Sub
ERR_H:
  frmApply.UnassData
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ssCalc_Click()

  If Not IsBoundModel Then
    MsgBox "Нет модели", vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If m_mg_Base Is Nothing Then
    MsgBox "Нет базовой модели", vbOKOnly, "Нельзя ранжировать"
    Exit Sub
  End If
    
  
  Set m_csp_Tmp = GetSFColl()
  If m_csp_Tmp Is Nothing Then
    MsgBox "Нет выбранных мер", vbExclamation Or vbOKOnly, "Нельзя применить"
    Exit Sub
  End If
  If m_csp_Tmp.Count = 0 Then
    MsgBox "Нет выбранных мер", vbExclamation Or vbOKOnly, "Нельзя применить"
    Exit Sub
  End If
  
  Dim collNC As collSF
  If Not m_csp_Tmp.CheckCompatible(collNC) Then
    Dim lRes As Long, sNC As String, spTmp As SafetyPrecaution
    For Each spTmp In collNC
      sNC = sNC & spTmp.Name & "; "
    Next spTmp
    sNC = Left(sNC, Len(sNC) - 2)
    lRes = MsgBox("Меры [" & sNC & "] не совместимы." & vbCrLf & "Игнорировать это ?", vbExclamation Or vbYesNo, "Такая комбинация мер не рекомендуется")
    If lRes = vbNo Then Exit Sub
  End If
        
  On Error GoTo ERR_H
  pbRun.Value = 0: pbRun2.Value = 0
  PVTime1.Time = "0:00:00"
  PVTimeEst.Time = "0:00:00"
  
  Dim icln As IClonable
  Set icln = m_csp_Tmp
  Set m_csp_Tmp2 = icln.Clone()
  
  haApp.VirtualApply m_mg_Base, m_csp_Tmp2, frmMon.GetCurPrior(), Me
  
  Timer1.Interval = 1000
  Timer1.Enabled = True
  ssRange.Enabled = False
  ssCalc.Enabled = False
  ssSetBase.Enabled = False
  frmCancel.blbRang.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub sscheckHisto_Click(Value As Integer)
  If sscheckHisto.Value = True Then
    If m_l_LastIdx = -1 Then
      sscheckHisto.Value = False
      MsgBox "Нет текущегокомплекса мер", vbOKOnly Or vbExclamation, "Ошибка"
      Exit Sub
    End If
    frmHisto.LoadColl m_csf_Coll(m_l_LastIdx), CurrSorting, m_csf_Coll.Key(m_l_LastIdx), ssPercent.Value = True
    frmHisto.Visible = True
  Else
    frmHisto.Visible = False
    frmHisto.Clear
  End If
End Sub

Private Sub ssRange_Click()
  If Not IsBoundModel Then
    MsgBox "Нет модели", vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  If m_mg_Base Is Nothing Then
    MsgBox "Нет базовой модели", vbOKOnly, "Нельзя ранжировать"
    Exit Sub
  End If
  
  'm_mg_Base.AverageDamage = pvnCost.Value
  'If haApp.GertNetMain.AverageDamage <> pvnCost.Value Then _
   ' haApp.GertNetMain.AverageDamage = pvnCost.Value
   
  pvnCost_LostFocusEvent
  
  On Error GoTo ERR_H
  pbRun.Value = 0: pbRun2.Value = 0
  PVTime1.Time = "0:00:00"
  PVTimeEst.Time = "0:00:00"
  
  Dim icln As IClonable
  Set icln = m_csf_Coll(m_l_LastIdx)
  Set m_csp_Tmp2 = icln.Clone()
  
  haApp.StartRang m_mg_Base, m_csp_Tmp2, frmMon.GetCurPrior(), Me
  
  Timer1.Interval = 1000
  Timer1.Enabled = True
  ssRange.Enabled = False
  ssCalc.Enabled = False
  ssSetBase.Enabled = False
  frmCancel.blbRang.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub Timer1_Timer()
  With PVTime1
    If .Second Mod 2 = 0 Then _
      PVTimeEst.Time = FormatDateTime(haApp.GN_Rang.RngCount, vbLongTime)
      
    .Second = .Second + 1
    If .Second >= 60 Then .Minute = .Minute + 1: .Second = 0
    If .Minute >= 60 Then .Hour = .Hour + 1: .Minute = 0
  End With
End Sub


Private Sub ssSetBase_Click()
'  Dim iCln As IClonable
'  Set iCln = frmMon.LastRun
'  Set m_mg_Base = iCln.Clone

  Set m_mg_Base = frmMon.LastRun
  
  Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double
  m_mg_Base.GetInfSampleK 73, lMin, lMax, dMx, dDx
  m_d_P0 = dMx
  pvnBaseP.ValueReal = dMx
  
  UpdateControls
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

Private Sub ssApply_LostFocus()
  HighLight ssApply, False
End Sub

Private Sub ssApply_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssApply, True
End Sub

Private Sub ssApply_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssApply, False
End Sub


Friend Sub UpdateControls()
  
'  If Not m_mg_Base Is Nothing Then
'    Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double, dMx2 As Double
'    m_mg_Base.GetInfSampleK 73, lMin, lMax, dMx, dDx
'    pvnBaseP.ValueReal = dMx
'  End If

  If haApp Is Nothing Then
    ssApply.Enabled = False
    ssRange.Enabled = False
    ssCalc.Enabled = False
    pvnCost.Enabled = False
    Exit Sub
  End If

  If frmMon Is Nothing Then
    ssSetBase.Enabled = False
  Else
    ssSetBase.Enabled = Not haApp.IsCalcRang And Not frmMon.LastRun Is Nothing
  End If
  
  ssApply.Enabled = haApp.CurrSFColl <> -1
  ssRange.Enabled = Not haApp.IsCalcRang And Not m_mg_Base Is Nothing And ssApply.Enabled
  ssCalc.Enabled = ssRange.Enabled
  pvnCost.Enabled = ssRange.Enabled
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

Friend Sub LoadCurrSP(Optional ByVal bCheck As Boolean = True)
  Dim bFlSrt As Boolean: bFlSrt = False
  If bCheck And m_l_LastIdx = haApp.CurrSFColl Then Exit Sub
  m_l_LastIdx = haApp.CurrSFColl
    
  pvnRealP.ValueReal = 0
  pvnCalcedP.ValueReal = 0
  pvnProfit.ValueReal = 0
  pvnProfitCost.ValueReal = 0
  
  On Error GoTo ERR_H
  tv.Redraw = False
  ClearTree
  
  If m_l_LastIdx = -1 Or Not Me.IsBoundModel Then
    UpdateControls
    tv.Redraw = True
    frmHisto.Clear
    Exit Sub
  End If
    
  
  Dim collSPrec As collSF
  Set collSPrec = m_csf_Coll(m_l_LastIdx)
  With tv.Branches.Get(pvtGetChild, 0)
    .Text = "Комплекс мер '" & m_csf_Coll.Key(m_l_LastIdx) & "'"
    .DataVariant = m_csf_Coll.Key(m_l_LastIdx)
    .Data = m_l_LastIdx
  End With
  
  
  Dim sf As SafetyPrecaution
  For Each sf In collSPrec
    With sf
      If .DeltaQ = 0 Then
        .Profit = 0
        .Ke = 0
      Else
        .Profit = .DeltaQ * pvnCost.Value
        If .Cost = 0 Then
          .Ke = 0
        Else
          .Ke = .DeltaQ * pvnCost.Value / .Cost
        End If
      End If
    End With
  Next sf
  
      
  Dim srtOld As CollSFSorting
  srtOld = collSPrec.Sorting
  bFlSrt = True
  collSPrec.Sorting = CurrSorting
  
  Dim bRanged As Boolean
  bRanged = False
  For Each sf In collSPrec
    If sf.DeltaQ <> 0 Then
      bRanged = True
      Exit For
    End If
  Next sf
  
  For Each sf In collSPrec
    AddToTree sf, bRanged
  Next sf
  
  collSPrec.Sorting = srtOld
  
  
  ExpandAll tv
  
  tv.Redraw = True
  
  If frmHisto.Visible Then
    If sscheckHisto.Value = True Then
      frmHisto.LoadColl m_csf_Coll(m_l_LastIdx), CurrSorting, m_csf_Coll.Key(m_l_LastIdx), ssPercent.Value = True
    Else
      'frmHisto.Clear
    End If
  End If
  
  Exit Sub
ERR_H:
  tv.Redraw = True
  If bFlSrt Then collSPrec.Sorting = srtOld
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub AddToTree(ByVal sf As SafetyPrecaution, ByVal bRanged As Boolean)
  Dim pvb As PVBranch
  Dim myFont As StdFont
  Set pvb = tv.Branches.Get(pvtGetChild, 0)
  Set pvb = pvb.Add(pvtPositionInOrder, 0, "Мера '" & sf.Name & "'")
  pvb.ForeColor = RGB(0, 200, 100)
  Set pvb.CustomItemPicture = imgSF.Picture
  pvb.CheckBoxType = pvtChkBoxRegular
  pvb.CheckBoxState = pvtChkBoxStateNotChecked
  Set myFont = tv.CreateFont
  myFont.Bold = True
  Set pvb.Font = myFont
  
  Dim ilk As ILongKey: Set ilk = sf
  pvb.Data = ilk.Get()
  pvb.DataVariant = sf.Name
  
  Dim pv2 As PVBranch
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "ID = " & CStr(ilk.Get()))
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgKey.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Затраты на внедрение = " & FormatCurrency(sf.Cost, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgMoney.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Дельта Q = " & IIf(bRanged, Format(sf.DeltaQ, haApp.FmtDbl), "<Не вычислено>"))
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgDQ.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Выгода = " & FormatCurrency(sf.Profit, 2, vbTrue, vbFalse, vbTrue))
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgMoney.Picture
  
  Dim sStr As String
  If bRanged Then
    If sf.Cost = 0 Then
      sStr = "не доступно (стоимость = 0)"
    Else
      sStr = Format(sf.Ke, haApp.FmtDbl2)
    End If
  Else
    sStr = "не вычислена дельта Q"
  End If
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Кэ (выгода/затраты) = " & sStr)
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgMoney.Picture
  
  Set pv2 = pvb.Add(pvtPositionInOrder, 0, "Воздействие на факторы опасности текущей базовой модели")
  pv2.ForeColor = RGB(0, 128, 128)
  Set pv2.CustomItemPicture = imgFChange.Picture
  
  Set pvb = pv2
  
  Dim collFac As collFac, mgTmp As MGertNet
  If m_mg_Base Is Nothing Then
    Set collFac = haApp.GertNetMain.Factors
  Else
    Set collFac = m_mg_Base.Factors
  End If
    
  
  Set mgTmp = IIf(m_mg_Base Is Nothing, haApp.GertNetMain, m_mg_Base)
  
  Dim fc As FChange, s1 As String, s2 As String
  For Each fc In sf.FChanges
    Dim f As Factor: Set f = collFac(fc.NameFactor)
    Dim lEnum As LingvoEnum
    Set lEnum = mgTmp.GetEnumForFactor(f)
    s1 = lEnum.Mark(f.Value)
    
    Dim iVal As Integer: iVal = f.Value + fc.Delta
    Dim bOverWrap As Boolean: bOverWrap = False
    If iVal < 0 Then iVal = 0: bOverWrap = True
    If iVal >= lEnum.Count Then iVal = lEnum.Count - 1: bOverWrap = True
    
    s2 = lEnum.Mark(iVal)
    If bOverWrap Then s2 = "(! " & s2 & " !)"
    
    Dim sFcDscr As String
    If fc.tch = TR_NoChange Then
      sFcDscr = ""
    Else
      sFcDscr = "; (" & StrFromTCH(fc.tch, False) & ") " & StrFromTrust(f.TrustLvl) & " ----> " & StrFromTrust(mgTmp.TstApplyTCH(f, fc))
    End If
    
    Set pv2 = pvb.Add(pvtPositionInOrder, 0, fc.NameFactor & " (" & TSign(fc.Delta) & CStr(Abs(fc.Delta)) & ")  " & s1 & " ----> " & s2 & sFcDscr & "  / '" & f.Name & "'")
    pv2.ForeColor = RGB(0, 0, 255)
    If Not bOverWrap Then
      Select Case UCase(Left(fc.NameFactor, 1))
        Case "H"
          Set pv2.CustomItemPicture = imgH.Picture
        Case "M"
          Set pv2.CustomItemPicture = imgM.Picture
        Case "C"
          Set pv2.CustomItemPicture = imgC.Picture
        Case "T"
          Set pv2.CustomItemPicture = imgT.Picture
      End Select
    Else
      Set pv2.CustomItemPicture = imgExclamation.Picture
    End If
  Next fc
End Sub

Private Sub ClearTree()
  Set HighlightedNode = Nothing
  If tv.Count < 1 Then
    Dim pvNodeMain As PVBranch
    Set pvNodeMain = tv.Branches.Add(pvtPositionInOrder, 0, "Комплекс мер")
    pvNodeMain.ForeColor = RGB(0, 0, 200)
    'pvNodeMain.StandardItemPicture = pvtpicFolders
    Set pvNodeMain.CustomItemPicture = imgComplSP.Picture
  Else
    tv.Branches.Get(pvtGetChild, 0).Clear
    tv.Branches.Get(pvtGetChild, 0).Text = "Комплекс мер"
  End If
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Dim myFont As StdFont
    
    If Not HighlightedNode Is Nothing Then
        Set HighlightedNode.Font = Nothing
    End If
    
    If Shift = 0 Then
        Set HighlightedNode = tv.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
        If HighlightedNode.IsValid And HighlightedNode.Level <> 1 Then
            'Set myFont = tv.CreateFont
            'myFont.Bold = True
            Set HighlightedNode.Font = m_myFont
        Else
          Set HighlightedNode = Nothing
        End If
    End If
End Sub

Public Sub ICallBack_EndOfWork(ByVal dt As Date, ByVal gn As MGertNet)
  If gn.UserProp = US_SP_Calc2 Then _
    TermCalc True, dt, gn
End Sub
Public Sub ICallBack_EndOfWork2(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc True, dt, gn
End Sub
Public Sub ICallBack_EndOfWork3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub


Public Sub ICallBack_StepOfWork(ByVal iPercent As Integer)
  pbRun2.Value = iPercent
End Sub
Public Sub ICallBack_StepOfWork2(ByVal iPercent As Integer)
  pbRun.Value = iPercent
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
  MsgBox sMsg, vbExclamation Or vbOKOnly, "Ошибка"
End Sub

Private Sub TermCalc(ByVal bOk As Boolean, ByVal dt As Date, Optional ByVal gn As MGertNet = Nothing)
  Timer1.Enabled = False
  frmCancel.blbRang.PulseStop
  PVTime1.Time = Format(dt, "h:mm:ss")
  
  CreateReport bOk, gn
  
  If Not bOk Then _
    MsgBox IIf(gn.UserProp = US_SP_Calc1, "Ранжировка мер прервана", "Оценочный прогон прерван"), vbInformation Or vbOKOnly, "Информация"
  
  Dim bEn As Boolean
  bEn = haApp.CurrSFColl <> -1
  
  ssRange.Enabled = bEn
  ssCalc.Enabled = bEn
  ssSetBase.Enabled = bEn
  Set m_csp_Tmp = Nothing
  Set m_csp_Tmp2 = Nothing
End Sub

Friend Sub CreateReport(ByVal bOk As Boolean, ByVal gn As MGertNet)
  On Error GoTo ERR_H
  
  Dim bm As New CBeam
  bm.Beam Me
  
  If bOk Then
    If gn.UserProp = US_SP_Calc1 Then
      SynchronizeColls m_l_LastIdx, m_csp_Tmp2
      LoadCurrSP False
    Else
      FillPFrom gn
    End If
  End If
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub FillPFrom(ByVal gn As MGertNet)
  Dim sf As SafetyPrecaution, dSummQ As Double
  Dim cSummCost As Currency, cSummProfit As Currency
  dSummQ = 0: cSummCost = 0: cSummProfit = 0
  For Each sf In m_csp_Tmp2
    dSummQ = dSummQ + sf.DeltaQ
    cSummCost = cSummCost + sf.Cost
    cSummProfit = cSummProfit + sf.Profit
  Next sf
    
  Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double
  Dim dMxKey As Double
  gn.GetInfSampleK 73, lMin, lMax, dMx, dDx
  
  dMxKey = dMx
  dMx = m_d_P0 - dMx
    
  With m_csp_Tmp2
    .Profit = dMx * pvnCost.Value
    .Cost = cSummCost
    .DeltaQ = dMx
    If cSummCost = 0 Then
      .Ke = 0
    Else
      .Ke = dMx * pvnCost.Value / cSummCost
    End If
  End With
  
  pvnRealP.ValueReal = dMx
  pvnCalcedP.ValueReal = dSummQ
  pvnProfit.Value = dMx * pvnCost.Value
  If cSummCost = 0 Then
    pvnProfitCost.Text = "<не доступно>"
  Else
    pvnProfitCost.ValueReal = dMx * pvnCost.Value / cSummCost
  End If
  
    
  If haApp.CreRepVApply Then
    Dim rep As New CRepVApply
    rep.InitReport gn, haApp, m_mg_Base, m_csp_Tmp2, cSummProfit, dSummQ, m_s_Currname, dMxKey, m_d_P0
    Dim ir As IRepItem
    Set ir = rep
    haApp.Rep2.Add ir
  
    UpdateGUI_Rep2 haApp
  End If
  
End Sub

Private Sub SynchronizeColls(ByVal lIdx As Long, ByVal cll As collSF)
  If lIdx < 1 Or lIdx > m_csf_Coll.Count Then
    MsgBox "Изменился текущий комплекс мер", vbExclamation Or vbOKOnly, "Нельзя синхронизировать комплексы мер"
    Exit Sub
  End If
  
  Dim cOrig As collSF
  Set cOrig = m_csf_Coll(lIdx)
  
  Dim sf As SafetyPrecaution, sf0 As SafetyPrecaution
  For Each sf In cll
    Dim ilk As ILongKey: Set ilk = sf
    On Error Resume Next
    Set sf0 = cOrig(ilk.Get())
    If Err.Number <> 0 Then
      MsgBox "Изменился текущий комплекс мер", vbExclamation Or vbOKOnly, "Нельзя синхронизировать комплексы мер"
      Exit Sub
    End If
    On Error GoTo 0
    
    With sf0
      .DeltaQ = sf.DeltaQ
      .Ke = sf.Ke
      .Profit = sf.Profit
    End With
  Next sf
  
  'cOrig.Cost
  
End Sub

Private Sub ssCopy_LostFocus()
  HighLight ssCopy, False
End Sub

Private Sub ssCopy_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssCopy, True
End Sub

Private Sub ssCopy_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight ssCopy, False
End Sub


Private Sub UserControl_Initialize()
  Caption = "Монитор комплексов мер"
  m_l_LastIdx = -1
End Sub

Private Sub UserControl_Terminate()
  HandsOffModel
  'Unload frmHisto
  Unload frmApply
End Sub
