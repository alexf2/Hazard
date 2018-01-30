VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{8F25C900-8346-11CF-AACD-444553540000}#6.0#0"; "pvtime.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Begin VB.Form frmMon 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Монитор модели"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin PVTreeView3Lib.PVTreeViewX tv 
      Height          =   1710
      Left            =   1335
      TabIndex        =   22
      Top             =   4800
      Width           =   3690
      _Version        =   393216
      _ExtentX        =   6509
      _ExtentY        =   3016
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
      Left            =   5385
      Top             =   585
   End
   Begin Threed.SSFrame ssfrmParams 
      Height          =   2460
      Left            =   315
      TabIndex        =   3
      Top             =   1620
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4339
      _Version        =   131074
      Font3D          =   1
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Параметры моделирования"
      ShadowStyle     =   1
      Begin PVNumericLib.PVNumeric pvnN 
         Height          =   345
         Left            =   495
         TabIndex        =   6
         Top             =   1530
         Width           =   1440
         _Version        =   393216
         _ExtentX        =   2540
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "4.000"
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
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   1
         ValueMax        =   100000000
         ValueSpinDelta  =   0
         ValueReal       =   4000
         DecimalMax      =   13
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnK 
         Height          =   345
         Left            =   480
         TabIndex        =   4
         Top             =   525
         Width           =   1440
         _Version        =   393216
         _ExtentX        =   2540
         _ExtentY        =   609
         _StockProps     =   253
         Text            =   "200"
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
         SpinButtons     =   2
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   1
         ValueMax        =   100000000
         ValueSpinDelta  =   0
         ValueReal       =   200
         DecimalMax      =   13
         ValidateLimit   =   76
         ChangeColor     =   -1  'True
      End
      Begin VB.ComboBox cbnPrior 
         BackColor       =   &H00FFFBF0&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmMon.frx":0000
         Left            =   2520
         List            =   "frmMon.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   525
         Width           =   2490
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1365
         Left            =   2505
         TabIndex        =   10
         ToolTipText     =   "Доступно только для имитационного метода расчёта"
         Top             =   945
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   2408
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "Затравка ГСЧ"
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvnRnd 
            Height          =   345
            Left            =   945
            TabIndex        =   13
            Top             =   870
            Width           =   1440
            _Version        =   393216
            _ExtentX        =   2540
            _ExtentY        =   609
            _StockProps     =   253
            Text            =   "1"
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
            ValueMin        =   1
            ValueMax        =   1000000
            ValueReal       =   1
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin Threed.SSOption ssoptVal 
            Height          =   225
            Left            =   165
            TabIndex        =   12
            Top             =   585
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            Caption         =   "Значение"
         End
         Begin Threed.SSOption ssoptAvto 
            Height          =   225
            Left            =   165
            TabIndex        =   11
            Top             =   285
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            Caption         =   "Таймер"
            Value           =   -1
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   615
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   344
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "K="
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   1590
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   344
         _Version        =   131074
         ForeColor       =   0
         BackColor       =   12632256
         Caption         =   "N="
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Приоритет вычислительной нити"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2490
         TabIndex        =   24
         Top             =   255
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Число инициаций модели в одном опыте"
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   480
         TabIndex        =   9
         Top             =   1050
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Число опытов"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   270
         Width           =   1080
      End
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   930
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   4650
      _Version        =   327680
      _ExtentX        =   8202
      _ExtentY        =   1640
      _StockProps     =   70
      ConvInfo        =   1418783674
      BackColor       =   0
      BevelOuter      =   5
      Begin PVProgressBarLib.PVProgressBar pbRun 
         Height          =   435
         Left            =   60
         TabIndex        =   2
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
         Left            =   1905
         TabIndex        =   1
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
   End
   Begin Threed.SSFrame ssfrmReport 
      Height          =   2595
      Left            =   5790
      TabIndex        =   15
      Top             =   1605
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4577
      _Version        =   131074
      Font3D          =   1
      BackColor       =   12632256
      Caption         =   "Результаты моделирования"
      ShadowStyle     =   1
      Begin PVNumericLib.PVNumeric pvnResult1 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   555
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
      End
      Begin PVNumericLib.PVNumeric PVNumeric1 
         Height          =   315
         Left            =   225
         TabIndex        =   19
         Top             =   1290
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
      End
      Begin PVNumericLib.PVNumeric PVNumeric2 
         Height          =   315
         Left            =   225
         TabIndex        =   21
         Top             =   2025
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
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Разность вероятностей"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Вероятность происшествия(пред.)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   1050
         Width           =   2610
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Вероятность происшествия"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         TabIndex        =   17
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Image imgC 
      Height          =   270
      Left            =   675
      Picture         =   "frmMon.frx":0047
      Top             =   5460
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgT 
      Height          =   270
      Left            =   150
      Picture         =   "frmMon.frx":01A1
      Top             =   5460
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   675
      Picture         =   "frmMon.frx":02FB
      Top             =   4950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgH 
      Height          =   270
      Left            =   225
      Picture         =   "frmMon.frx":0455
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin Threed.SSCommand ssStart 
      Default         =   -1  'True
      Height          =   510
      Left            =   315
      TabIndex        =   14
      Top             =   4245
      Width           =   5340
      _ExtentX        =   9419
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
      Picture         =   "frmMon.frx":05AF
      Caption         =   "Запуск моделирования"
      PictureAlignment=   9
      PictureDisabledFrames=   1
      PictureDisabled =   "frmMon.frx":06C1
   End
End
Attribute VB_Name = "frmMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ICallBack

Private Const m_def_LPad = 100
Private Const m_def_UPad = 70

Private m_mg_GertNet As MGertNet
Private m_mg_LastRun As MGertNet

Private HighlightedNode As PVBranch

Friend Property Get LastRun() As MGertNet
  Set LastRun = m_mg_LastRun
End Property

Public Property Get MinW() As Single
  MinW = 7000
End Property

Public Property Get MinH() As Single
  MinH = 6800
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
'    Dim arr(0) As Control
'    Set arr(0) = ssStart
'    EnblControls2 arr, Not haApp.IsCalcRun
     If frmMEditor.ssoptAnalytic.Value = True Or frmMEditor.ssoptAnalytic2.Value = True Then
       ssoptAvto.Enabled = False
       ssoptVal.Enabled = False
       pvnRnd.Enabled = False
    Else
      ssoptAvto.Enabled = True
      ssoptVal.Enabled = True
      pvnRnd.Enabled = IIf(ssoptAvto.Value = True, False, True)
    End If
  End If
End Function


Private Sub cbnPrior_Click()
  If Not haApp Is Nothing Then
    If Not haApp.GN_Run Is Nothing Then _
      haApp.GN_Run.CalcSpeed = GetCurPrior()
  End If
End Sub



Private Sub Form_Load()
  cbnPrior.ListIndex = 3
End Sub

Friend Function GetCurPrior() As GERTNETLib.ThrdPriority
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


Private Sub Form_Resize()
  Dim sWExtr As Single, sHExtr As Single
  frxModelMon.GetExtraWH sWExtr, sHExtr
  
  VideoSoftElastic1.Move m_def_LPad, m_def_UPad, ScaleWidth - 2 * m_def_LPad - sWExtr
  ssfrmParams.Move VideoSoftElastic1.Left + (VideoSoftElastic1.Width - ssfrmParams.Width) / 2, VideoSoftElastic1.Top + VideoSoftElastic1.Height + m_def_UPad
  ssStart.Move VideoSoftElastic1.Left, ssfrmParams.Top + ssfrmParams.Height + m_def_UPad, VideoSoftElastic1.Width
  ssfrmReport.Move VideoSoftElastic1.Left, ssStart.Top + ssStart.Height + m_def_UPad
    
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  
  PVTime1.Move (VideoSoftElastic1.Width - PVTime1.Width) / 2, 50
  pbRun.Move 60 + sW, PVTime1.Top + PVTime1.Height + 10, VideoSoftElastic1.Width - 2 * (70 + sW)
  
  tv.Move ssfrmReport.Left + ssfrmReport.Width + m_def_LPad, ssfrmReport.Top, VideoSoftElastic1.Width - ssfrmReport.Width - m_def_LPad, ssfrmReport.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set HighlightedNode = Nothing
  HandsOffModel
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_mg_GertNet = Nothing
  Set m_mg_LastRun = Nothing
  
  haApp.CloseCalc
  pbRun.Value = 0
  PVTime1.Time = "0:00:00"
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
  pvnResult1.ValueInteger = 0
  PVNumeric1.ValueInteger = 0
  PVNumeric2.ValueInteger = 0
  TvClear
End Sub

Private Sub pvnK_Decrement(ByVal DecrementValue As Double)
  ChkID pvnK, True
End Sub

Friend Sub ChkID(ByVal pv As PVNumeric, ByVal bDec As Boolean)
  Dim l As Currency
  l = 1
  Do While l <= pv.ValueCurrency
   l = l * 10
  Loop
  l = CDbl(l) / 100# * 10#
  If bDec Then
    pv.ValueCurrency = pv.ValueCurrency - l
  Else
    pv.ValueCurrency = pv.ValueCurrency + l
  End If
End Sub

Private Sub pvnK_Increment(ByVal IncrementValue As Double)
  ChkID pvnK, False
End Sub

Private Sub pvnN_Decrement(ByVal DecrementValue As Double)
  ChkID pvnN, True
End Sub

Private Sub pvnN_Increment(ByVal IncrementValue As Double)
  ChkID pvnN, False
End Sub

Private Sub ssoptAvto_Click(Value As Integer)
  If ssoptAvto.Value = True Then pvnRnd.Enabled = False
End Sub

Private Sub ssoptVal_Click(Value As Integer)
  If ssoptVal.Value = True Then pvnRnd.Enabled = True
End Sub

Private Sub ssStart_Click()
  If Not IsBoundModel Then
    MsgBox "Нет модели", vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  pbRun.Value = 0
  PVTime1.Time = "0:00:00"
  haApp.StartCalc CLng(pvnK.ValueReal), CLng(pvnN.ValueReal), IIf(ssoptAvto.Value = True, -1, CLng(pvnRnd.ValueCurrency)), GetCurPrior(), Me
  Timer1.Interval = 1000
  Timer1.Enabled = True
  ssStart.Enabled = False
  frmCalibrate.ssStart.Enabled = False
  frmCancel.blbRunOnce.PulseStart
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
End Sub

Private Sub Timer1_Timer()
  With PVTime1
    .Second = .Second + 1
    If .Second >= 60 Then .Minute = .Minute + 1: .Second = 0
    If .Minute >= 60 Then .Hour = .Hour + 1: .Minute = 0
  End With
End Sub

Public Sub ICallBack_EndOfWork(ByVal dt As Date, ByVal gn As MGertNet)
  TermCalc True, dt, gn
End Sub
Public Sub ICallBack_EndOfWork2(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub ICallBack_EndOfWork3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub


Public Sub ICallBack_StepOfWork(ByVal iPercent As Integer)
  pbRun.Value = iPercent
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
  MsgBox sMsg, vbExclamation Or vbOKOnly, "Ошибка"
End Sub

Private Sub TermCalc(ByVal bOk As Boolean, ByVal dt As Date, Optional ByVal gn As MGertNet = Nothing)
  Timer1.Enabled = False
  frmCancel.blbRunOnce.PulseStop
  PVTime1.Time = Format(dt, "h:mm:ss")
  
  CreateReport bOk, gn
  If bOk Then Set m_mg_LastRun = gn: frmSP.UpdateControls: frmOptimiz.UpdateControls
  'haApp.CloseCalc
  
  If Not bOk Then _
    MsgBox "Прогон модели прерван", vbInformation Or vbOKOnly, "Информация"
  
  ssStart.Enabled = True
  frmCalibrate.ssStart.Enabled = True
End Sub

Friend Sub CreateReport(ByVal bOk As Boolean, ByVal gn As MGertNet)
  On Error GoTo ERR_H
  
  Dim bm As New CBeam
  bm.Beam Me
  
  If bOk Then
    Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double, dMx2 As Double
    gn.GetInfSampleK 73, lMin, lMax, dMx, dDx
    pvnResult1.ValueReal = dMx
    
    If Not m_mg_LastRun Is Nothing Then
      m_mg_LastRun.GetInfSampleK 73, lMin, lMax, dMx2, dDx
      PVNumeric1.ValueReal = dMx2
      PVNumeric2.ValueReal = dMx - dMx2
      
      TvClear
      
      Dim cFac As Factor, cFacOld As Factor, ibsKey As IBSTRKey
      Dim bFlAdd As Boolean
      bFlAdd = False
      For Each cFac In gn.Factors
        Set ibsKey = cFac
        Set cFacOld = m_mg_LastRun.Factors(ibsKey.Get())
        If cFac.Value <> cFacOld.Value Or cFac.TrustLvl <> cFacOld.TrustLvl Then AddFactor cFacOld, cFac: bFlAdd = True
      Next cFac
      If bFlAdd Then ExpandAll tv
    End If
    
          
    If haApp.CreRepRun Then
      Dim rep As New CRepRun
      rep.InitReport gn, haApp, m_mg_LastRun
      Dim ir As IRepItem
      Set ir = rep
      haApp.Rep1.Add ir
    
      UpdateGUI_Rep1 haApp
    End If
    
  End If
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
End Sub


Private Sub TvClear()
  If tv.Count < 1 Then
    Dim pvNodeMain As PVBranch, pvNodeH As PVBranch, pvNodeM As PVBranch, pvNodeT As PVBranch, pvNodeC As PVBranch
    Set pvNodeMain = tv.Branches.Add(pvtPositionInOrder, 0, "Модифицированные факторы")
    pvNodeMain.ForeColor = RGB(237, 232, 83)
    pvNodeMain.StandardItemPicture = pvtpicFolders
    
    Set pvNodeH = pvNodeMain.Add(pvtPositionInOrder, 0, "Человек")
    pvNodeH.ForeColor = RGB(14, 239, 61)
    Set pvNodeH.CustomItemPicture = imgH.Picture
    
    Set pvNodeM = pvNodeMain.Add(pvtPositionInOrder, 1, "Машина")
    pvNodeM.ForeColor = RGB(14, 239, 61)
    Set pvNodeM.CustomItemPicture = imgM.Picture
    
    Set pvNodeT = pvNodeMain.Add(pvtPositionInOrder, 2, "Технология")
    pvNodeT.ForeColor = RGB(14, 239, 61)
    Set pvNodeT.CustomItemPicture = imgT.Picture
    
    Set pvNodeC = pvNodeMain.Add(pvtPositionInOrder, 3, "Среда")
    pvNodeC.ForeColor = RGB(14, 239, 61)
    Set pvNodeC.CustomItemPicture = imgC.Picture
    
    ExpandAll tv
  Else
    Set pvNodeMain = tv.Branches.Get(pvtGetChild, 0)
    Do While pvNodeMain.IsValid
      Set pvNodeH = pvNodeMain.Get(pvtGetChild, 0)
      If pvNodeH.IsValid Then pvNodeH.Clear
      Set pvNodeMain = pvNodeMain.Get(pvtGetNextSibling, 0)
    Loop
  End If
End Sub


Private Function FIdx(ByVal f As Factor)
  Dim ibsKey As IBSTRKey: Set ibsKey = f
  Select Case Left(ibsKey.Get(), 1)
    Case "H"
      FIdx = 0
    Case "M"
      FIdx = 1
    Case "T"
      FIdx = 2
    Case "C"
      FIdx = 3
  End Select
  
End Function

Private Sub AddFactor(ByVal fOld As Factor, ByVal fNew As Factor)
  Dim pvB As PVBranch
  Set pvB = tv.Branches.Get(pvtGetChild, 0).Get(pvtGetChild, FIdx(fOld))
  Dim hh As Boolean
  hh = pvB.IsValid
  Dim ibsKey As IBSTRKey: Set ibsKey = fOld
  
  Dim sDescrOld As String, sTrustOld As String, sTrust As String
  Dim dValOld As Double, dVal As Double, sDescr As String
  
  m_mg_LastRun.GetFactorPresent fOld, sDescrOld, sTrustOld, dValOld
  m_mg_GertNet.GetFactorPresent fNew, sDescr, sTrust, dVal
  
  Dim sRes As String, sSub As Integer
  sSub = fNew.Value - fOld.Value
  sRes = ibsKey.Get() & ":  "
  If dValOld <> dVal Then _
    sRes = sRes & sDescrOld & "---(" & TSign(sSub) & CStr(Abs(sSub)) & ")--->" & sDescr
    
  sSub = CInt(fNew.TrustLvl) - CInt(fOld.TrustLvl)
  If fOld.TrustLvl <> fNew.TrustLvl Then _
    sRes = sRes & IIf(dValOld <> dVal, ";", "") & " КД:[" & sTrustOld & "---(" & TSign(sSub) & CStr(Abs(sSub)) & ")--->" & sTrust & "]"
    
  sRes = sRes & "  /" & Chr$(147) & fNew.Name & Chr$(148)
  
  Set pvB = pvB.Add(pvtPositionInOrder, 0, sRes)
  pvB.ForeColor = RGB(108, 245, 250)
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
        End If
    End If
End Sub
