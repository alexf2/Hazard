VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#6.0#0"; "pvprgbar.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmTM 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Числовые характеристики модели"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frmTM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   6195
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   10927
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin Hazard.CHisto Graph1 
         Height          =   3630
         Left            =   3075
         TabIndex        =   13
         Top             =   810
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   6403
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontLbl {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox List1 
         BackColor       =   &H00FFFBF0&
         Height          =   2550
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   825
         Width           =   2805
      End
      Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
         Height          =   570
         Left            =   1665
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   105
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
            TabIndex        =   16
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
      Begin Threed.SSFrame ssfAlho 
         Height          =   2025
         Left            =   105
         TabIndex        =   17
         ToolTipText     =   "Алгоритм, используемый для прогонов модели"
         Top             =   3465
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   3572
         _Version        =   131074
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   12632256
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Срез"
         ShadowStyle     =   1
         Begin PVNumericLib.PVNumeric pvn1 
            Height          =   315
            Left            =   675
            TabIndex        =   3
            ToolTipText     =   "Левая доля отображаемой части выборки (1- отобразить всё, 0.5 - половину)"
            Top             =   285
            Width           =   810
            _Version        =   393216
            _ExtentX        =   1429
            _ExtentY        =   556
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
            ValueReal       =   1
            DecimalMax      =   16
            ValidateLimit   =   76
            ChangeColor     =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   600
            Left            =   1575
            TabIndex        =   19
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   1058
            _Version        =   131074
            BackStyle       =   1
            BevelOuter      =   0
            Begin Threed.SSOption ssoVent 
               Height          =   285
               Left            =   0
               TabIndex        =   21
               ToolTipText     =   "Собрать статистику по кранам (на входе и на правом выходе)"
               Top             =   0
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   503
               _Version        =   131074
               ForeColor       =   0
               BackColor       =   12632256
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Краны"
               Value           =   -1
            End
            Begin Threed.SSOption ssoSt 
               Height          =   285
               Left            =   0
               TabIndex        =   20
               ToolTipText     =   "Собрать статистику по состояниям"
               Top             =   285
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   131074
               ForeColor       =   0
               BackColor       =   12632256
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Состояния"
            End
         End
         Begin Threed.SSOption ssopt33 
            Height          =   285
            Left            =   1950
            TabIndex        =   10
            Top             =   885
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "/III"
         End
         Begin Threed.SSOption ssopt44 
            Height          =   285
            Left            =   1950
            TabIndex        =   11
            Top             =   1200
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "/IV"
         End
         Begin Threed.SSOption ssopt11 
            Height          =   285
            Left            =   720
            TabIndex        =   6
            Top             =   885
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "/I"
         End
         Begin Threed.SSOption ssopt22 
            Height          =   285
            Left            =   720
            TabIndex        =   7
            Top             =   1200
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "/II"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Вид:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   195
            TabIndex        =   18
            Top             =   345
            Width           =   420
            WordWrap        =   -1  'True
         End
         Begin Threed.SSCommand ssCopy 
            Default         =   -1  'True
            Height          =   375
            Left            =   675
            TabIndex        =   12
            ToolTipText     =   "Копирование выборки в буфер обмена Windows"
            Top             =   1545
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   661
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmTM.frx":000C
            Caption         =   "Копировать"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSOption ssopt4 
            Height          =   285
            Left            =   1425
            TabIndex        =   9
            Top             =   1200
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "IV"
         End
         Begin Threed.SSOption ssopt3 
            Height          =   285
            Left            =   1425
            TabIndex        =   8
            Top             =   885
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "III"
         End
         Begin Threed.SSOption ssopt2 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   1200
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "II"
         End
         Begin Threed.SSOption ssopt1 
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   885
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   131074
            ForeColor       =   0
            BackColor       =   12632256
            BackStyle       =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "I"
            Value           =   -1
         End
      End
      Begin Threed.SSCommand ssStart 
         Height          =   510
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Извлечь числовые характеристики модели"
         Top             =   105
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   900
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
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
         Picture         =   "frmTM.frx":016A
         Caption         =   "Выполнить"
         ButtonStyle     =   2
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "frmTM.frx":027C
      End
      Begin Threed.SSCommand ssOk 
         Height          =   525
         Left            =   3555
         TabIndex        =   14
         ToolTipText     =   "Выполняет модификации и закрывает диалог"
         Top             =   5445
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmTM.frx":038E
         Caption         =   "Закрыть"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400
Private Const LST_WID = (1# / 4#)


Private m_b_LockResize As Boolean
Private m_b_Res As Boolean
Private WithEvents m_mg As MGertNet
Attribute m_mg.VB_VarHelpID = -1

Public Property Get CanClose() As Boolean
  If m_mg Is Nothing Then
    CanClose = True
  Else
    CanClose = IIf(m_mg.IsRunning, False, True)
  End If
End Property

Private Sub Form_Initialize()
  m_b_LockResize = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssOk_Click
  ElseIf KeyAscii = vbKeyReturn Then
    'ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    'ssOk_Click
    If m_mg.IsRunning Then
      Cancel = 1
      Exit Sub
    End If
    
    ssOk_Click
    
  End If
End Sub

Friend Sub MClose()
  PostMessage hWnd, WM_CLOSE, 0, 0
End Sub


Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 5000 Then Width = 5000
    If Height < 5000 Then Height = 5000
  Else
    m_b_LockResize = False
    Exit Sub
  
  End If
  
  On Error GoTo ERR_H
  
  Dim lBW As Long, lBH As Long
  lBW = ScaleX(2 * GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips)
  lBH = ScaleY(2 * GetSystemMetrics(SM_CYBORDER), vbPixels, vbTwips)
  
  SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
  
  ssStart.Move L_PAD, T_PAD
  VideoSoftElastic1.Move ssStart.Left + ssStart.Width + L_PAD, ssStart.Top + (ssStart.Height - VideoSoftElastic1.Height) / 2, ScaleWidth - ssStart.Width - 3 * L_PAD
  
  Dim sW As Single: sW = ScaleX(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
  Dim sH As Single: sH = ScaleY(VideoSoftElastic1.BevelOuterWidth, vbPixels, vbTwips)
    
  pbRun.Move 50 + sW, 100, VideoSoftElastic1.Width - 2 * (60 + sW), VideoSoftElastic1.Height - 200
  
  List1.Move L_PAD, ssStart.Top + ssStart.Height + T_PAD, List1.Width, ScaleHeight - ssStart.Height - 4 * T_PAD - ssfAlho.Height
  Graph1.Move List1.Left + List1.Width + L_PAD, List1.Top, ScaleWidth - 3 * L_PAD - List1.Width, ScaleHeight - VideoSoftElastic1.Height - 4 * T_PAD - ssOk.Height
  ssfAlho.Move List1.Left + (List1.Width - ssfAlho.Width) / 2, List1.Top + List1.Height + T_PAD
  ssOk.Move Graph1.Left + (Graph1.Width - ssOk.Width) / 2, Graph1.Top + Graph1.Height + (ScaleHeight - Graph1.Top - Graph1.Height - ssOk.Height) / 2
    
    
  Dim r As RECT
  r.Left = 0: r.Top = 0
  r.Right = ScaleX(ScaleWidth, vbTwips, vbPixels)
  
  r.Bottom = ScaleY(ScaleHeight, vbTwips, vbPixels)
  ValidateRect hWnd, r
  InvalidateRect SSPanel1.hWnd, r, 0
  
  Dim cc1 As Control
  For Each cc1 In Me.Controls
    With cc1
      If .Name <> "SSPanel1" And .Name <> "ssCopy" Then
        If TypeOf cc1 Is Label Then GoTo L_NEXT_CC
        If TypeOf cc1 Is SSOption Then GoTo L_NEXT_CC
        'If TypeOf cc1 Is Timer Then GoTo L_NEXT_CC
        If TypeOf cc1 Is PVProgressBar Then GoTo L_NEXT_CC
        If TypeOf cc1 Is PVNumeric Then GoTo L_NEXT_CC
'        On Error Resume Next
'        If Not cc1.Container Is Me Then GoTo L_NEXT_CC
'        On Error GoTo ERR_H
        
        r.Left = ScaleX(.Left, vbTwips, vbPixels)
        r.Top = ScaleY(.Top, vbTwips, vbPixels)
        r.Right = r.Left + ScaleX(.Width, vbTwips, vbPixels)
        r.Bottom = r.Top + ScaleY(.Height, vbTwips, vbPixels)
        ValidateRect SSPanel1.hWnd, r
        
        r.Left = 0: r.Top = 0
        r.Right = ScaleX(.Width, vbTwips, vbPixels)
        r.Bottom = ScaleY(.Height, vbTwips, vbPixels)
        'On Error Resume Next
        InvalidateRect .hWnd, r, 1
        'On Error GoTo ERR_H
      End If
    End With
L_NEXT_CC:
  Next cc1
    
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_mg = Nothing
End Sub

Private Sub Graph1_Click()
'kkkkkkkk
  If Not ssOk.Enabled Then Exit Sub
  
'  Dim ns As Tag_TypANodesStat, arr() As Double
'  m_mg.GetModelStat CurrCut, ns, arr, True
'
'  Dim lSz As Long, dSumm As Double, dTmp As Double
'  dSumm = 0
'  For lSz = LBound(arr) To UBound(arr) Step 4
'    dTmp = arr(lSz)
'  Next lSz
End Sub

Private Sub m_mg_OnCancel(ByVal dtTime As Date)
  TermCalc False, dtTime
End Sub

Private Sub m_mg_OnEndOfWork(ByVal dtTime As Date)
  TermCalc True, dtTime
End Sub

Private Sub m_mg_OnErrorCalc(ByVal bsErrMsg As String)
  Dim dt As Date
  dt = m_mg.TimeWork
  TermCalc False, dt
  MsgBox bsErrMsg, vbExclamation Or vbOKOnly, "Ошибка"
End Sub

Private Sub m_mg_OnStepOfWork(ByVal shPercent As Integer)
  pbRun.Value = shPercent
End Sub

Private Sub TermCalc(ByVal bOk As Boolean, ByVal dt As Date)

  Dim pTmp As StdPicture
  Set pTmp = ssStart.PictureDisabled
  Set ssStart.PictureDisabled = ssStart.Picture
  Set ssStart.Picture = pTmp
  
  If bOk Then
    CreateReport
    ssCopy.Enabled = True
    ssopt1.Enabled = True
    ssopt2.Enabled = True
    ssopt3.Enabled = True
    ssopt4.Enabled = True
  Else
    MsgBox "Прогон модели прерван", vbInformation Or vbOKOnly, "Информация"
  End If
  ssOk.Enabled = True
End Sub

Private Property Get CurrCut() As Long
  If m_mg.RunMode = MT_Analytical Or m_mg.RunMode = MT_AnalyticalIndistinct Then
    If ssopt1.Value = True Then
      CurrCut = 0
    ElseIf ssopt2.Value = True Then
      CurrCut = 1
    ElseIf ssopt3.Value = True Then
      CurrCut = 2
    Else
      CurrCut = 3
    End If
    CurrCut = 2 * CurrCut
  Else
    If ssopt1.Value = True Then
      CurrCut = 0
    ElseIf ssopt11.Value = True Then
      CurrCut = 1
    ElseIf ssopt2.Value = True Then
      CurrCut = 2
    ElseIf ssopt22.Value = True Then
      CurrCut = 3
    ElseIf ssopt3.Value = True Then
      CurrCut = 4
    ElseIf ssopt33.Value = True Then
      CurrCut = 5
    ElseIf ssopt4.Value = True Then
      CurrCut = 6
    ElseIf ssopt44.Value = True Then
      CurrCut = 7
    End If
  End If
End Property

Private Sub pvn1_lostFocus()
  If ssopt1.Enabled Then FillHisto
End Sub

Private Sub ssCopy_Click()

  Dim bm As CBeam
  Set bm = New CBeam
  
  bm.Beam Me
  
  Dim ns As Tag_TypANodesStat, arr() As Double
  m_mg.GetModelStat CurrCut, ns, arr, True
  Dim sStr As String
  
  Dim l1 As Long, l2 As Long
  l2 = UBound(arr)
  For l1 = LBound(arr) To l2 Step 4
    If m_mg.RunMode = MT_Analytical Then
      sStr = sStr & Sprintf("%15.12g", , arr(l1 + 1)) & "   " & Sprintf("%25.15f", , arr(l1)) & vbCrLf
    Else
      sStr = sStr & Sprintf("%15.12g", , arr(l1 + 1)) & "   " & Sprintf("%15.12g", , arr(l1 + 2)) & "       " & Sprintf("%25.15f", , arr(l1)) & vbCrLf
    End If
  Next l1
  
  Clipboard.Clear
  Clipboard.SetText sStr, vbCFText
End Sub

Private Sub ssCopy_LostFocus()
  HighLight2 ssCopy, False, Me.hWnd
End Sub

Private Sub ssCopy_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCopy, True, Me.hWnd
End Sub

Private Sub ssCopy_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCopy, False, Me.hWnd
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
  If Not ssOk.Enabled Then Exit Sub
  
  m_b_Res = True
  Set m_mg = Nothing
  'MClose
  Unload Me
End Sub


Friend Sub AssData(ByVal m As MGertNet)
  m_b_Res = False
  Dim icln As IClonable
  Set icln = m
  Set m_mg = icln.Clone()
  
  'm_mg.RunMode = MT_Analytical
  m_mg.StatOn = TSF_Full
  'm_mg.StatIn = True
  
  If m_mg.RunMode = MT_Analytical Or m_mg.RunMode = MT_AnalyticalIndistinct Then
    ssopt1.Enabled = True: ssopt11.Enabled = False
    ssopt2.Enabled = True: ssopt22.Enabled = False
    ssopt3.Enabled = True: ssopt33.Enabled = False
    ssopt4.Enabled = True: ssopt44.Enabled = False
  Else
    ssopt1.Enabled = True: ssopt11.Enabled = True
    ssopt2.Enabled = True: ssopt22.Enabled = True
    ssopt3.Enabled = True: ssopt33.Enabled = True
    ssopt4.Enabled = True: ssopt44.Enabled = True
  End If
  
End Sub

Friend Sub UnassData()
  Set m_mg = Nothing
End Sub


Private Sub ssopt1_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt11_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt2_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt22_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt3_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt33_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt4_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssopt44_Click(Value As Integer)
  FillHisto
End Sub

Private Sub ssStart_Click()
  
  On Error GoTo ERR_H
  
  If m_mg.IsRunning Then
    m_mg.Cancel
    Exit Sub
  End If
  
  pbRun.Value = 0
     
  m_mg.StatIn = IIf(ssoVent.Value = True, True, False)
  m_mg.CalcSpeed = frmMon.GetCurPrior()
  m_mg.Run frmMon.ValK, frmMon.ValN, False, frmMon.RndVal
    


  ssOk.Enabled = False
  ssCopy.Enabled = False
  Dim pTmp As StdPicture
  Set pTmp = ssStart.Picture
  Set ssStart.Picture = ssStart.PictureDisabled
  Set ssStart.PictureDisabled = pTmp
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub FillHisto()
  'Dim clrOld As ColorDataConstants
  
  If pvn1.ValueReal < 0.1 Then
    MsgBox "Доля выборки должна быть не менее 0.1", vbOKOnly Or vbExclamation, "Ошибка"
    Exit Sub
  End If
  
  Dim ns As Tag_TypANodesStat, arr() As Double
  
  m_mg.GetModelStat CurrCut, ns, arr, True
      
  Graph1.Trunc = pvn1.ValueReal
  Graph1.SetArrData arr
  
'  With Graph1
'    .DrawMode = GraphLib.DrawModeConstants.gphNoAction
'
'    .AutoInc = AutoIncConstants.gphOff
'    .RandomData = RandomDataConstants.gphOff
'    '.DataReset = gphGraphData
'    .DataReset = gphAllData
'    .NumSets = 1
'    .ThisSet = 1
'    .NumPoints = (UBound(arr) - LBound(arr) + 1) / 4
'    '.TickEvery = 2
'
'    '.Background = gphBlue
'
'    .ThisSet = 1
'
'    .ThisPoint = 1
'    Dim lPtCnt As Long, sf As SafetyPrecaution
'    lPtCnt = 1
'
'
'    On Error GoTo ERR_H
'
'
'    Dim lIdx As Long
'    For lIdx = LBound(arr) To UBound(arr) Step 4
'
'      .ThisPoint = lPtCnt
'
'      '.XPosData = arr(lIdx + 1)
'
'      .GraphData = arr(lIdx)
'
'      '.XPosData = arr(lIdx + 1)
'
'
'      Dim clrNew As ColorDataConstants
'      Do
'        clrNew = RndClr()
'      Loop While clrNew = clrOld
'
'      .ColorData = clrNew
'      clrOld = clrNew
'
'      '.LegendText = sf.Name
'
'      '.LabelText = CStr(arr(lIdx + 1))
'
'
'      lPtCnt = lPtCnt + 1
'    Next lIdx
'
'    .PrintStyle = PrintStyleConstants.gphColor
'    .GridStyle = GridStyleConstants.gphBoth
'
'    '.Labels = LabelsConstants.gphOn
'
'    .ToolTipText = "Ось X - индекс; ось Y - вероятность появления"
'
'    .FontUse = gphOtherTitles
'    .FontStyle = GraphLib.FontStyleConstants.gphItalic
'    .DrawMode = GraphLib.DrawModeConstants.gphBlit
'  End With
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub CreateReport()
  On Error GoTo ERR_H
  
  Static arrVents As Variant
  Static arrSt As Variant
  If IsEmpty(arrVents) Then _
    arrVents = Array("I", "I", "II", "II", "III", "III", "IV", "IV")
  If IsEmpty(arrSt) Then _
    arrSt = Array("Гомеостазис", "Нарушение ==", "Адаптация D", "Опасная ситуация", "Адаптация F", "Кр.ситуация", "Адаптация H", "Происшествие")
  
  List1.Text = ""
  Dim sStr As String
  Const sPad1 = "       "
  Dim l As Long, ns As Tag_TypANodesStat, arr() As Double
  For l = 0 To 7
    m_mg.GetModelStat l, ns, arr, False
    
    If m_mg.StatIn Then
      sStr = sStr & IIf(l Mod 2 = 0, "До крана", "После крана") & " " & arrVents(l) & ":" & vbCrLf
    Else
      sStr = sStr & arrSt(l) & ":" & vbCrLf
    End If
                
    If m_mg.RunMode = MT_Analytical Or m_mg.RunMode = MT_AnalyticalIndistinct Then
      sStr = sStr & sPad1 & "Сумма вероятностей = " & Format(ns.dSummP, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Средняя вероятность = " & Format(ns.dAvgP, haApp.FmtDbl) & vbCrLf
    
      sStr = sStr & sPad1 & "Минимальное значение индекса опасности = " & Format(ns.dMinVal, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Максимальное значение индекса опасности = " & Format(ns.dMaxVal, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Мат. ожидание индекса опасности = " & Format(ns.dAvg, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Ср.кв. отклонение индекса опасности = " & Format(ns.dDx, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Мощность множества = " & Sprintf("%I64u", , ns.cySz) & vbCrLf
    Else
      sStr = sStr & sPad1 & "Минимальное значение выборки = " & Format(ns.dMinVal, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Максимальное значение выборки = " & Format(ns.dMaxVal, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Частота = " & Format(ns.dAvg, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Ср.кв. отклонение меры возможности = " & Format(ns.dDx, haApp.FmtDbl) & vbCrLf
      sStr = sStr & sPad1 & "Число реализаций = " & Sprintf("%I64u", , ns.cySz) & vbCrLf
    End If
    
     
    If l < 7 Then
      List1.Text = List1.Text & sStr & "--------------------" & vbCrLf
    Else
      List1.Text = List1.Text & sStr
    End If
    sStr = ""
  Next l
  
  FillHisto
    
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Friend Sub Clear()
'  Graph1.DataReset = gphAllData
'  Graph1.DrawMode = GraphLib.DrawModeConstants.gphBlit
  Dim arr() As Double
  Graph1.SetArrData arr
End Sub

Friend Function RndClr() As ColorDataConstants
  Static arrConst As Variant
  If IsEmpty(arrConst) Then _
    arrConst = Array(ColorDataConstants.gphBlack, ColorDataConstants.gphBlue, ColorDataConstants.gphBrown, _
      ColorDataConstants.gphCyan, ColorDataConstants.gphDarkGray, ColorDataConstants.gphGreen, ColorDataConstants.gphLightBlue, _
      ColorDataConstants.gphLightCyan, ColorDataConstants.gphLightGray, ColorDataConstants.gphLightGreen, _
      ColorDataConstants.gphLightMagenta, ColorDataConstants.gphLightRed, ColorDataConstants.gphMagenta, ColorDataConstants.gphRed, ColorDataConstants.gphYellow)
        
  RndClr = arrConst(CInt(UBound(arrConst) - LBound(arrConst)) * Rnd + LBound(arrConst))
        
End Function


