VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{FA9F41FF-762A-11D1-A98C-004845001083}#47.1#0"; "GTSlider.ocx"
Begin VB.Form frmRepOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Автоматические отчёты"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmRepOpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   6509
      _Version        =   131074
      PictureBackgroundStyle=   2
      ClipControls    =   0   'False
      BorderWidth     =   0
      BevelOuter      =   0
      Begin VsOcxLib.VideoSoftIndexTab vsTabOpt 
         Height          =   2820
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   5745
         _Version        =   327680
         _ExtentX        =   10134
         _ExtentY        =   4974
         _StockProps     =   102
         Caption         =   "Генерация отчётов|Формат вероятностей|Формат прочих величин"
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
            Height          =   1875
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   900
            Width           =   5265
            _Version        =   327680
            _ExtentX        =   9287
            _ExtentY        =   3307
            _StockProps     =   70
            ConvInfo        =   1418783674
            BevelOuter      =   0
            Begin Threed.SSCheck ssOpt 
               Height          =   225
               Left            =   1050
               TabIndex        =   7
               Top             =   1140
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   397
               _Version        =   131074
               Caption         =   "Оптимизация"
            End
            Begin Threed.SSCheck ssVApply 
               Height          =   225
               Left            =   1050
               TabIndex        =   6
               Top             =   750
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   397
               _Version        =   131074
               Caption         =   "Оценка эффективности набора мер"
            End
            Begin Threed.SSCheck ssRun 
               Height          =   225
               Left            =   1050
               TabIndex        =   5
               Top             =   360
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   397
               _Version        =   131074
               Caption         =   "Прогон"
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
            Height          =   1875
            Left            =   75045
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   900
            Width           =   5265
            _Version        =   327680
            _ExtentX        =   9287
            _ExtentY        =   3307
            _StockProps     =   70
            ConvInfo        =   1418783674
            BevelOuter      =   0
            Begin Threed.SSPanel SSPanel4 
               Height          =   210
               Left            =   2265
               TabIndex        =   13
               Top             =   510
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   370
               _Version        =   131074
               Font3D          =   3
               ForeColor       =   0
               Caption         =   "Число цифр:"
               ClipControls    =   0   'False
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
            End
            Begin Threed.SSFrame SSFrame2 
               Height          =   1050
               Left            =   150
               TabIndex        =   9
               Top             =   285
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   1852
               _Version        =   131074
               Font3D          =   3
               Caption         =   "Формат"
               Begin Threed.SSOption ssFloatPr 
                  Height          =   270
                  Left            =   150
                  TabIndex        =   11
                  Top             =   630
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   476
                  _Version        =   131074
                  Caption         =   "Плавающий"
               End
               Begin Threed.SSOption ssNormalPr 
                  Height          =   270
                  Left            =   150
                  TabIndex        =   10
                  Top             =   270
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   476
                  _Version        =   131074
                  Caption         =   "Фиксированный"
                  Value           =   -1
               End
            End
            Begin Threed.SSPanel spDigitsPr 
               Height          =   210
               Left            =   3420
               TabIndex        =   14
               Top             =   510
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   370
               _Version        =   131074
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "5"
               ClipControls    =   0   'False
               BevelOuter      =   0
               Alignment       =   1
            End
            Begin GreenTreeSlider.GTSlider gtsPr 
               Height          =   540
               Left            =   2205
               TabIndex        =   15
               Top             =   765
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   953
               LargeChange     =   3
               SmallChange     =   1
               Max             =   15
               Min             =   1
               ThumbWidth      =   10
               TickCount       =   15
               BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               ThumbMaskColor  =   0
               PropA           =   1
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic3 
            Height          =   1875
            Left            =   75345
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   900
            Width           =   5265
            _Version        =   327680
            _ExtentX        =   9287
            _ExtentY        =   3307
            _StockProps     =   70
            ConvInfo        =   1418783674
            BevelOuter      =   0
            Begin Threed.SSPanel SSPanel2 
               Height          =   210
               Left            =   2265
               TabIndex        =   16
               Top             =   480
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   370
               _Version        =   131074
               Font3D          =   3
               ForeColor       =   0
               Caption         =   "Число цифр:"
               ClipControls    =   0   'False
               BevelOuter      =   0
               AutoSize        =   2
               Alignment       =   0
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   1050
               Left            =   150
               TabIndex        =   17
               Top             =   225
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   1852
               _Version        =   131074
               Font3D          =   3
               Caption         =   "Формат"
               Begin Threed.SSOption ssNormalOther 
                  Height          =   270
                  Left            =   150
                  TabIndex        =   19
                  Top             =   270
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   476
                  _Version        =   131074
                  Caption         =   "Фиксированный"
                  Value           =   -1
               End
               Begin Threed.SSOption ssFloatOther 
                  Height          =   270
                  Left            =   150
                  TabIndex        =   18
                  Top             =   630
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   476
                  _Version        =   131074
                  Caption         =   "Плавающий"
               End
            End
            Begin Threed.SSPanel spDigitsOther 
               Height          =   210
               Left            =   3405
               TabIndex        =   20
               Top             =   480
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   370
               _Version        =   131074
               CaptionStyle    =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "5"
               ClipControls    =   0   'False
               BevelOuter      =   0
               Alignment       =   1
            End
            Begin GreenTreeSlider.GTSlider gtsOther 
               Height          =   540
               Left            =   2205
               TabIndex        =   21
               Top             =   705
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   953
               LargeChange     =   3
               SmallChange     =   1
               Max             =   15
               Min             =   1
               ThumbWidth      =   10
               TickCount       =   15
               BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               ThumbMaskColor  =   0
               PropA           =   1
            End
         End
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   525
         Left            =   570
         TabIndex        =   1
         Top             =   3030
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepOpt.frx":000C
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   525
         Left            =   3375
         TabIndex        =   2
         Top             =   3015
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmRepOpt.frx":0176
         Caption         =   "Подтверждение"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmRepOpt"
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
  'Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
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


Private Sub gtsPr_Change()
  spDigitsPr.Caption = gtsPr.Value
End Sub

Private Sub gtsPr_Scroll()
  spDigitsPr.Caption = gtsPr.Value
End Sub

Private Sub gtsOther_Change()
  spDigitsOther.Caption = gtsOther.Value
End Sub

Private Sub gtsOther_Scroll()
  spDigitsOther.Caption = gtsOther.Value
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

Private Sub ssOk_Click()
  m_b_Res = True
  Hide
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

Public Sub AssData(ByVal haApp As HazardApp)
  m_b_Res = False
  Me.ssRun.Value = haApp.CreRepRun
  Me.ssOpt.Value = haApp.CreRepOpt
  Me.ssVApply.Value = haApp.CreRepVApply
  
  With haApp.GertNetMain
    If haApp.HaveDoc Then
      If .NFormatPr = NF_Normal Then _
        ssNormalPr.Value = True _
      Else ssFloatPr.Value = True
      
      If .NFormatOther = NF_Normal Then _
        ssNormalOther.Value = True _
      Else ssFloatOther.Value = True
      
      gtsPr.Value = .NDigitsPr
      gtsOther.Value = .NDigitsOther
      
    End If
  End With
End Sub

