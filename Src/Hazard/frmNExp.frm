VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmNExp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Число опытов для поиска Pпр."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frmNExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4471
      _Version        =   131074
      BackColor       =   12632256
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVNumericLib.PVNumeric pvnResMax 
         Height          =   315
         Left            =   2625
         TabIndex        =   5
         ToolTipText     =   "Верхняя оценка делается согласно неравенству Чебышева (т.е. без учёта закона распределения)"
         Top             =   1245
         Width           =   1140
         _Version        =   393216
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         ReadOnly        =   -1  'True
         FreeFormEntry   =   -1  'True
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -2147483648
         ValueMax        =   2147483647
         LimitValueByType=   3
         DecimalMax      =   0
         ValidateLimit   =   -16364
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnK 
         Height          =   315
         Left            =   652
         TabIndex        =   2
         ToolTipText     =   "Определяет ширину интервала отклонения мат.ожидания: K*Сигма"
         Top             =   1245
         Width           =   1140
         _Version        =   393216
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,25"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         FreeFormEntry   =   -1  'True
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0.00001
         ValueMax        =   1
         ValueReal       =   0.25
         DecimalMin      =   1
         DecimalMax      =   16
         ValidateLimit   =   -16364
         ChangeColor     =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnResMin 
         Height          =   315
         Left            =   2625
         TabIndex        =   4
         ToolTipText     =   "Нижняя оценка числа опытов делается исходя из допущения о нормальном законе распределения"
         Top             =   570
         Width           =   1140
         _Version        =   393216
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         ReadOnly        =   -1  'True
         FreeFormEntry   =   -1  'True
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   -2147483648
         ValueMax        =   2147483647
         LimitValueByType=   3
         DecimalMax      =   0
         ValidateLimit   =   -16364
         ChangeColor     =   -1  'True
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFBF0&
         Height          =   315
         ItemData        =   "frmNExp.frx":000C
         Left            =   630
         List            =   "frmNExp.frx":0022
         TabIndex        =   1
         Text            =   "0.95"
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Число опытов max"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2625
         TabIndex        =   12
         Top             =   1020
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Число опытов min"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2625
         TabIndex        =   11
         Top             =   345
         Width           =   1365
      End
      Begin Threed.SSCommand ssCalc 
         Height          =   360
         Left            =   1950
         TabIndex        =   3
         Top             =   810
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   635
         _Version        =   131074
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "=>"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   337
         TabIndex        =   10
         Top             =   1305
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   337
         TabIndex        =   9
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   652
         TabIndex        =   8
         Top             =   1020
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   7
         Top             =   195
         Width           =   105
      End
      Begin Threed.SSCommand ssOk 
         Height          =   435
         Left            =   1658
         TabIndex        =   6
         Top             =   1875
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmNExp.frx":0049
         Caption         =   "Закрыть"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmNExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_b_Res As Boolean
Private m_mg As MGertNet
Attribute m_mg.VB_VarHelpID = -1

Private Sub Form_Initialize()
  m_b_Res = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssOk_Click
  ElseIf KeyAscii = vbKeyReturn Then
    'ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    'ssOk_Click
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_mg = Nothing
End Sub

Private Sub MsgErrBnd(ByVal pv As PVNumeric)
  
End Sub

Private Function GetP() As Double
  Select Case Combo1.ListIndex
    Case 0, -1
      GetP = 1.96
    Case 1
      GetP = 1.645
    Case 2
      GetP = 1.44
    Case 3
      GetP = 1.282
    Case 4
      GetP = 1.1625
    Case 5
      GetP = 1.0366
  End Select
End Function

Private Function GetP0() As Double
  Select Case Combo1.ListIndex
    Case 0, -1
      GetP0 = 0.95
    Case 1
      GetP0 = 0.9
    Case 2
      GetP0 = 0.85
    Case 3
      GetP0 = 0.8
    Case 4
      GetP0 = 0.75
    Case 5
      GetP0 = 0.7
  End Select
End Function


Private Sub ssCalc_Click()
  If pvnK.ValueReal < pvnK.ValueMin Or pvnK.ValueReal > pvnK.ValueMax Then
    MsgBox "K должно быть: " & CStr(pvnK.ValueMin) & " <= K <= " & CStr(pvnK.ValueMax), vbExclamation Or vbOKOnly, "Ошибка"
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double
  m_mg.GetInfSampleK 73, lMin, lMax, dMx, dDx
  
'  If dDx = 0 Then
'    MsgBox "Пробное число прогонов не достаточно. Требуется выбрать такое число, чтобы было получено ненулевое ср.кв. отклонение." & CStr(pvnK.ValueMax), vbExclamation Or vbOKOnly, "Ошибка"
'    Exit Sub
'  End If
  
  Dim dTmp As Double, dTmp2 As Double
  dTmp = GetP()
  dTmp2 = pvnK.ValueReal
  pvnResMin.ValueReal = CLng(dTmp * dTmp / (dTmp2 * dTmp2)) + 1
  pvnResMax.ValueReal = CLng(1# / (pvnK.ValueReal * pvnK.ValueReal) / (1# - GetP0())) + 1
  
  Exit Sub
ERR_H:
  MsgBox "Перед оценкой числа опытов требуется осуществить пробный прогон модели", vbOKOnly Or vbExclamation, "Нет данных"
  Err.Clear
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


Private Sub ssCalc_LostFocus()
  HighLight2 ssCalc, False, Me.hWnd
End Sub

Private Sub ssCalc_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCalc, True, Me.hWnd
End Sub

Private Sub ssCalc_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCalc, False, Me.hWnd
End Sub


Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Private Sub ssOk_Click()
  m_b_Res = True
  Set m_mg = Nothing
  Unload Me
End Sub


Friend Sub AssData(ByVal m As MGertNet)
  m_b_Res = False
  Set m_mg = m
End Sub

Friend Sub UnassData()
  Set m_mg = Nothing
End Sub


