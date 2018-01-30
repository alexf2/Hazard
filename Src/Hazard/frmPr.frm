VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Object = "{FA9F41FF-762A-11D1-A98C-004845001083}#47.1#0"; "GTSlider.ocx"
Begin VB.Form frmPr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Форма насыщения"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   HasDC           =   0   'False
   Icon            =   "frmPr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   3285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5794
      _Version        =   131074
      BackColor       =   12632256
      PictureBackgroundStyle=   2
      ClipControls    =   0   'False
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVNumericLib.PVNumeric pvnForm 
         Height          =   315
         Left            =   2708
         TabIndex        =   3
         Top             =   915
         Width           =   1785
         _Version        =   393216
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0,5"
         ForeColor       =   0
         BackColor       =   16776176
         Appearance      =   1
         BackColor       =   16776176
         ForeColor       =   0
         SpinButtons     =   0
         SuppressThousand=   0   'False
         DecimalSeparator=   ","
         LimitValue      =   -1  'True
         ValueMin        =   0.00000001
         ValueMax        =   214748364
         ValueSpinDelta  =   0
         ValueReal       =   0.5
         DecimalMax      =   9
         ValidateLimit   =   1
         ChangeColor     =   -1  'True
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   210
         Left            =   653
         TabIndex        =   6
         Top             =   1470
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         CaptionStyle    =   1
         ForeColor       =   0
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
         Caption         =   "Расположение медианы:"
         ClipControls    =   0   'False
         BevelOuter      =   0
         Alignment       =   1
      End
      Begin Threed.SSPanel spValue 
         Height          =   210
         Left            =   3713
         TabIndex        =   5
         Top             =   1455
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   370
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0.50"
         ClipControls    =   0   'False
         BevelOuter      =   0
         Alignment       =   4
      End
      Begin GreenTreeSlider.GTSlider sldPlacement 
         Height          =   1080
         Left            =   555
         TabIndex        =   4
         ToolTipText     =   "Лингвистическая оценка фактора"
         Top             =   1365
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   1905
         BevelInner      =   1
         BevelOuter      =   2
         BackColor       =   12632256
         LargeChange     =   3
         SmallChange     =   1
         Max             =   11
         Min             =   1
         ThumbStyle      =   2
         ThumbWidth      =   9
         BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TickLength      =   5
         TickLongLength  =   5
         Value           =   1
         ThumbMaskColor  =   0
         PropA           =   1
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   210
         Left            =   1298
         TabIndex        =   7
         Top             =   960
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Параметр формы"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSCommand ssNo 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   353
         TabIndex        =   10
         ToolTipText     =   "Отказаться от модификации"
         Top             =   2700
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPr.frx":000C
         Caption         =   "Отмена"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssApply 
         Default         =   -1  'True
         Height          =   375
         Left            =   3083
         TabIndex        =   9
         ToolTipText     =   "Вносит изменения в модель без закрытия диалога"
         Top             =   2700
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPr.frx":0176
         Caption         =   "Применить"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssClose 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1718
         TabIndex        =   8
         ToolTipText     =   "Закрывает диалог, сохраняя модификации"
         Top             =   2700
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPr.frx":02E0
         Caption         =   "Закрыть"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSOption optCauchi 
         Height          =   300
         Left            =   195
         TabIndex        =   2
         Top             =   570
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   131074
         BackStyle       =   1
         Caption         =   "Насыщение по Коши"
      End
      Begin Threed.SSOption optUniform 
         Height          =   300
         Left            =   188
         TabIndex        =   1
         ToolTipText     =   "Используется равномерное распределение"
         Top             =   165
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         _Version        =   131074
         BackStyle       =   1
         Caption         =   "Максимальная энтропия"
         Value           =   -1
      End
   End
End
Attribute VB_Name = "frmPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_g As GERTNETLib.MGertNet
Private m_f As GERTNETLib.Factor
Private m_le As GERTNETLib.LingvoEnum

Private m_stKeep As GERTNETLib.PrpbDistrTyp
Private m_dCPKeep As Double
Private m_dCSKeep As Double

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssNo_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssApply_Click
  End If
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssNo_Click
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set m_g = Nothing
  Set m_f = Nothing
  Set m_le = Nothing
End Sub

Private Sub optCauchi_Click(Value As Integer)
  If optCauchi.Value Then
    pvnForm.Enabled = True
    sldPlacement.Enabled = True
    spValue.Enabled = True
  End If
End Sub

Private Sub optUniform_Click(Value As Integer)
  If optUniform.Value Then
    pvnForm.Enabled = False
    sldPlacement.Enabled = False
    spValue.Enabled = False
  End If
End Sub

Private Sub pvnForm_Validate(Cancel As Boolean)
  If pvnForm.ValueReal < pvnForm.ValueMin Or pvnForm.ValueReal > pvnForm.ValueMax Then
    Cancel = True
    MsgBox "Ошибочное значение параметра формы. Требуется " & pvnForm.ValueMin & " <= Val <= " & pvnForm.ValueMax, vbExclamation Or vbOKOnly, "Ошибка"
  End If
End Sub


Private Sub sldPlacement_Change()
  spValue.Caption = CStr(m_le.Quality(sldPlacement.Value - 1))
End Sub

Private Sub sldPlacement_Scroll()
  spValue.Caption = CStr(m_le.Quality(sldPlacement.Value - 1))
End Sub

Private Sub ssNo_Click()
  m_f.DistrType = m_stKeep
  m_f.CochiPlacement = m_dCPKeep
  m_f.CochiScale = m_dCSKeep
  RefreshView
  
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


Private Sub ssClose_Click()
  Hide
End Sub

Private Sub ssClose_LostFocus()
  HighLight2 ssClose, False, Me.hWnd
End Sub

Private Sub ssClose_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, True, Me.hWnd
End Sub

Private Sub ssClose_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, False, Me.hWnd
End Sub

Private Sub ssApply_LostFocus()
  HighLight2 ssApply, False, Me.hWnd
End Sub

Private Sub ssApply_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssApply, True, Me.hWnd
End Sub

Private Sub ssApply_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssApply, False, Me.hWnd
End Sub

Private Property Get Flag() As Boolean
  Flag = IIf(m_f.DistrType = PDT_Cauchy, True, False)
End Property

Private Sub ssApply_Click()
  Dim bFl As Boolean: bFl = False
  pvnForm_Validate bFl
  If bFl Then Exit Sub
  
  m_f.DistrType = IIf(optUniform.Value, PDT_Uniform, PDT_Cauchy)
  If m_f.DistrType = PDT_Cauchy Then
    m_f.CochiPlacement = m_le.Quality(sldPlacement.Value - 1)
    m_f.CochiScale = pvnForm.ValueReal
  End If
  
  RefreshView
End Sub

Private Sub RefreshView()
  m_g.Run 5, 10, False, -1, True
  frmMEditor.ShowSPFunc m_f
End Sub


Friend Sub AssData(ByVal g As MGertNet, ByVal f As Factor)
  Set m_g = g
  Set m_f = f
  
  m_stKeep = f.DistrType
  m_dCPKeep = f.CochiPlacement
  m_dCSKeep = f.CochiScale
  
  Dim ibs As IBSTRKey
  Set ibs = f
  Caption = "Форма насыщения для '" & ibs.Get() & "'"
  pvnForm.ValueReal = f.CochiScale
  Set m_le = g.GetEnumForFactor(f)
  Dim l As Long
  For l = 0 To 10
    If m_le.Quality(l) = f.CochiPlacement Then
      sldPlacement.Value = l + 1
      Exit For
    End If
  Next l
      
  pvnForm.Enabled = Flag
  sldPlacement.Enabled = Flag
  spValue.Enabled = Flag
  If Flag Then optCauchi.Value = True Else optUniform.Value = True
End Sub

Friend Sub UnassData()
  Set m_g = Nothing
  Set m_f = Nothing
  Set m_le = Nothing
End Sub


