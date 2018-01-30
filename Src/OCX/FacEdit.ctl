VERSION 5.00
Object = "{FA9F41FF-762A-11D1-A98C-004845001083}#47.1#0"; "GTSlider.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#105.0#0"; "AlexOCX.ocx"
Begin VB.UserControl FacEdit 
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   ScaleHeight     =   2040
   ScaleWidth      =   6180
   Begin Threed.SSPanel spValue 
      Height          =   210
      Left            =   2130
      TabIndex        =   7
      ToolTipText     =   "Лингвистическая оценка фактора"
      Top             =   240
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   370
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Очень очень высокий"
      ClipControls    =   0   'False
      BevelOuter      =   0
   End
   Begin Threed.SSPanel ssDescr 
      Height          =   450
      Left            =   75
      TabIndex        =   9
      ToolTipText     =   "Фактор опасности"
      Top             =   1470
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   794
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SHN"
      ClipControls    =   0   'False
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel ssShortName 
      Height          =   360
      Left            =   75
      TabIndex        =   10
      ToolTipText     =   "Краткое обозначение фактора опасности"
      Top             =   45
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   635
      _Version        =   131074
      CaptionStyle    =   1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SHN"
      ClipControls    =   0   'False
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin GreenTreeSlider.GTSlider sldLingvo 
      Height          =   1110
      Left            =   1875
      TabIndex        =   0
      ToolTipText     =   "Лингвистическая оценка фактора"
      Top             =   150
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1958
      BevelInner      =   1
      BevelOuter      =   2
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
      TickStyle       =   1
      Value           =   1
      ThumbMaskColor  =   0
      PropA           =   1
   End
   Begin Threed.SSCommand cmdPr 
      Height          =   360
      Left            =   810
      TabIndex        =   4
      ToolTipText     =   "Настроить форму насыщения"
      Top             =   120
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Picture         =   "FacEdit.ctx":0000
      AutoSize        =   2
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdAskVal 
      Height          =   300
      Left            =   165
      TabIndex        =   6
      ToolTipText     =   "Установить значение оценки"
      Top             =   540
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
      _Version        =   131074
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "?"
      AutoSize        =   2
      ButtonStyle     =   3
   End
   Begin AlexOCX.AngularLabel lblTrustLevel 
      Height          =   977
      Left            =   1334
      TabIndex        =   11
      Top             =   300
      Width           =   197
      _ExtentX        =   344
      _ExtentY        =   1720
      FontCtx         =   "FacEdit.ctx":016A
      Caption         =   "Нормальный"
      AutoSize        =   -1  'True
   End
   Begin Threed.SSCommand cmdSetQ 
      Height          =   345
      Left            =   75
      TabIndex        =   8
      ToolTipText     =   "Настроить индексы опасности"
      Top             =   990
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      _Version        =   131074
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1/0"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdAss 
      Height          =   300
      Left            =   810
      TabIndex        =   5
      ToolTipText     =   "Назначить енумератор"
      Top             =   540
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Picture         =   "FacEdit.ctx":01C2
      AutoSize        =   2
      ButtonStyle     =   3
   End
   Begin Threed.SSOption opt3 
      Height          =   240
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Низкая уверенность"
      Top             =   1125
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      _Version        =   131074
   End
   Begin Threed.SSOption opt2 
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Средняя уверенность"
      Top             =   690
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      _Version        =   131074
      Value           =   -1
   End
   Begin Threed.SSOption opt1 
      Height          =   240
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Высокая уверенность"
      Top             =   270
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      _Version        =   131074
   End
End
Attribute VB_Name = "FacEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFacED

'Default Property Values:
Const m_def_MinWidthTwip = 5565#
Const m_def_HeightTwip = 2040#
Const m_def_SldRightXTwip = 135#
Const m_def_SSDescrRightXTwip = 83#
Const m_def_SPValueLeftXTwip = 255#
Const m_def_ACtiveColor = &H8000&
'Property Variables:
Private WithEvents m_f_Factor As GERTNETLib.Factor
Attribute m_f_Factor.VB_VarHelpID = -1
Private m_gn_GertNet As GERTNETLib.MGertNet

'Event Declarations:
'Event ClickAssEnum(ByVal feSender As FacEdit)
'Event ClickSetupValue(ByVal feSender As FacEdit)
'Event FactorChanged(ByVal fac As Factor)

'Int. var.
Private m_bLockResize As Boolean
Private m_bLockUpdate As Boolean
Private m_l_KeepBK As Long

Public Sub SetFocus()
  On Error GoTo ERR_H
  Extender.SetFocus
  Exit Sub
ERR_H:
End Sub

Public Property Get MHwnd() As Long
  MHwnd = UserControl.hwnd
End Property


Public Property Get MinimumHeightTwip() As Single
  MinimumHeightTwip = m_def_HeightTwip
End Property

Public Property Get MinimumWidthTwip() As Single
  MinimumWidthTwip = m_def_MinWidthTwip
End Property

Public Property Set HFactor(ByVal f As Factor)
  Set m_f_Factor = f
  SetupValues
End Property

Public Property Get HFactor() As Factor
  Set HFactor = m_f_Factor
End Property

Public Property Set GNet(ByVal f As MGertNet)
  Set m_gn_GertNet = f
End Property

Public Property Get GNet() As MGertNet
  Set GNet = m_gn_GertNet
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bFl As Boolean)
  UserControl.Enabled = bFl
  sldLingvo.Enabled = bFl
  cmdAss.Enabled = bFl
  cmdSetQ.Enabled = bFl
  opt1.Enabled = bFl
  opt2.Enabled = bFl
  opt3.Enabled = bFl
  PropertyChanged "Enabled"
End Property


Private Sub cmdAskVal_Click()
  RaiseClickAskVal
End Sub

Private Sub cmdAss_Click()
  RaiseClickAssEnum
End Sub

Private Sub cmdPr_Click()
  RaiseClickPr
End Sub

Private Sub cmdSetQ_Click()
  RaiseClickSetupValue
End Sub

Private Sub m_f_Factor_OnDirtyChanged(ByVal bFl As Boolean)

  Dim sTmp As String, bTst As Boolean
  sTmp = ssShortName.Caption
  bTst = Right(sTmp, 1) = "*"
  If bTst <> bFl Then
    Dim ibsKey As IBSTRKey
    Set ibsKey = m_f_Factor
    ssShortName.Caption = ibsKey.Get & IIf(bFl = True, "*", "")
  End If
End Sub

Private Sub opt1_Click(Value As Integer)
  lblTrustLevel.Caption = "Высокий"
  If opt1.Value = True And IsBound() And Not m_bLockUpdate Then
    m_f_Factor.TrustLvl = TL_High
    RaiseFactorChanged
  End If
End Sub

Private Sub opt2_Click(Value As Integer)
  lblTrustLevel.Caption = "Средний"
  If opt2.Value = True And IsBound() And Not m_bLockUpdate Then
    m_f_Factor.TrustLvl = TL_Normal
    RaiseFactorChanged
  End If
End Sub

Private Sub opt3_Click(Value As Integer)
  lblTrustLevel.Caption = "Низкий"
  If opt3.Value = True And IsBound() And Not m_bLockUpdate Then
    m_f_Factor.TrustLvl = TL_Low
    RaiseFactorChanged
  End If
End Sub

Private Function IsBound() As Boolean
  IsBound = Not ((m_gn_GertNet Is Nothing) Or (m_f_Factor Is Nothing))
End Function

Private Sub sldLingvo_Change()
  If Not m_bLockUpdate And IsBound() Then
    m_f_Factor.Value = sldLingvo.Value - 1
    RaiseFactorChanged
  End If
    
  If IsBound() Then
    Dim sDescr As String, dbVal As Double
    m_gn_GertNet.GetFactorPresent pF:=m_f_Factor, strDescription:=sDescr, dVal:=dbVal
    spValue.Caption = sDescr & " (" & CStr(dbVal) & ")"
  Else
    spValue.Caption = "<пусто>"
  End If
End Sub

Private Sub sldLingvo_Scroll()
  If IsBound() Then
    Dim fTmp As GERTNETLib.Factor
    Dim iCl As IClonable
    Set iCl = m_f_Factor
    Set fTmp = iCl.Clone()
    
    Dim sDescr As String, dbVal As Double
    fTmp.Value = sldLingvo.Value - 1
    m_gn_GertNet.GetFactorPresent pF:=fTmp, strDescription:=sDescr, dVal:=dbVal
    spValue.Caption = sDescr & " (" & CStr(dbVal) & ")"
  End If
End Sub

Private Sub spValue_Click()
  SetFocus
End Sub

Private Sub ssDescr_Click()
  SetFocus
End Sub

Private Sub ssShortName_Click()
  SetFocus
End Sub

Private Sub UserControl_EnterFocus()
  'm_l_KeepBK = ssShortName.ForeColor
  ssShortName.ForeColor = m_def_ACtiveColor
  ssShortName.Font3D = ssRaisedLight
  
  lblTrustLevel.ForeColor = &HFF0000
  lblTrustLevel.Bold = True
  spValue.ForeColor = &HFF0000
  spValue.Font3D = ssInsetLight
  spValue.Font.Bold = True
  ssDescr.Font3D = ssInsetLight
  ssDescr.Font.Bold = True
  
  RaiseFactorFocus
End Sub

Private Sub UserControl_ExitFocus()
'  ssShortName.ForeColor = m_l_KeepBK
'  ssShortName.Font3D = ssNoneFont3D
'
'  ssDescr.Font3D = ssNoneFont3D
'  spValue.Font3D = ssNoneFont3D
'  spValue.ForeColor = &H0
'  spValue.Font.Bold = False
'  lblTrustLevel.ForeColor = &H0
'  lblTrustLevel.Bold = False
'  ssDescr.Font.Bold = False
End Sub


Private Sub UserControl_Initialize()
  m_bLockResize = False
  m_bLockUpdate = False
  Enabled = False
  SetupValues
End Sub

Private Sub UserControl_Paint()
  Dim sP1 As Single, sP2 As Single
  sP1 = ScaleY(1, vbPixels, ScaleMode)
  sP2 = ScaleY(2, vbPixels, ScaleMode)
    
  Line (ScaleLeft, ScaleTop)-(ScaleLeft + ScaleWidth, ScaleTop), &HC0C0C0
  Line (ScaleLeft, ScaleTop + sP1)-(ScaleLeft + ScaleWidth, ScaleTop + sP1), &H808080
  
  Line (ScaleLeft, ScaleTop + ScaleHeight - sP2)-(ScaleLeft + ScaleWidth, ScaleTop + ScaleHeight - sP2), &H808080
  Line (ScaleLeft, ScaleTop + ScaleHeight - sP1)-(ScaleLeft + ScaleWidth, ScaleTop + ScaleHeight - sP1), &HFFFFFF
End Sub

Private Sub UserControl_Resize()
    
  If Not m_bLockResize Then
    m_bLockResize = True
    On Error GoTo ERR_H
        
        
    'Const m_def_MinWidthTwip = 5565
    'Const m_def_HeightTwip = 2040

    Dim w As Single, h As Single, iScm As Integer
    iScm = ContainerScale()
'    w = ScaleX(Extender.Width, iScm, vbTwips)
'    h = ScaleY(Extender.Height, iScm, vbTwips)
    
    w = Extender.Width
    h = Extender.Height
    
    If h <> m_def_HeightTwip Then h = m_def_HeightTwip
    If w < m_def_MinWidthTwip Then w = m_def_MinWidthTwip
        
    sldLingvo.Width = w - sldLingvo.Left - m_def_SldRightXTwip
    ssDescr.Width = w - ssDescr.Left - m_def_SSDescrRightXTwip
    
    spValue.Left = sldLingvo.Left + m_def_SPValueLeftXTwip
    spValue.Width = sldLingvo.Width - 2 * m_def_SPValueLeftXTwip
    'spValue.Move sldLingvo.Left + m_def_SPValueLeftXTwip, , sldLingvo.Width - 2 * m_def_SPValueLeftXTwip
        
    'Extender.Width = ScaleX(w, vbTwips, iScm)
    'Extender.Height = ScaleY(h, vbTwips, iScm)
    
    'Extender.Width = w
    'Extender.Height = h
    Extender.Move Extender.Left, Extender.Top, w, h
        
    m_bLockResize = False
  End If
  Exit Sub
ERR_H:
  m_bLockResize = False
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub UserControl_Terminate()
  Set m_f_Factor = Nothing
  Set m_gn_GertNet = Nothing
End Sub

Public Sub RefreshValues()
  SetupValues
End Sub

Private Sub SetupValues()

  m_bLockUpdate = True
  On Error GoTo ERR_H

  If m_f_Factor Is Nothing Then
    sldLingvo.Value = 1
    Enabled = False
    ssShortName.Caption = "<Н>"
    ssDescr.Caption = "<нет>"
    opt1.Value = True
    spValue.Caption = "<пусто>"
    cmdSetQ.Caption = "1/0"
  Else
    If Not Enabled Then Enabled = True
    sldLingvo.Value = m_f_Factor.Value + 1
    Dim ibsKey As IBSTRKey
    Set ibsKey = m_f_Factor
    If m_f_Factor.IsDirty Then
      ssShortName.Caption = ibsKey.Get & "*"
    Else
      ssShortName.Caption = ibsKey.Get
    End If
    ssDescr.Caption = m_f_Factor.Name
    Select Case m_f_Factor.TrustLvl
      Case TL_High
        opt1.Value = True
      Case TL_Normal
        opt2.Value = True
      Case TL_Low
        opt3.Value = True
    End Select
    
    Dim l As Long, l2 As Long, sTmp As String
    l2 = m_f_Factor.NIdx
    For l = 0 To l2 - 1
      sTmp = sTmp & CStr(m_f_Factor.Idx(l))
      If l < l2 - 1 Then sTmp = sTmp & "/"
    Next l
    cmdSetQ.Caption = sTmp
    
  End If
  
  m_bLockUpdate = False
  Exit Sub
ERR_H:
  m_bLockUpdate = False
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Function ContainerScale() As Integer
' **** Gets the proper mode to scale to
    Dim eScaleTo As ScaleModeConstants
    
    On Error Resume Next
    
    ContainerScale = Extender.Container.ScaleMode
    
    If Err Then
        Err = 0
        ContainerScale = Parent.ScaleMode
        If Err Then
            Err = 0
            ContainerScale = vbTwips
        End If
    End If
    
    On Error GoTo 0
End Function


Private Sub RaiseClickPr()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.ClickPr m_f_Factor
End Sub

Private Sub RaiseClickAssEnum()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.ClickAssEnum m_f_Factor
End Sub
Private Sub RaiseClickAskVal()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.ClickAskVal m_f_Factor
End Sub
Private Sub RaiseClickSetupValue()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.ClickSetupValue m_f_Factor
End Sub
Private Sub RaiseFactorChanged()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.FactorChanged m_f_Factor
End Sub
Private Sub RaiseFactorFocus()
  Dim ife As IFECallback
  Set ife = UserControl.Parent
  ife.FactorFocus Me
End Sub

Public Sub ClearFocus()

  ssShortName.ForeColor = RGB(0, 0, 0)
  ssShortName.Font3D = ssNoneFont3D
  
  ssDescr.Font3D = ssNoneFont3D
  spValue.Font3D = ssNoneFont3D
  spValue.ForeColor = &H0
  spValue.Font.Bold = False
  lblTrustLevel.ForeColor = &H0
  lblTrustLevel.Bold = False
  ssDescr.Font.Bold = False
End Sub

Public Property Get IFacED_HFactor() As Factor
  Set IFacED_HFactor = HFactor
End Property

Public Sub IFacED_SetFocus()
  SetFocus
End Sub

Public Sub IFacED_ClearFocus()
  ClearFocus
End Sub


