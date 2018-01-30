VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#6.0#0"; "pvcurr.ocx"
Begin VB.Form frmEditSP 
   Caption         =   "Мера"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmEditSP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   9446
      _Version        =   131074
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVCurrencyLib.PVCurrency pvnCost 
         Height          =   315
         Left            =   5400
         TabIndex        =   9
         Top             =   3480
         Width           =   1725
         _Version        =   393216
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "100,00р."
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
         Value           =   100
         ValueMax        =   910000000000000
         ValidateLimit   =   1
      End
      Begin PVTreeView3Lib.PVTreeViewX tvFacAll 
         Height          =   2400
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   3210
         _Version        =   393216
         _ExtentX        =   5662
         _ExtentY        =   4233
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   9450828
         Appearance      =   1
         SelectMode      =   1
         StandardDefaultPicture=   10
         Sort            =   -1  'True
         AlwaysShowSelection=   -1  'True
         AllowInPlaceEditing=   0   'False
         BackColor       =   9450828
         ForeColor       =   0
         EnableToolTips  =   0   'False
         SelectedTextBackColor=   16711680
         SelectedTextForeColor=   16777215
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
      Begin Threed.SSPanel sspAllFac 
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   60
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Все доступные факторы опасности"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin VB.ListBox lstSP 
         BackColor       =   &H00E1FFFF&
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
         Height          =   1920
         IntegralHeight  =   0   'False
         ItemData        =   "frmEditSP.frx":000C
         Left            =   3840
         List            =   "frmEditSP.frx":000E
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   3270
         Width           =   1350
      End
      Begin VB.ListBox lstSPAll 
         BackColor       =   &H00E1FFFF&
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
         Height          =   1920
         IntegralHeight  =   0   'False
         ItemData        =   "frmEditSP.frx":0010
         Left            =   255
         List            =   "frmEditSP.frx":0012
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   3375
         Width           =   1350
      End
      Begin TrueOleDBGrid60.TDBGrid tdbgFac 
         Height          =   2790
         Left            =   3990
         OleObjectBlob   =   "frmEditSP.frx":0014
         TabIndex        =   4
         Top             =   90
         Width           =   3345
      End
      Begin Threed.SSPanel sspAllSP 
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   3105
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Все доступные меры"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   210
         Left            =   3765
         TabIndex        =   14
         Top             =   3045
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   255
         BackStyle       =   1
         Caption         =   "Несовместимые меры"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSCommand ssSPRight 
         Height          =   345
         Left            =   2415
         TabIndex        =   7
         ToolTipText     =   "Удалить несовместимые меры"
         Top             =   4260
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         _Version        =   131074
         ForeColor       =   0
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">>"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssSPLeft 
         Height          =   345
         Left            =   2430
         TabIndex        =   6
         ToolTipText     =   "Добавить несовместимые меры"
         Top             =   3660
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         _Version        =   131074
         ForeColor       =   0
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<<"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssFacRight 
         Height          =   345
         Left            =   3510
         TabIndex        =   3
         ToolTipText     =   "Удалить факторы"
         Top             =   1470
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         _Version        =   131074
         ForeColor       =   0
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">>"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssFacLeft 
         Height          =   345
         Left            =   3525
         TabIndex        =   2
         ToolTipText     =   "Добавить факторы"
         Top             =   885
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         _Version        =   131074
         ForeColor       =   0
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<<"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5310
         TabIndex        =   11
         ToolTipText     =   "Отменяет любые модификации и закрывает диалог"
         Top             =   4635
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         Picture         =   "frmEditSP.frx":43B0
         Caption         =   "Отмена"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOk 
         Default         =   -1  'True
         Height          =   525
         Left            =   5355
         TabIndex        =   10
         ToolTipText     =   "Выполняет модификации и закрывает диалог"
         Top             =   3915
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Windowless      =   0   'False
         Picture         =   "frmEditSP.frx":451A
         Caption         =   "Подтверждение"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSPanel sspCost 
         Height          =   210
         Left            =   5400
         TabIndex        =   15
         Top             =   3210
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Затраты"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin VB.Label NoDrag 
         Caption         =   "NoDrag"
         DragIcon        =   "frmEditSP.frx":4684
         Height          =   255
         Left            =   2865
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label SaveIcon 
         Caption         =   "SaveIcon"
         Height          =   255
         Left            =   2970
         TabIndex        =   18
         Top             =   3750
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label RemoveIcon 
         Caption         =   "RemoveIcon"
         DragIcon        =   "frmEditSP.frx":498E
         Height          =   255
         Left            =   2955
         TabIndex        =   17
         Top             =   4065
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label AddIcon 
         Caption         =   "AddIcon"
         DragIcon        =   "frmEditSP.frx":4C98
         Height          =   255
         Left            =   2925
         TabIndex        =   16
         Top             =   4395
         Visible         =   0   'False
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmEditSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TDragTypes
  DT_None = 0
  DT_RemoveSP = 1
  DT_AddSP = 2
  DT_RemoveFac = 4
  DT_AddFac = 8
End Enum

Private m_dt_DragType As TDragTypes


Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400
Private Const LST_WID = (1# / 5#)

'Private m_sp As SafetyPrecaution
Private m_spOriginal As SafetyPrecaution
Private m_gn As MGertNet
Private m_coll_SF As collSF
Private m_dba_x As XArrayDB
Private m_b_Res As Boolean

Private m_b_LockResize As Boolean


Private m_pvNodeH As PVBranch
Private m_pvNodeM As PVBranch
Private m_pvNodeC As PVBranch
Private m_pvNodeT As PVBranch

Private HighlightedNode As PVBranch
Private m_myFont As StdFont

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If m_dt_DragType <> DT_None Then
    If (Button And vbLeftButton) = 0 Then
      EndDragMode DT_None
    End If
  End If
End Sub

Private Function EndDragMode(ByVal mask As TDragTypes) As Boolean
  EndDragMode = (mask And m_dt_DragType) <> 0
  m_dt_DragType = DT_None
End Function


'Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'  DragValid Source, DT_None, State
'End Sub

Private Sub Form_Initialize()
  m_b_LockResize = False
  m_b_Res = False
  m_dt_DragType = DT_None
End Sub

Private Sub BeginDragMode(ByVal ctl As Control, ByVal dt As TDragTypes)
  m_dt_DragType = dt
  ctl.Drag vbBeginDrag
End Sub

Friend Property Get CurrDragIcon() As StdPicture
  Select Case m_dt_DragType
    Case DT_None
      Set CurrDragIcon = Nothing
    Case DT_RemoveSP
      Set CurrDragIcon = RemoveIcon.DragIcon
    Case DT_AddSP
      Set CurrDragIcon = AddIcon.DragIcon
    Case DT_RemoveFac
      Set CurrDragIcon = RemoveIcon.DragIcon
    Case DT_AddFac
      Set CurrDragIcon = AddIcon.DragIcon
  End Select
End Property

Private Function DragValid(ByVal src As Control, ByVal mask As TDragTypes, ByVal State As Integer) As Integer
  If (mask And m_dt_DragType) = m_dt_DragType Then
    DragValid = True
    'Exit Function
  Else
    DragValid = False
  End If
      
  
  Select Case State
    Case vbEnter
      'SaveIcon.DragIcon = src.DragIcon
      If Not DragValid Then
        src.DragIcon = NoDrag.DragIcon
      Else
        src.DragIcon = CurrDragIcon
      End If
    Case vbLeave
      src.DragIcon = NoDrag.DragIcon
  End Select
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  tvFacAll.StandardDefaultPicture = pvtNone
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
  
  Set m_myFont = tvFacAll.CreateFont
  m_myFont.Bold = True
  
  With tdbgFac
     .Columns(0).Alignment = dbgRight
     .Columns(0).Locked = True
     .Columns(0).DefaultValue = "<Пусто>"
     .Columns(0).HeadAlignment = dbgCenter
     .Columns(0).ForeColor = RGB(&H80, &H80, &H80)
     
     .Columns(1).Alignment = dbgLeft
     .Columns(1).Locked = False
     .Columns(1).DefaultValue = 0
     .Columns(1).HeadAlignment = dbgCenter
     .Columns(1).Font.Bold = True
     
     .Columns(2).Alignment = dbgLeft
     .Columns(2).Locked = False
     .Columns(2).DefaultValue = 0
     .Columns(2).HeadAlignment = dbgCenter
     .Columns(2).Font.Bold = True
     
     '.Columns(2).Visible = False
     
     .EvenRowStyle.BackColor = &H80FFFF
     .OddRowStyle.BackColor = &HC0FFFF
     .AlternatingRowStyle = True
  End With
  
  With tvFacAll
    Set m_pvNodeH = .Branches.Add(pvtPositionInOrder, 0, "Человек")
    m_pvNodeH.ForeColor = RGB(237, 232, 83)
    Set m_pvNodeH.CustomItemPicture = frmComplSP.PimgH.Picture
    m_pvNodeH.Data = PVT_H
    
    Set m_pvNodeM = .Branches.Add(pvtPositionInOrder, 0, "Машина")
    m_pvNodeM.ForeColor = RGB(237, 232, 83)
    Set m_pvNodeM.CustomItemPicture = frmComplSP.PimgM.Picture
    m_pvNodeM.Data = PVT_M
    
    Set m_pvNodeC = .Branches.Add(pvtPositionInOrder, 0, "Среда")
    m_pvNodeC.ForeColor = RGB(237, 232, 83)
    Set m_pvNodeC.CustomItemPicture = frmComplSP.PimgC.Picture
    m_pvNodeC.Data = PVT_C
    
    Set m_pvNodeT = .Branches.Add(pvtPositionInOrder, 0, "Технология")
    m_pvNodeT.ForeColor = RGB(237, 232, 83)
    Set m_pvNodeT.CustomItemPicture = frmComplSP.PimgT.Picture
    m_pvNodeT.Data = PVT_T
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ssCancel_Click
  End If
  
End Sub


Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 5500 Then Width = 5500
    If Height < 5700 Then Height = 5700
  Else
    m_b_LockResize = False
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  Dim lBW As Long, lBH As Long
  lBW = ScaleX(2 * GetSystemMetrics(SM_CXBORDER), vbPixels, vbTwips)
  lBH = ScaleY(2 * GetSystemMetrics(SM_CYBORDER), vbPixels, vbTwips)
  
  SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
      
  Dim sBaseW As Single
  sBaseW = ScaleWidth - 4 * L_PAD - ssFacLeft.Width
  
  sspAllFac.Move L_PAD, T_PAD
  tvFacAll.Move L_PAD, sspAllFac.Top + sspAllFac.Height, sBaseW * 2# / 3#, (ScaleHeight - 3 * T_PAD - lBH - 2 * sspAllFac.Height) / 1.5
  ssFacLeft.Move tvFacAll.Left + tvFacAll.Width + L_PAD, tvFacAll.Top + (tvFacAll.Height - 2 * ssFacLeft.Height + T_PAD) / 2
  ssFacRight.Move ssFacLeft.Left, ssFacLeft.Top + ssFacLeft.Height + T_PAD
  tdbgFac.Move ssFacLeft.Left + ssFacLeft.Width + L_PAD, sspAllFac.Top, sBaseW * 1# / 3# - lBW, tvFacAll.Height + sspAllFac.Height
  
  sspAllSP.Move L_PAD, tvFacAll.Top + tvFacAll.Height + T_PAD
  lstSPAll.Move L_PAD, sspAllSP.Top + sspAllSP.Height, (ScaleWidth - 5 * L_PAD - lBW - ssSPLeft.Width - ssOk.Width) / 2, ScaleHeight - (sspAllSP.Top + sspAllSP.Height) - T_PAD - lBH
  ssSPLeft.Move lstSPAll.Left + lstSPAll.Width + L_PAD, lstSPAll.Top + (lstSPAll.Height - 2 * ssSPLeft.Height - T_PAD) / 2
  ssSPRight.Move ssSPLeft.Left, ssSPLeft.Top + ssSPLeft.Height + T_PAD
  
  SSPanel2.Move ssSPLeft.Left + ssSPLeft.Width + L_PAD, sspAllSP.Top
  lstSP.Move SSPanel2.Left, SSPanel2.Top + SSPanel2.Height, lstSPAll.Width, lstSPAll.Height
  
  sspCost.Move lstSP.Left + lstSP.Width + L_PAD, lstSP.Top + (lstSP.Height - 2 * ssOk.Height - 2 * T_PAD - pvnCost.Height - sspCost.Height) / 2
  pvnCost.Move sspCost.Left, sspCost.Top + sspCost.Height
  ssOk.Move lstSP.Left + lstSP.Width + L_PAD, pvnCost.Top + pvnCost.Height + T_PAD
  ssCancel.Move ssOk.Left, ssOk.Top + ssOk.Height + T_PAD
  
  Dim r As RECT
  r.Left = 0: r.Top = 0
  r.Right = ScaleX(ScaleWidth, vbTwips, vbPixels)
  r.Bottom = ScaleY(ScaleHeight, vbTwips, vbPixels)
  ValidateRect hWnd, r
  InvalidateRect SSPanel1.hWnd, r, 0
  
  Dim cc1 As Control
  For Each cc1 In Me.Controls
    With cc1
      If .Name <> "SSPanel1" Then
        If TypeOf cc1 Is Label Then GoTo L_NEXT_CC
        
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
  
  'DoEvents
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnassData
  
  Set m_pvNodeH = Nothing
  Set m_pvNodeM = Nothing
  Set m_pvNodeC = Nothing
  Set m_pvNodeT = Nothing
    
End Sub

Private Sub lstSP_DragDrop(Source As Control, x As Single, y As Single)
  Dim dtDragType As TDragTypes
  dtDragType = m_dt_DragType
  If Not EndDragMode(DT_RemoveSP Or DT_AddSP) Then
     Exit Sub
  End If
  
  If dtDragType = DT_AddSP Then
    ssSPRight_Click
  End If
End Sub

Private Sub lstSP_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  DragValid Source, DT_RemoveSP Or DT_AddSP, State
End Sub

Private Sub lstSP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button And vbLeftButton Then
    If lstSP.SelCount > 0 Then
      lstSP.DragIcon = RemoveIcon.DragIcon
      BeginDragMode lstSP, DT_RemoveSP
    End If
  End If
End Sub

Private Sub lstSPAll_DragDrop(Source As Control, x As Single, y As Single)
  Dim dtDragType As TDragTypes
  dtDragType = m_dt_DragType
  If Not EndDragMode(DT_RemoveSP Or DT_AddSP) Then
     Exit Sub
  End If
  
  If dtDragType = DT_RemoveSP Then
    ssSPLeft_Click
  End If
End Sub

Private Sub lstSPAll_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  DragValid Source, DT_RemoveSP Or DT_AddSP, State
End Sub

Private Sub lstSPAll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button And vbLeftButton Then
    If lstSPAll.SelCount > 0 Then
      lstSPAll.DragIcon = AddIcon.DragIcon
      BeginDragMode lstSPAll, DT_AddSP
    End If
  End If
End Sub

Private Sub pvnCost_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub pvnCost_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub


Private Sub ssCancel_Click()
  m_b_Res = False
  
  If tdbgFac.EditActive Then _
    tdbgFac.CurrentCellModified = False: tdbgFac.EditActive = False
    
  Hide
End Sub


Private Sub ssCancel_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssCancel_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub ssCancel_LostFocus()
  HighLight2 ssCancel, False, Me.hWnd
End Sub

Private Sub ssCancel_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, True, Me.hWnd
End Sub

Private Sub ssCancel_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCancel, False, Me.hWnd
End Sub


Private Sub ssFacLeft_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssFacLeft_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub ssFacLeft_LostFocus()
  HighLight2 ssFacLeft, False, Me.hWnd
End Sub

Private Sub ssFacLeft_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssFacLeft, True, Me.hWnd
End Sub

Private Sub ssFacLeft_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssFacLeft, False, Me.hWnd
End Sub


Private Sub ssFacRight_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssFacRight_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub ssFacRight_LostFocus()
  HighLight2 ssFacRight, False, Me.hWnd
End Sub

Private Sub ssFacRight_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssFacRight, True, Me.hWnd
End Sub

Private Sub ssFacRight_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssFacRight, False, Me.hWnd
End Sub


Private Sub ssOk_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssOk_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub sspAllFac_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub sspAllFac_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'  DragValid Source, DT_None, State
'End Sub

Private Sub sspAllSP_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub sspAllSP_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'  DragValid Source, DT_None, State
'End Sub

Private Sub SSPanel1_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub SSPanel1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub SSPanel2_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub SSPanel2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'  DragValid Source, DT_None, State
'End Sub

Private Sub sspCost_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub sspCost_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'  DragValid Source, DT_None, State
'End Sub

Private Sub ssSPLeft_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssSPLeft_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub ssSPRight_DragDrop(Source As Control, x As Single, y As Single)
  EndDragMode DT_None
End Sub

'Private Sub ssSPRight_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'    DragValid Source, DT_None, State
'End Sub

Private Sub ssSPRight_LostFocus()
  HighLight2 ssSPRight, False, Me.hWnd
End Sub

Private Sub ssSPRight_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssSPRight, True, Me.hWnd
End Sub

Private Sub ssSPRight_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssSPRight, False, Me.hWnd
End Sub

Private Sub ssSPLeft_LostFocus()
  HighLight2 ssSPLeft, False, Me.hWnd
End Sub

Private Sub ssSPLeft_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssSPLeft, True, Me.hWnd
End Sub

Private Sub ssSPLeft_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssSPLeft, False, Me.hWnd
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



Public Sub AssData(ByVal m As MGertNet, ByVal sp As SafetyPrecaution, ByVal coll As collSF)
    
  Set HighlightedNode = Nothing
    
  Dim iClon As IClonable
  Set iClon = sp
  'Set m_sp = iClon.Clone()
  Set m_spOriginal = sp
  Set m_gn = m
  Set m_coll_SF = coll
  
  m_b_Res = False
  Caption = "Редактировать меру '" & sp.Name & "'"
      
  pvnCost.Value = sp.Cost
      
  lstSPAll.Clear
  
  Dim spSP As SafetyPrecaution
  Dim larrIdx() As Long, l As Long
  larrIdx = sp.NonCompatibles
  
  For Each spSP In coll
    If spSP.Name <> sp.Name And Not FindL(spSP.ID, larrIdx) Then
      Dim ilk As ILongKey
      Set ilk = spSP
      lstSPAll.AddItem spSP.Name
      lstSPAll.ItemData(lstSPAll.NewIndex) = ilk.Get()
    End If
  Next spSP
  
  lstSP.Clear
  
  For l = LBound(larrIdx) To UBound(larrIdx)
    Set ilk = coll(larrIdx(l))
    lstSP.AddItem coll(larrIdx(l)).Name
    lstSP.ItemData(lstSP.NewIndex) = ilk.Get()
  Next l
  
  Set m_dba_x = New XArrayDB
  m_dba_x.ReDim 0, sp.FChanges.Count - 1, 0, 2
  m_dba_x.DefaultColumnType(0) = XTYPE_STRING
  m_dba_x.DefaultColumnType(1) = XTYPE_LONG
  m_dba_x.DefaultColumnType(2) = XTYPE_LONG
  'm_dba_x.DefaultColumnType(2) = XTYPE_LONG
  
  Dim fc As FChange
  l = 0
  For Each fc In sp.FChanges
    m_dba_x(l, 0) = fc.NameFactor
    m_dba_x(l, 1) = fc.Delta
    m_dba_x(l, 2) = CLng(fc.tch)
    l = l + 1
  Next fc
      
  Set tdbgFac.Array = m_dba_x
  tdbgFac.ReBind
  
  On Error GoTo ERR_H
  tvFacAll.Redraw = False
  TvClearGl tvFacAll
  
  Dim fac As Factor
  For Each fac In m_gn.Factors
    Dim ibsKey As IBSTRKey
    Set ibsKey = fac
    For Each fc In sp.FChanges
      If fc.NameFactor = ibsKey.Get() Then GoTo L_NEXT_FAC
    Next fc
    
    AddToTree fac
    
L_NEXT_FAC:
  Next fac
  
  
  ExpandAll tvFacAll
  tvFacAll.Redraw = True
  Exit Sub
ERR_H:
  tvFacAll.Redraw = True
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub ssFacRight_Click()
  Dim lCntSel As Long
  lCntSel = CountSelected(tvFacAll)
  If lCntSel < 1 Then
    MsgBox "Чтобы добавить факторы опасности нужно их выделить в списке мышью (удерживая Shift или Control)"
    Exit Sub
  End If
  
  Set HighlightedNode = Nothing
  
  Dim node As PVBranch, node2 As PVBranch, coll As New CollectionX

  tdbgFac.Close False
  Dim lIdx As Long
  lIdx = m_dba_x.Count(1)
  m_dba_x.AppendRows lCntSel
  

  Set node = tvFacAll.Branches.Get(pvtGetChild, 0)
  While node.IsValid
    Set node2 = node.Get(pvtGetChild, 0)
    While node2.IsValid
      If node2.IsSelected Then
        coll.Add node2
        m_dba_x(lIdx, 0) = node2.DataVariant
        m_dba_x(lIdx, 1) = 1
        m_dba_x(lIdx, 2) = 0
        'm_dba_x(lIdx, 2) =
        lIdx = lIdx + 1
      End If
      Set node2 = node2.Get(pvtGetNextSibling, 0)
    Wend
    Set node = node.Get(pvtGetNextSibling, 0)
  Wend
  
  tdbgFac.ReBind
  
  On Error GoTo ERR_H
  
  tvFacAll.Redraw = False
  For Each node In coll
    node.Remove
  Next node
  
  ExpandAll tvFacAll
  tvFacAll.Redraw = True
  Exit Sub
ERR_H:
  tvFacAll.Redraw = True
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub

Private Sub ssFacLeft_Click()
  If tdbgFac.SelBookmarks.Count < 1 Then
    MsgBox "Чтобы удалить факторы опасности нужно их выделить в сетке мышью (удерживая Shift или Control)"
    Exit Sub
  End If
  
  tvFacAll.Redraw = False
  
  Dim vvt, coll As New CollectionX
  For Each vvt In tdbgFac.SelBookmarks
    coll.Add CLng(vvt)
    AddToTree m_gn.Factors(m_dba_x(vvt, 0))
  Next vvt

  ExpandAll tvFacAll
  tvFacAll.Redraw = True

  SortCollX coll
  tdbgFac.Close False
  For Each vvt In coll
    m_dba_x.DeleteRows vvt
  Next vvt
  
  tdbgFac.ReBind
  Exit Sub
L_ERRH:
  tvFacAll.Redraw = True
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Sub AddToTree(ByVal fac As Factor)

  Dim ibsKey As IBSTRKey
  Set ibsKey = fac

  Dim pvBr As PVBranch
  Select Case UCase(Left(ibsKey.Get(), 1))
    Case "H"
      Set pvBr = m_pvNodeH
    Case "M"
      Set pvBr = m_pvNodeM
    Case "C"
      Set pvBr = m_pvNodeC
    Case "T"
      Set pvBr = m_pvNodeT
  End Select

  
  Set pvBr = pvBr.Add(pvtPositionInOrder, 0, ibsKey.Get() & "  '" & fac.Name & "'")
  pvBr.ForeColor = RGB(108, 245, 250)
  pvBr.DataVariant = ibsKey.Get()

End Sub

Private Sub ssSPLeft_Click()
  
  If lstSP.SelCount < 1 Then
    MsgBox "Чтобы удалить несовместимые меры нужно их выделить в списке мышью (удерживая Shift или Control)"
    Exit Sub
  End If
  
  MoveListSelCtx lstSP, lstSPAll
End Sub

Friend Sub MoveListSelCtx(ByVal lst1 As ListBox, ByVal lst2 As ListBox)
  Dim l As Long, coll As New CollectionX
  For l = 0 To lst1.ListCount - 1
    If lst1.Selected(l) Then
      coll.Add l
      lst2.AddItem lst1.List(l)
      lst2.ItemData(lst2.NewIndex) = lst1.ItemData(l)
    End If
  Next l
  
  SortCollX coll
  Dim vvt
  For Each vvt In coll
    lst1.RemoveItem CInt(vvt)
  Next vvt
End Sub

Private Sub ssSPRight_Click()
  If lstSPAll.SelCount < 1 Then
    MsgBox "Чтобы добавить несовместимые меры нужно их выделить в списке мышью (удерживая Shift или Control)"
    Exit Sub
  End If
  
  MoveListSelCtx lstSPAll, lstSP
End Sub

Public Sub UnassData()
  tdbgFac.Close True
  
  'Set m_sp = Nothing
  Set m_spOriginal = Nothing
  Set m_gn = Nothing
  Set m_coll_SF = Nothing
    
  If Not m_dba_x Is Nothing Then
    If m_dba_x.Count(1) > 0 Then _
      m_dba_x.DeleteRows 0, m_dba_x.Count(1)
  End If
  Set m_dba_x = Nothing
      
  Set HighlightedNode = Nothing
  tvFacAll.Redraw = False
  TvClearGl tvFacAll
  tvFacAll.Redraw = True
End Sub

Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property

Private Sub ssOk_Click()
    
  Dim bm As New CBeam
  bm.Beam Me
    
  On Error Resume Next
  'If tdbgFac.EditActive Or tdbgFac.CurrentCellModified Then _

  tdbgFac.Update
    
  If Err.Number <> 0 Then
    MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка ввода"
    Exit Sub
  End If
  On Error GoTo 0
  
  If SynchronizeData Then Hide
  m_b_Res = True
End Sub

Private Function TranslateIntToTCH(ByVal l As Long) As TrustChange
  Select Case l
    Case 0
      TranslateIntToTCH = TR_NoChange
    Case 1
      TranslateIntToTCH = TR_MinusOne
    Case 2
      TranslateIntToTCH = TR_MinusTwo
    Case 3
      TranslateIntToTCH = TR_PlusOne
    Case 4
      TranslateIntToTCH = TR_PlusTwo
    Case 5
      TranslateIntToTCH = TR_SetLow
    Case 6
      TranslateIntToTCH = TR_SetNormal
    Case 7
      TranslateIntToTCH = TR_SetHigh
    Case Else
      TranslateIntToTCH = TR_NoChange
  End Select
End Function

Private Function SynchronizeData() As Boolean
  SynchronizeData = False
  
  Dim larrIdx() As Long, l As Long
  
  For l = m_dba_x.LowerBound(1) To m_dba_x.UpperBound(1)
    If m_dba_x(l, 1) < -10 Or m_dba_x(l, 1) > 10 Then
      MsgBox "Дельта оценки для фактора '" & m_dba_x(l, 0) & "' выходит за пределы [-10,10]", vbExclamation Or vbOKOnly, "Ошибка"
      Exit Function
    End If
    
    If m_dba_x(l, 2) < 0 Or m_dba_x(l, 2) > 7 Then
      MsgBox "Изменение коэффициента доверия для фактора '" & m_dba_x(l, 0) & "' имеет недопустимое значение", vbExclamation Or vbOKOnly, "Ошибка"
      Exit Function
    End If
  Next l
  
  
  m_spOriginal.Cost = pvnCost.Value
    
  If lstSP.ListCount > 0 Then
    ReDim larrIdx(lstSP.ListCount - 1) As Long
    For l = LBound(larrIdx) To UBound(larrIdx)
      larrIdx(l) = lstSP.ItemData(l)
    Next
  End If
  
  m_spOriginal.NonCompatibles = larrIdx
      
  lstSPAll.Clear
  lstSP.Clear
    
  m_spOriginal.FChanges.Clear
  Dim lIdx As Long
  For l = m_dba_x.LowerBound(1) To m_dba_x.UpperBound(1)
    lIdx = m_spOriginal.FChanges.Add(New FChange)
    m_spOriginal.FChanges(lIdx).NameFactor = m_dba_x(l, 0)
    m_spOriginal.FChanges(lIdx).Delta = m_dba_x(l, 1)
    m_spOriginal.FChanges(lIdx).tch = TranslateIntToTCH(m_dba_x(l, 2))
  Next l
  
  SynchronizeData = True
    
End Function

Friend Sub TvClearGl(ByVal pv As PVTreeViewX)

  Dim pvNodeMain As PVBranch
  Set pvNodeMain = pv.Branches.Get(pvtGetChild, 0)
  
  Do While pvNodeMain.IsValid
    pvNodeMain.Clear
    Set pvNodeMain = pvNodeMain.Get(pvtGetNextSibling, 0)
  Loop
    
End Sub


Private Sub tdbgFac_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
  If tdbgFac.SelBookmarks.Count > 0 Then
    tdbgFac.DragIcon = RemoveIcon.DragIcon
    BeginDragMode tdbgFac, DT_RemoveFac
  End If
End Sub

Private Sub tdbgFac_DragDrop(Source As Control, x As Single, y As Single)
  Dim dtDragType As TDragTypes
  dtDragType = m_dt_DragType
  If Not EndDragMode(DT_RemoveFac Or DT_AddFac) Then
     Exit Sub
  End If
  
  If dtDragType = DT_AddFac Then
    ssFacRight_Click
  End If
End Sub

Private Sub tdbgFac_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  DragValid Source, DT_RemoveFac Or DT_AddFac, State
End Sub

Private Sub tdbgFac_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueOleDBGrid60.StyleDisp)
  If ColIndex <> 0 Then
    CellTip = vbNullString
    Exit Sub
  End If
  
  If Not IsNull(tdbgFac.Bookmark) Then
    CellTip = m_gn.Factors(m_dba_x(tdbgFac.FirstRow + RowIndex, 0)).Name
  End If
End Sub

Private Sub tdbgFac_Validate(Cancel As Boolean)
   On Error Resume Next
   tdbgFac.Update
   If Err.Number <> 0 Then Cancel = True
End Sub

Private Sub tdbgFac_LostFocus()
    
  On Error Resume Next
  'If tdbgCompl.EditActive Or tdbgCompl.CurrentCellModified Then
  If tdbgFac.CurrentCellModified Then
    tdbgFac.Update
    If Err.Number <> 0 Then
      tdbgFac.CurrentCellModified = False
      tdbgFac.EditActive = False
    End If
  End If
    
End Sub


Private Sub tvFacAll_DragDrop(Source As Control, x As Single, y As Single)
  Dim dtDragType As TDragTypes
  dtDragType = m_dt_DragType
  If Not EndDragMode(DT_RemoveFac Or DT_AddFac) Then
     Exit Sub
  End If
  
  If dtDragType = DT_RemoveFac Then
    ssFacLeft_Click
  End If
End Sub

Private Sub tvFacAll_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  DragValid Source, DT_RemoveFac Or DT_AddFac, State
End Sub

Private Sub tvFacAll_LButtonDown(ByVal node As PVTreeView3Lib.PVIBranch, ByVal x As Single, ByVal y As Single)
  
  Dim lCntSel As Long
  lCntSel = CountSelected(tvFacAll)
  
  If lCntSel > 0 Then
    tvFacAll.DragIcon = AddIcon.DragIcon
    BeginDragMode tvFacAll, DT_AddFac
  End If
  
End Sub

Private Sub tvFacAll_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Dim myFont As StdFont
    
  If Not HighlightedNode Is Nothing Then
      Set HighlightedNode.Font = Nothing
  End If
  
  If Shift = 0 Then
      Set HighlightedNode = tvFacAll.HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
      If HighlightedNode.IsValid Then
          'Set myFont = tvFacAll.CreateFont
          'myFont.Bold = True
          Set HighlightedNode.Font = m_myFont
      End If
  End If
End Sub
