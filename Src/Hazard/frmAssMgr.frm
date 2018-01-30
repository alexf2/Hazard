VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{DB625B91-05EA-11D2-97DD-00400520799C}#6.0#0"; "pvtreex.ocx"
Begin VB.Form frmAssMgr 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Назначить модуль оценки значений ФО"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   Icon            =   "frmAssMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   6120
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   10795
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin PVTreeView3Lib.PVTreeViewX tv 
         Height          =   2610
         Left            =   135
         TabIndex        =   0
         Top             =   330
         Width           =   5925
         _Version        =   393216
         _ExtentX        =   10451
         _ExtentY        =   4604
         _StockProps     =   237
         ForeColor       =   0
         BackColor       =   9450828
         BorderStyle     =   1
         Appearance      =   1
         BorderStyle     =   1
         StandardDefaultPicture=   10
         Sort            =   -1  'True
         AlwaysShowSelection=   -1  'True
         AllowInPlaceEditing=   0   'False
         BackColor       =   9450828
         ForeColor       =   0
         EnableToolTips  =   0   'False
         SelectedTextBackColor=   16711680
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
      Begin VB.TextBox txtInPlace 
         BackColor       =   &H00FFFBF0&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1425
         TabIndex        =   15
         Text            =   "Text1"
         ToolTipText     =   "Вызов редактора - F2"
         Top             =   5685
         Visible         =   0   'False
         Width           =   2250
      End
      Begin Threed.SSPanel sspAllSP 
         Height          =   210
         Left            =   165
         TabIndex        =   11
         Top             =   90
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Модули оценки значений ФО"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1290
         Left            =   3315
         TabIndex        =   12
         Top             =   3045
         Visible         =   0   'False
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   2275
         _Version        =   131074
         BackColor       =   12632256
         BackStyle       =   1
         Caption         =   "Модули оценки ФО"
         ClipControls    =   0   'False
         Begin Threed.SSCommand ssAbout 
            Height          =   330
            Left            =   150
            TabIndex        =   6
            ToolTipText     =   "Вызывает окно 'About' выбранного модуля"
            Top             =   795
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":000C
            Caption         =   "Информация"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand ssRemoveAll 
            Height          =   330
            Left            =   150
            TabIndex        =   5
            ToolTipText     =   "Удаляет неустановленные модули и конфигурации"
            Top             =   315
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   582
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":011E
            Caption         =   "Удалить неустановленные"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1290
         Left            =   165
         TabIndex        =   13
         Top             =   3045
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   2275
         _Version        =   131074
         BackColor       =   12632256
         BackStyle       =   1
         Caption         =   "Конфигурации"
         ClipControls    =   0   'False
         Begin Threed.SSCommand ssRemove 
            Height          =   330
            Left            =   135
            TabIndex        =   3
            ToolTipText     =   "Удаляет выбранную конфигурацию"
            Top             =   795
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":0294
            Caption         =   "Удалить"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand ssCreate 
            Height          =   330
            Left            =   135
            TabIndex        =   1
            ToolTipText     =   "Создаёт новую конфигурацию для выбранного модуля"
            Top             =   330
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":03F2
            Caption         =   "Создать"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand ssLink 
            Height          =   360
            Left            =   1275
            TabIndex        =   2
            ToolTipText     =   "Редактировать связь конфигурации с файлом"
            Top             =   330
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   635
            _Version        =   131074
            PictureMaskColor=   255
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":0550
            Caption         =   "Изменить связь"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand ssEdit 
            Height          =   330
            Left            =   1275
            TabIndex        =   4
            ToolTipText     =   "Редактировать конфигурацию"
            Top             =   795
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   582
            _Version        =   131074
            PictureFrames   =   1
            BackStyle       =   1
            Enabled         =   0   'False
            PictureUseMask  =   -1  'True
            Picture         =   "frmAssMgr.frx":06AE
            Caption         =   "Редактировать"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   345
         Top             =   5625
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "cfg"
         Filter          =   "Все файлы (*.*)|*.*"
      End
      Begin Threed.SSPanel ssSelectedMod 
         Height          =   270
         Left            =   1815
         TabIndex        =   16
         Top             =   5850
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   476
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
         Caption         =   "Модуль не выбран"
         BevelOuter      =   0
      End
      Begin Threed.SSCommand ssAdvanced 
         Height          =   330
         Left            =   4185
         TabIndex        =   14
         ToolTipText     =   "Режим редактора конфигураций"
         Top             =   5625
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmAssMgr.frx":07C0
         Caption         =   "Редактор  >>"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOkEdit 
         Height          =   525
         Left            =   915
         TabIndex        =   7
         ToolTipText     =   "Сохраняет изменения конфигураций"
         Top             =   4395
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmAssMgr.frx":08D2
         Caption         =   "Сохранить изменения конфигураций"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOkModule 
         Default         =   -1  'True
         Height          =   525
         Left            =   165
         TabIndex        =   8
         ToolTipText     =   "Подтверждает выбор модуля ФО и конфигурации"
         Top             =   5130
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   926
         _Version        =   131074
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         Enabled         =   0   'False
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
         Picture         =   "frmAssMgr.frx":0A3C
         Caption         =   "Подтвердить выбор модуля и конфигурации"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4185
         TabIndex        =   9
         ToolTipText     =   "Закрывает окно. Если модификации не подтверждены, то они не выполняются."
         Top             =   5025
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmAssMgr.frx":0BA6
         Caption         =   "Закрыть"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin VB.Image imgCfgNotInstalled 
         Height          =   270
         Left            =   480
         Picture         =   "frmAssMgr.frx":0D10
         Top             =   4440
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgCfg 
         Height          =   270
         Left            =   465
         Picture         =   "frmAssMgr.frx":0E6A
         Top             =   4770
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgNotInstalled 
         Height          =   270
         Left            =   105
         Picture         =   "frmAssMgr.frx":0FC4
         Top             =   4770
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgModule 
         Height          =   270
         Left            =   105
         Picture         =   "frmAssMgr.frx":111E
         Top             =   4440
         Visible         =   0   'False
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmAssMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_LockResize As Boolean
Private m_b_Res As Boolean

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 400
Private Const L_PAD2 = 400

Public m_g As MGertNet

'Private m_sArr() As String

Private m_collPlugins As CollectionX
Private m_bAdvLock As Boolean

Private HighlightedNode As PVBranch
Private m_pvEdit As PVBranch
Private m_myFont As StdFont

'Public Property Get ModName() As String
'  ModName = m_sModName
'End Property
'
'Public Property Let ModName(ByVal sName As String)
'  m_sModName = sName
'  m_b_Res = False
'
'  Set Font = List1.Font
'
'  List1.Clear
'  Dim saTmp() As String
'  m_sArr = saTmp
'
'  m_sArr = GetProgIDs()
'  Dim sWidth As Single: sWidth = 0
'  Dim l As Long, lIdx As Long
'  lIdx = -1
'  For l = LBound(m_sArr) To UBound(m_sArr)
'    Dim ifass As IFactorAssign, sExtName As String
'
'    On Error Resume Next
'    Set ifass = CreateObject(m_sArr(l))
'    If Err.Number = 0 Then
'      sExtName = ifass.Description
'    Else
'      sExtName = "?"
'    End If
'    On Error GoTo 0
'
'    Set ifass = Nothing
'
'    Dim sOut As String
'    sOut = sExtName & " (" & m_sArr(l) & ")"
'    List1.AddItem sOut
'    List1.ItemData(List1.NewIndex) = l
'
'    Dim sW As Single
'    sW = TextWidth(sOut)
'    If sW > sWidth Then sWidth = sW
'
'    If lIdx = -1 Then
'      If StrComp(m_sArr(l), m_sModName, vbTextCompare) = 0 Then lIdx = l
'    End If
'  Next l
'
'  If List1.ListCount > 0 And lIdx <> -1 Then _
'    List1.Selected(lIdx) = True
'
'  SendMessage List1.hWnd, LB_SETHORIZONTALEXTENT, ScaleX(sWidth, Me.ScaleMode, vbPixels), 0
'
'End Property

Friend Sub AssData(ByVal g As MGertNet)
  Set m_g = g
  Set m_collPlugins = New CollectionX
  
  With tv
    If Trim(g.ModuleProgID) = "" Then
      '.ToolTipText = "Модуль не назначен"
      ssSelectedMod.Caption = "Модуль не назначен"
    ElseIf Trim(g.ModuleConfig) = "" Then
      '.ToolTipText = "Конфигурация не выбрана"
      ssSelectedMod.Caption = "Конфигурация не выбрана"
    Else
      '.ToolTipText = "Назначен модуль '" & g.ModuleProgID & "' с конфигурацией '" & g.ModuleConfig
      ssSelectedMod.Caption = "Назначен модуль '" & g.ModuleProgID & "' с конфигурацией '" & g.ModuleConfig
    End If
  End With
  
  Set HighlightedNode = Nothing
  
  With m_collPlugins
    .DefaultItem = Nothing
    .AutoDefaultGetKey = True
    
      
    LoadFromRegistry
  
    Dim saTmp() As String, l As Long
    saTmp = GetProgIDs()
    
    For l = LBound(saTmp) To UBound(saTmp)
      If .Item(saTmp(l)) Is .DefaultItem Then _
        AddPlugin saTmp(l), ""
    Next l
  End With
  
  ssRemoveAll.Enabled = (CheckMounted <> 0)
    
  ResolveChanges
  LoadToTV
  TvSelectCurrCfg
End Sub

Private Sub TvSelectCurrCfg()
  Dim pvb As PVBranch, sProgId As String, sNameCfg As String
  Dim pvb2 As PVBranch, pvTmp As PVBranch
  
  sProgId = m_g.ModuleProgID: sNameCfg = m_g.ModuleConfig
  If Trim(sProgId) <> "" And tv.Count > 0 Then
    Set pvb = tv.Branches.Get(pvtGetChild, 0)
    Do While pvb.IsValid
      If StrComp(pvb.DataVariant.ProgId, sProgId, vbTextCompare) = 0 Then
        If Trim(sNameCfg) <> "" Then
          Set pvb2 = pvb.Get(pvtGetChild, 0)
          Do While pvb2.IsValid
            If StrComp(pvb2.DataVariant.Name, sNameCfg, vbTextCompare) = 0 Then
              'tv.ToolTipText = "Назначен модуль '" & sProgId & "' с конфигурацией '" & sNameCfg & "'"
              ssSelectedMod.Caption = "Назначен модуль '" & sProgId & "' с конфигурацией '" & sNameCfg & "'"
              pvb2.EnsureVisible
              pvb2.Select pvtSelectNode
              Exit Sub
            End If
            Set pvb2 = pvb2.Get(pvtGetNextSibling, 0)
          Loop
        Else
          'tv.ToolTipText = "Назначен модуль '" & sProgId & "'"
          ssSelectedMod.Caption = "Назначен модуль '" & sProgId & "'"
          pvb.EnsureVisible
          pvb.Select pvtSelectNode
          Exit Sub
        End If
      End If
      Set pvb = pvb.Get(pvtGetNextSibling, 0)
    Loop
  End If
  
  If tv.Count > 0 Then
    Set pvb = tv.Branches.Get(pvtGetChild, 0)
    'tv.ToolTipText = "Модуль не назначен"
    ssSelectedMod.Caption = "Модуль не назначен"
    'm_g.ModuleProgID = "": m_g.ModuleConfig = ""
    pvb.EnsureVisible
    pvb.Select pvtSelectNode
  End If
End Sub

Private Sub LoadToTV()
  With tv
    If .Count > 0 Then .Branches.Clear
    Dim pvNode As PVBranch
    
    On Error GoTo ERR_H
    .Redraw = False
    
    Dim cpe As CPluginEntry, pce As CPlgCfgEntry
    For Each cpe In m_collPlugins
      If cpe.Dirty <> Ad_Deleted Then
        Set pvNode = .Branches.Add(pvtPositionInOrder, 0, cpe.Description & " (" & cpe.ProgId & ")")
        pvNode.ForeColor = RGB(237, 232, 83)
        If cpe.Mounted Then _
          Set pvNode.CustomItemPicture = imgModule.Picture _
        Else Set pvNode.CustomItemPicture = imgNotInstalled.Picture
        Set pvNode.DataVariant = cpe
        
        For Each pce In cpe.Configs
          If pce.Dirty <> Ad_Deleted Then _
            AddCfgToTV pvNode, pce
        Next pce
      End If
    Next cpe
    
    ExpandAll tv
    .Redraw = True
    
  End With
  
  Exit Sub
ERR_H:
  tv.Redraw = True
  Err.Raise Err.Number
End Sub

Private Sub AddCfgToTV(ByVal pvb As PVBranch, ByVal cpe As CPlgCfgEntry)
  Dim pvNode As PVBranch
  
  Set pvNode = pvb.Add(pvtPositionInOrder, 0, cpe.Name)
  With pvNode
    .ForeColor = RGB(14, 239, 61)
    
    If cpe.Mounted Then _
      Set .CustomItemPicture = imgCfg.Picture _
    Else Set .CustomItemPicture = imgCfgNotInstalled.Picture
    Set .DataVariant = cpe
    
    Set pvNode = .Add(pvtPositionInOrder, 0, cpe.FileName)
    pvNode.ForeColor = RGB(108, 245, 250)
  End With
End Sub

Private Function CheckMounted() As Long
  Dim pl As CPluginEntry
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  Dim lCntNotMounted As Long
  lCntNotMounted = 0
  
  For Each pl In m_collPlugins
  
    Dim ifass As IFactorAssign, sExtName As String

    With pl
      Err.Clear
      On Error Resume Next
      Set ifass = CreateObject(.ProgId)
      If Err.Number = 0 Then
        .Mounted = True
        sExtName = ifass.Description
        If StrComp(sExtName, .Description, vbTextCompare) <> 0 Then
          .Description = sExtName
          .Modify
        End If
      Else
        .Mounted = False
        lCntNotMounted = lCntNotMounted + 1
        sExtName = "?"
        If .Description = "" Then .Description = sExtName
      End If
      On Error GoTo 0
    
      Set ifass = Nothing
      
      lCntNotMounted = lCntNotMounted + .CheckMounted
    End With

  Next pl
  
  CheckMounted = lCntNotMounted
End Function

Friend Sub UnassData()
  Set m_g = Nothing
  Set m_collPlugins = Nothing
  Set HighlightedNode = Nothing
  Set m_pvEdit = Nothing
  tv.Branches.Clear
End Sub


Private Sub Form_Initialize()
  m_b_LockResize = False
  m_bAdvLock = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssCancel_Click
  ElseIf KeyAscii = vbKeyReturn Then
    If ssOkModule.Enabled Then _
      ssOkModule_Click
  End If
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
  Set m_myFont = tv.CreateFont
  m_myFont.Bold = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then _
    Cancel = 1: txtInPlace_LostFocus: ssCancel_Click
End Sub

Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  Err.Clear
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 6500 Then Width = 6500
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
  
  sspAllSP.Move L_PAD, T_PAD
    
  If SSFrame1.Visible Then
    tv.Move sspAllSP.Left, sspAllSP.Top + sspAllSP.Height, SSPanel1.Width - 2 * L_PAD, SSPanel1.Height - 5 * T_PAD - sspAllSP.Height - SSFrame1.Height - ssOkEdit.Height - ssOkModule.Height - ssSelectedMod.Height
    ssSelectedMod.Move tv.Left, tv.Top + tv.Height + T_PAD / 2, tv.Width
    SSFrame1.Move tv.Left, ssSelectedMod.Top + ssSelectedMod.Height + T_PAD / 2
    SSFrame2.Move tv.Left + tv.Width - SSFrame2.Width, SSFrame1.Top
    ssOkEdit.Move SSFrame1.Left, SSFrame1.Top + SSFrame1.Height + T_PAD, tv.Width
    ssOkModule.Move tv.Left, ssOkEdit.Top + ssOkEdit.Height + T_PAD, tv.Width - ssCancel.Width - L_PAD
    ssCancel.Move ssOkModule.Left + ssOkModule.Width, ssOkModule.Top
  Else
    tv.Move sspAllSP.Left, sspAllSP.Top + sspAllSP.Height, SSPanel1.Width - 2 * L_PAD, SSPanel1.Height - 4 * T_PAD - ssOkModule.Height - ssAdvanced.Height - sspAllSP.Height - ssSelectedMod.Height
    ssSelectedMod.Move tv.Left, tv.Top + tv.Height + T_PAD / 2, tv.Width
    ssOkModule.Move tv.Left, ssSelectedMod.Top + ssSelectedMod.Height + T_PAD / 2, tv.Width - L_PAD - ssCancel.Width
    ssCancel.Move ssOkModule.Left + ssOkModule.Width + L_PAD, ssOkModule.Top
    ssAdvanced.Move ssCancel.Left, ssCancel.Top + ssCancel.Height + T_PAD, ssCancel.Width
  End If
  
'  Dim r As RECT
'  r.Left = 0
'  r.Top = 0
'  r.Right = ScaleX(ScaleWidth, vbTwips, vbPixels)
'  r.Bottom = ScaleY(ScaleHeight, vbTwips, vbPixels)
'  ValidateRect hWnd, r
'  ValidateRect SSPanel1.hWnd, r
'  ValidateRect tv.hWnd, r
  
  'DoEvents
  
'  r.Left = 0
'  r.Top = 0
'  r.Right = ScaleX(tv.Width, vbTwips, vbPixels)
'  r.Bottom = ScaleY(tv.Height, vbTwips, vbPixels)
'  InvalidateRect tv.hWnd, r, 0
  
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ssAdvanced_Click()
  If m_bAdvLock Then Exit Sub
  m_bAdvLock = True
  On Error Resume Next
  ssAdvanced.Visible = False
  SSFrame1.Visible = True
  SSFrame2.Visible = True
  ssOkEdit.Visible = True
  Form_Resize
  m_bAdvLock = False
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


Private Sub ssAbout_Click()
  Dim pvNode As PVBranch
  Set pvNode = tv.Branches.Get(pvtGetPrevSelected, 0)
  
  If pvNode.IsValid Then
    Dim ifass As IFactorAssign
    On Error GoTo ERR_H
    Set ifass = CreateObject(pvNode.DataVariant.ProgId)
    ifass.About Me, Me.hWnd
  Else
    MsgBox "Нужно выбрать модуль", vbInformation Or vbOKOnly, "Сообщение"
  End If
  
  Exit Sub
ERR_H:
  MsgBox "Не удаётся загрузить компонент '" & pvNode.DataVariant.ProgId & "':" & vbCrLf & Err.Description, vbOKOnly Or vbCritical, "Ошибка"
  Err.Clear
End Sub

Private Sub ssCreate_Click()
  Dim cpe As CPluginEntry
  Dim pv As PVBranch
  Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
  If pv.IsValid Then
    If pv.Level = 0 Then
      Set cpe = pv.DataVariant
    ElseIf pv.Level = 1 Then
      Set pv = pv.Get(pvtGetParent, 0)
      Set cpe = pv.DataVariant
    ElseIf pv.Level = 2 Then
      Set pv = pv.Get(pvtGetParent, 0).Get(pvtGetParent, 0)
      Set cpe = pv.DataVariant
    Else
      MsgBox "Внутренняя ошибка", vbExclamation Or vbOKOnly, "Ошибка"
      Exit Sub
    End If
    
    Dim cp As CPlgCfgEntry
    Set cp = cpe.Configs.Item("<Новая>")
    If Not cp Is cpe.Configs.DefaultItem Then
      If cp.Dirty = Ad_Deleted Then
        Dim sOKey As String
        sOKey = UniqueKey(cpe.Configs)
        cp.Name = sOKey
        cpe.Configs.Key("<Новая>") = sOKey
        'cpe.Configs.Remove "<Новая>"
      Else
        MsgBox "Слот для новой конфигурации уже есть", vbExclamation Or vbOKOnly, "Ошибка"
        Exit Sub
      End If
    End If
    Dim cpce As CPlgCfgEntry
    Set cpce = cpe.AddConfig("<Новая>", "<Нет>")
    cpce.Mounted = False
    cpce.Dirty = Ad_New
    
    
    Set pv = pv.Add(pvtPositionInOrder, 0, cpce.Name)
    pv.ForeColor = RGB(14, 239, 61)
    Set pv.DataVariant = cpce
    Set pv.CustomItemPicture = imgCfgNotInstalled.Picture
    Set pv = pv.Add(pvtPositionInOrder, 0, cpce.FileName)
    pv.ForeColor = RGB(108, 245, 250)
    pv.EnsureVisible
    Set pv = pv.Get(pvtGetParent, 0)
    pv.EnsureVisible
    pv.Select pvtSelectNode
    tv.SetFocus
    Dim iShift As Integer
    tv_KeyDown vbKeyF2, iShift
    'tv.BeginInPlaceEdit
    
  End If
End Sub

Private Sub ssCreate_LostFocus()
  HighLight2 ssCreate, False, Me.hWnd
End Sub

Private Sub ssCreate_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCreate, True, Me.hWnd
End Sub

Private Sub ssCreate_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssCreate, False, Me.hWnd
End Sub


Private Sub ssEdit_Click()
  On Error GoTo ERR_H
  Dim pvb As PVBranch, sCfgName As String, sFileName As String, sProgId As String
  Set pvb = tv.Branches.Get(pvtGetNextSelected, 0)
  With pvb
    If .IsValid Then
      If TypeOf .DataVariant Is CPlgCfgEntry Then
        sCfgName = .DataVariant.Name
        sFileName = .DataVariant.FileName
        Set pvb = pvb.Get(pvtGetParent, 0)
        sProgId = pvb.DataVariant.ProgId
      'ElseIf TypeOf pvb.DataVariant Is CPluginEntry Then
      Else
        MsgBox "Это не редактируется", vbExclamation Or vbOKOnly, "Ошибка"
        Exit Sub
      End If
      Dim ifass As IFactorAssign
      Set ifass = CreateObject(sProgId)
      ifass.Configure Me, hWnd, sCfgName, sFileName
    End If
  End With
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка модуля '" & sProgId & "'"
  Err.Clear
End Sub

Private Sub ssLink_Click()
  With CommonDialog1
    .CancelError = True
    .DialogTitle = "Выбрать файл конфигурации"
    .Flags = FileOpenConstants.cdlOFNExplorer Or FileOpenConstants.cdlOFNHideReadOnly Or FileOpenConstants.cdlOFNLongNames
    If Trim(.InitDir) = "" Then _
      .InitDir = App.Path
                
    .FileName = ""
                
    Dim pv As PVBranch, pvPar As PVBranch
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    Set pvPar = pv.Get(pvtGetParent, 0)
    If Not pv.IsValid Or Not pvPar.IsValid Then
      MsgBox "Не выбрана конфигурация", vbExclamation Or vbOKOnly, "Ошибка"
      Exit Sub
    End If
                                
    On Error GoTo ERR_H0
    .ShowOpen
    DoEvents
                                              
    Dim sDir As String
    CutPath .FileName, sDir
    .InitDir = sDir
                                          
    pvPar.DataVariant.FileName = .FileName
    pvPar.DataVariant.Modify
    pv.Text = .FileName
                                          
    UpdateMounted pvPar
                                                
    'TvSelectCurrCfg
    
    Exit Sub
ERR_H0:
  End With
  Err.Clear
  Exit Sub

End Sub

Private Sub UpdateMounted(ByVal pv As PVBranch)
  Err.Clear
  On Error Resume Next
  FileLen pv.DataVariant.FileName
  If Err.Number <> 0 Then
    Dim vbRes As VbMsgBoxResult
    vbRes = MsgBox("Файл '" & pv.DataVariant.FileName & "' - не существует. Создать ?", vbQuestion Or vbYesNo, "Вопрос")
    If vbRes = vbYes Then
      Dim stg As IStorage
      On Error GoTo ERR_H2
      Set stg = CreateCF(pv.DataVariant.FileName, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
      pv.DataVariant.Mounted = True
      Set pv.CustomItemPicture = imgCfg.Picture
    Else
      pv.DataVariant.Mounted = False
      Set pv.CustomItemPicture = imgCfgNotInstalled.Picture
    End If
  Else
    pv.DataVariant.Mounted = True
    Set pv.CustomItemPicture = imgCfg.Picture
  End If
  
  Exit Sub
  
ERR_H2:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssLink_LostFocus()
  HighLight2 ssLink, False, Me.hWnd
End Sub

Private Sub ssLink_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssLink, True, Me.hWnd
End Sub

Private Sub ssLink_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssLink, False, Me.hWnd
End Sub

Private Function CheckFileNames(sFN As String) As Boolean
  Dim cpe As CPluginEntry, cpce As CPlgCfgEntry
  For Each cpe In m_collPlugins
    For Each cpce In cpe.Configs
      If Trim(cpce.FileName) = "" Then
        CheckFileNames = False
        sFN = cpce.Name
        Exit Function
      End If
    Next cpce
  Next cpe
  CheckFileNames = True
End Function

Private Sub ssOkEdit_Click()
  On Error GoTo ERR_H
  Dim sFN As String
  If Not CheckFileNames(sFN) Then
    MsgBox "Для конфигурации '" & sFN & "' не выбран файл", vbExclamation Or vbOKOnly, "Нельзя сохранить"
    Exit Sub
  Else
    ResolveChanges
  End If
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssOkModule_Click()
  Dim pvb As PVBranch
  Set pvb = tv.Branches.Get(pvtGetNextSelected, 0)
  
  On Error GoTo ERR_H
  
  If pvb.IsValid Then
    If Not ssOkModule.Enabled Or Not pvb.Level = 1 Then Exit Sub
    
    If TypeOf pvb.DataVariant Is CPlgCfgEntry Then
      m_g.ModuleConfig = pvb.DataVariant.Name
      Set pvb = pvb.Get(pvtGetParent, 0)
      m_g.ModuleProgID = pvb.DataVariant.ProgId
      Dim sTmp As String
      sTmp = "Назначен модуль '" & pvb.DataVariant.Description & "' с конфигурацией '" & m_g.ModuleConfig & "'"
      'MsgBox sTmp, vbInformation Or vbOKOnly, "Информация"
      
      'tv.ToolTipText = sTmp
      ssSelectedMod.Caption = sTmp
      
      Exit Sub
    End If
  End If
  MsgBox "Нельзя назначить", vbExclamation Or vbOKOnly, "Ошибка"
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssRemove_Click()
  On Error GoTo ERR_H
  Dim pvb As PVBranch
  Set pvb = tv.Branches.Get(pvtGetNextSelected, 0)
  If pvb.IsValid Then
    If pvb.Level <> 2 Then
      If pvb.Level = 0 Then If pvb.DataVariant.Mounted Then Exit Sub
      
      Dim pvb2 As PVBranch
      Set pvb2 = pvb.Get(pvtGetNextSibling, 0)
      If Not pvb2.IsValid Then
        Set pvb2 = pvb.Get(pvtGetPrevSibling, 0)
      End If
            
      
      pvb.DataVariant.Dirty = Ad_Deleted
      pvb.Remove
      
      If pvb2.IsValid Then pvb2.Select pvtSelectNode _
      Else TvSelectCurrCfg
    End If
  End If
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssRemove_LostFocus()
  HighLight2 ssRemove, False, Me.hWnd
End Sub

Private Sub ssRemove_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemove, True, Me.hWnd
End Sub

Private Sub ssRemove_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemove, False, Me.hWnd
End Sub


Private Sub ssEdit_LostFocus()
  HighLight2 ssEdit, False, Me.hWnd
End Sub

Private Sub ssEdit_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssEdit, True, Me.hWnd
End Sub

Private Sub ssEdit_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssEdit, False, Me.hWnd
End Sub


Private Sub ssRemoveAll_Click()
  On Error GoTo ERR_H
  
  Dim cpe As CPluginEntry, cpce As CPlgCfgEntry
  For Each cpe In m_collPlugins
    With cpe
      If Not .Mounted Then
        'm_collPlugins.Remove .ProgId
        .Dirty = Ad_Deleted
      Else
        For Each cpce In .Configs
          If Not cpce.Mounted Then _
            cpce.Dirty = Ad_Deleted
        Next cpce
      End If
    End With
  Next cpe
  
  LoadToTV
  TvSelectCurrCfg
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssRemoveAll_LostFocus()
  HighLight2 ssRemoveAll, False, Me.hWnd
End Sub

Private Sub ssRemoveAll_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemoveAll, True, Me.hWnd
End Sub

Private Sub ssRemoveAll_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssRemoveAll, False, Me.hWnd
End Sub


Private Sub ssAbout_LostFocus()
  HighLight2 ssAbout, False, Me.hWnd
End Sub

Private Sub ssAbout_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAbout, True, Me.hWnd
End Sub

Private Sub ssAbout_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAbout, False, Me.hWnd
End Sub


Private Sub ssOkEdit_LostFocus()
  HighLight2 ssOkEdit, False, Me.hWnd
End Sub

Private Sub ssOkEdit_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkEdit, True, Me.hWnd
End Sub

Private Sub ssOkEdit_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkEdit, False, Me.hWnd
End Sub


Private Sub ssOkModule_LostFocus()
  HighLight2 ssOkModule, False, Me.hWnd
End Sub

Private Sub ssOkModule_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkModule, True, Me.hWnd
End Sub

Private Sub ssOkModule_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssOkModule, False, Me.hWnd
End Sub



Private Sub ssAdvanced_LostFocus()
  HighLight2 ssAdvanced, False, Me.hWnd
End Sub

Private Sub ssAdvanced_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAdvanced, True, Me.hWnd
End Sub

Private Sub ssAdvanced_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssAdvanced, False, Me.hWnd
End Sub


Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property



Private Sub Clear()
  If m_collPlugins.Count > 0 Then _
    m_collPlugins.Remove Start:=1, Count:=m_collPlugins.Count
End Sub

Private Function AddPlugin(ByVal sProgId As String, ByVal sDescription As String) As CPluginEntry
  Dim pl As New CPluginEntry, plTmp As CPluginEntry
  With pl
    .ProgId = sProgId: .Description = sDescription: .Dirty = Ad_New
  End With
  
  With m_collPlugins
    Set plTmp = .Item(sProgId)
    If Not plTmp Is .DefaultItem Then
      If plTmp.Dirty <> Ad_Deleted Then _
        Err.Raise vbObjectError + 3, "frmAssMgr.AddPlugin", "Нельзя добавить: плагин '" & sProgId & "' уже имеется"
      .Remove sProgId
    End If
    
    .Add pl, sProgId
  End With
  
  Set AddPlugin = pl
  Set pl = Nothing
End Function

Private Sub LoadFromRegistry()
  Dim lNSubKeys As Long, lMaxlen As Long, lRes As Long
  
  Dim hKey As Long, sKeyFull As String
  sKeyFull = S_RegAppKey & "Plugins" '\"
  lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyFull, 0, KEY_ALL_ACCESS, hKey)
  If lRes <> ERROR_SUCCESS Then _
    Err.Raise vbObjectError + 2, "frmAssMgr.LoadFromRegistry", "Ошибка открытия ветви: [" & sKeyFull & "] - '" & Win32ErrInfo(lRes) & "'"
  
  Clear
  
  lRes = RegQueryInfoKey(hKey, vbNullString, 0&, 0&, VarPtr(lNSubKeys), VarPtr(lMaxlen), 0&, 0&, 0&, 0&, 0&, 0&)
  If lRes = ERROR_SUCCESS Then
    Dim sTmp As String, l As Long, lLenIn As Long
        
    For l = lNSubKeys - 1 To 0 Step -1
      sTmp = String$(lMaxlen + 1, vbNullChar)
      lLenIn = lMaxlen + 1
      lRes = RegEnumKeyEx(hKey, l, sTmp, VarPtr(lLenIn), 0&, vbNullString, 0&, 0&)
      If lRes <> ERROR_SUCCESS Then
        RegCloseKey hKey
        Err.Raise vbObjectError + 1, "frmAssMgr.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
      Else
        If lLenIn > 0 Then
          If Chr$(0) = Mid$(sTmp, lLenIn, 1) Then lLenIn = lLenIn - 1
        End If
        sTmp = Left$(sTmp, lLenIn)
        
        Dim hKey2 As Long
        lRes = RegOpenKeyEx(hKey, sTmp, 0, KEY_ALL_ACCESS, hKey2)
        If lRes <> ERROR_SUCCESS Then
          RegCloseKey hKey
          Err.Raise vbObjectError + 2, "frmAssMgr.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
        Else
          Dim sTmp2 As String, lPath As Long: lPath = MAX_PATH * 2
          sTmp2 = String$(MAX_PATH * 2, vbNullChar)
          lRes = RegQueryValueEx(hKey2, "", 0&, 0&, sTmp2, lPath)
          
          If lPath > 0 Then
            If Chr$(0) = Mid$(sTmp2, lPath, 1) Then lPath = lPath - 1
          End If
          sTmp2 = Left$(sTmp2, lPath)
          
          
          If lRes <> ERROR_SUCCESS Then
            RegCloseKey hKey
            RegCloseKey hKey2
            Err.Raise vbObjectError + 2, "CPluginEntry.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
          End If
            
          On Error GoTo ERR_H
          
          Dim pl As CPluginEntry
          Set pl = New CPluginEntry
          With pl
            .ProgId = sTmp
            .OldProgId = sTmp
            .Description = sTmp2
            .Dirty = AD_None
          End With
          m_collPlugins.Add pl, sTmp
          
          pl.LoadFromRegistry hKey2
          
          On Error GoTo 0
          RegCloseKey hKey2
          Set pl = Nothing
        End If
      End If
    Next l
  Else
    RegCloseKey hKey
    Err.Raise vbObjectError + 2, "frmAssMgr.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
  End If
  
  RegCloseKey hKey
  Exit Sub
  
ERR_H:
  RegCloseKey hKey
  RegCloseKey hKey2
  Err.Raise Err.Number
End Sub

Private Sub ResolveChanges()
  Dim pl As CPluginEntry, lRes As Long, hKey As Long, hKey2 As Long
  Dim collRemov As New CollectionX
  
  Dim sKeyFull As String
  sKeyFull = S_RegAppKey & "Plugins" '\"
  lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyFull, 0, KEY_ALL_ACCESS, hKey)
  If lRes <> ERROR_SUCCESS Then _
    Err.Raise vbObjectError + 2, "frmAssMgr.ResolveChanges", "Ошибка открытия ветви: [" & sKeyFull & "] - '" & Win32ErrInfo(lRes) & "'"
  
  On Error GoTo ERR_H
  
  Dim cllTmp As CollectionX
  Set cllTmp = GetResolveReady(m_collPlugins)
  
  For Each pl In cllTmp
    With pl
      If .Dirty <> AD_None Then

        Select Case .Dirty

          Case Ad_Deleted
            If Not IsNull(.OldProgId) Then _
              If ExitsKey(hKey, .OldProgId) Then MkRemoveKey hKey, .ProgId
              
            collRemov.Add .ProgId

          Case Ad_New, Ad_Updated
            If .Dirty = Ad_New Then _
              If ExitsKey(hKey, .ProgId) Then MkRemoveKey hKey, .ProgId

            lRes = RegCreateKeyEx(hKey, .ProgId, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey2, 0&)
            If lRes <> ERROR_SUCCESS Then _
              Err.Raise vbObjectError + 2, "frmAssMgr.ResolveChanges", "Ошибка создания ключа '" & .ProgId & "' в реестре: '" & Win32ErrInfo(lRes) & "'"

            lRes = RegSetValueEx2(hKey2, "", 0&, REG_SZ, .Description, Len(.Description))
            
            If lRes <> ERROR_SUCCESS Then
              RegCloseKey hKey2
              Err.Raise vbObjectError + 2, "frmAssMgr.ResolveChanges", "Ошибка записи ключа '" & .ProgId & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
            End If
            
            RegCloseKey hKey2
            .OldProgId = .ProgId
         
        End Select
      End If
      If .Dirty <> Ad_Deleted Then
        If .Dirty2 Then
          lRes = RegCreateKeyEx(hKey, .ProgId, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey2, 0&)
          If lRes <> ERROR_SUCCESS Then _
            Err.Raise vbObjectError + 2, "frmAssMgr.ResolveChanges", "Ошибка создания ключа '" & .ProgId & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
            
          On Error Resume Next
          .ResolveChanges hKey2
          RegCloseKey hKey2
          If Err.Number <> 0 Then GoTo ERR_H
          On Error GoTo ERR_H
        End If
      End If
      .Dirty = AD_None
    End With
  Next pl
  
  RegCloseKey hKey
  RemoveCollItems m_collPlugins, collRemov
  
  Exit Sub
ERR_H:
  RegCloseKey hKey
  RemoveCollItems m_collPlugins, collRemov
  Err.Raise Err.Number
End Sub


Private Sub tv_AfterSelChange(ByVal node As PVTreeView3Lib.PVIBranch)
  With node
    If .IsValid Then
      If .Level = 2 Then
        ssLink.Enabled = True
        ssEdit.Enabled = False
        ssRemove.Enabled = False
        ssCreate.Enabled = False
        ssAbout.Enabled = False
        ssOkModule.Enabled = False
      ElseIf .Level = 1 Then
        ssLink.Enabled = False
        ssEdit.Enabled = True
        ssRemove.Enabled = True
        ssCreate.Enabled = False
        ssAbout.Enabled = False
        ssOkModule.Enabled = True
      Else
        ssLink.Enabled = False
        ssEdit.Enabled = False
        ssRemove.Enabled = IIf(.DataVariant.Mounted = True, False, True)
        ssCreate.Enabled = True
        ssAbout.Enabled = True
        ssOkModule.Enabled = False
      End If
    Else
      ssLink.Enabled = False
      ssEdit.Enabled = False
      ssRemove.Enabled = False
      ssCreate.Enabled = False
      ssAbout.Enabled = False
      ssOkModule.Enabled = False
    End If
  End With
End Sub

Private Sub tv_InPlaceEditBegin(ByVal node As PVTreeView3Lib.PVIBranch, fAllow As Boolean)
'  If node.IsValid Then
'    If node.Level <> 0 Then
'      fAllow = True
'      Set m_pvEdit = node
'    Else
'      fAllow = False
'    End If
'  End If
End Sub

Private Sub tv_InPlaceEditEnd(ByVal szText As String, fAllow As Boolean)
'  fAllow = True
'  If m_pvEdit Is Nothing Then
'    'fAllow = False
'    MsgBox "Внутренняя ошибка", vbExclamation Or vbOKOnly
'  Else
'    If m_pvEdit.Level = 2 Then
'      Dim pv As PVBranch
'      Set pv = m_pvEdit.Get(pvtGetParent, 0)
'      If Trim(szText) = "" Then
'        MsgBox "Нельзя назначить пустое имя файла", vbExclamation Or vbOKOnly, "Ошибка"
'        'fAllow = False
'        m_pvEdit.Text = pv.DataVariant.FileName
'        Set m_pvEdit = Nothing
'        Exit Sub
'      End If
'      pv.DataVariant.FileName = szText
'      pv.DataVariant.Modify
'      UpdateMounted pv
'      'fAllow = True
'    ElseIf m_pvEdit.Level = 1 Then
'      Set pv = m_pvEdit.Get(pvtGetParent, 0)
'      Dim cpe As CPluginEntry
'      Set cpe = pv.DataVariant
'      If cpe.Configs.Item(szText) Is cpe.Configs.DefaultItem Then
'        cpe.Configs.Key(m_pvEdit.DataVariant.Name) = szText
'        m_pvEdit.DataVariant.Name = szText
'        m_pvEdit.DataVariant.Modify
'        'fAllow = True
'      Else
'        'fAllow = False
'        MsgBox "Нельзя назначить имя '" & szText & "' - такое уже существует", vbExclamation Or vbOKOnly, "Ошибка"
'        SendMessageStr tv.GetEditHWND, WM_SETTEXT, 0, m_pvEdit.DataVariant.Name
'        m_pvEdit.Text = m_pvEdit.DataVariant.Name
'        'tv.CancelEvent
'      End If
'    End If
'    Set m_pvEdit = Nothing
'  End If
End Sub


Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim pv As PVBranch
  If KeyCode = vbKeyReturn Then
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    If pv.IsValid Then
      If pv.Level = 1 Then
        ssOkModule_Click
      End If
    End If
    
  ElseIf KeyCode = vbKeyDelete Then
    If ssAdvanced.Visible Then Beep Else ssRemove_Click
    
  ElseIf KeyCode = vbKeyF2 Then
    Set pv = tv.Branches.Get(pvtGetNextSelected, 0)
    If pv.IsValid Then
      If pv.Level = 0 Or ssAdvanced.Visible Then
        Beep
        Exit Sub
      End If
      
      On Error GoTo ERR_H
      
      Set m_pvEdit = pv
      
      pv.Select pvtSelectNode
      tv.BeginInPlaceEdit
      Dim r As RECT, x As Single, y As Single, w As Single, h As Single
      GetWindowRect tv.GetEditHWND(), r
      tv.EndInPlaceEdit
      
      Dim pt As POINTAPI
      pt.x = r.Left: pt.y = r.Top
      ScreenToClient SSPanel1.hWnd, pt
      x = pt.x: y = pt.y
      
      pt.x = r.Right: pt.y = r.Bottom
      ScreenToClient SSPanel1.hWnd, pt
      w = pt.x - x: h = pt.y - y
      
      x = ScaleX(x, vbPixels, vbTwips)
      w = ScaleX(w, vbPixels, vbTwips)
      y = ScaleY(y, vbPixels, vbTwips)
      h = ScaleY(h, vbPixels, vbTwips)
      
      w = tv.Width - 2 * x
      txtInPlace.Move x, y, w, h
            
      If pv.Level = 1 Then
        txtInPlace.Text = pv.DataVariant.Name
      Else
        Set pv = pv.Get(pvtGetParent, 0)
        txtInPlace.Text = pv.DataVariant.FileName
      End If
      txtInPlace.Visible = True
      txtInPlace.ZOrder 0
      txtInPlace.SetFocus
      
      
      Me.KeyPreview = False
      ssCancel.Cancel = False
      ssOkModule.Default = False
    End If
  End If
  
  Exit Sub
  
ERR_H:
  Set m_pvEdit = Nothing
  txtInPlace.Visible = False
  KeyPreview = True
  ssCancel.Cancel = True
  ssOkModule.Default = True
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Dim myFont As StdFont
    
  If Not HighlightedNode Is Nothing Then
    Set HighlightedNode.Font = Nothing
  End If
  
  With tv
    If Shift = 0 Then
      Set HighlightedNode = .HitTest(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
      If HighlightedNode.IsValid Then
          'Set myFont = .CreateFont
          'myFont.Bold = True
          Set HighlightedNode.Font = m_myFont
      End If
    End If
  End With
End Sub

Private Sub txtInPlace_KeyPress(KeyAscii As Integer)
  If m_pvEdit Is Nothing Then Exit Sub
  
  On Error GoTo ERR_H
  
  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    GoTo L_UNLOCK
    
  ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Dim szText As String: szText = Trim(txtInPlace.Text)
    If m_pvEdit.Level = 2 Then
      Dim pv As PVBranch
      Set pv = m_pvEdit.Get(pvtGetParent, 0)
      If szText = "" Then
        MsgBox "Нельзя назначить пустое имя файла", vbExclamation Or vbOKOnly, "Ошибка"
        'fAllow = False
        'm_pvEdit.Text = pv.DataVariant.FileName
      Else
        pv.DataVariant.FileName = szText
        m_pvEdit.Text = szText
        pv.DataVariant.Modify
        UpdateMounted pv
        'fAllow = True
      End If
      
      GoTo L_UNLOCK
      
    ElseIf m_pvEdit.Level = 1 Then
      Set pv = m_pvEdit.Get(pvtGetParent, 0)
      If StrComp(m_pvEdit.DataVariant.Name, szText, vbTextCompare) = 0 Then GoTo L_UNLOCK
      Dim cpe As CPluginEntry
      Set cpe = pv.DataVariant
      
      Dim cp As CPlgCfgEntry
      Set cp = cpe.Configs.Item(szText)
      If Not cp Is cpe.Configs.DefaultItem Then
        If cp.Dirty = Ad_Deleted Then
          Dim sOKey As String
          sOKey = UniqueKey(cpe.Configs)
          cp.Name = sOKey
          cpe.Configs.Key(szText) = sOKey
          'cpe.Configs.Remove szText
        Else
          MsgBox "Нельзя назначить имя '" & szText & "' - такое уже существует", vbExclamation Or vbOKOnly, "Ошибка"
          Exit Sub
        End If
      End If
      
      
      Dim sTmp As String: sTmp = m_pvEdit.DataVariant.Name
      cpe.Configs.Key(sTmp) = szText
      m_pvEdit.DataVariant.Name = szText
      m_pvEdit.DataVariant.Modify
      m_pvEdit.Text = szText
      
      Set cp = cpe.AddConfig(UniqueKey(cpe.Configs), "F*")
      cp.Dirty = Ad_Deleted: cp.OldName = sTmp
      'fAllow = True
      
      GoTo L_UNLOCK
      
    End If
  End If
  
  Exit Sub
  
L_UNLOCK:
  On Error GoTo 0
  Me.KeyPreview = True
  ssCancel.Cancel = True
  ssOkModule.Default = True
  Set m_pvEdit = Nothing
  txtInPlace.Visible = False
  tv.SetFocus
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  GoTo L_UNLOCK
End Sub

Private Sub txtInPlace_LostFocus()
  If Not m_pvEdit Is Nothing Then
    Me.KeyPreview = True
    ssCancel.Cancel = True
    ssOkModule.Default = True
    Set m_pvEdit = Nothing
    txtInPlace.Visible = False
    tv.SetFocus
  End If
End Sub

Friend Function GetPathForCfg(ByVal sPlugin As String, ByVal sCfg As String) As String

  Dim hKey As Long, sKeyFull As String, lRes As Long
  sKeyFull = S_RegAppKey & "Plugins\" & sPlugin & "\" & sCfg
  
  lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyFull, 0, KEY_ALL_ACCESS, hKey)
  If lRes <> ERROR_SUCCESS Then
    Err.Raise vbObjectError + 2, "frmAssMgr.GetPathForCfg", "Ошибка открытия ветви: [" & sKeyFull & "] - '" & Win32ErrInfo(lRes) & "'"
        
  Else
    Dim sTmp2 As String, lPath As Long: lPath = MAX_PATH * 2
    sTmp2 = String$(MAX_PATH * 2, vbNullChar)
    lRes = RegQueryValueEx(hKey, "", 0&, 0&, sTmp2, lPath)
    
    If lPath > 0 Then
      If Chr$(0) = Mid$(sTmp2, lPath, 1) Then lPath = lPath - 1
    End If
    sTmp2 = Left$(sTmp2, lPath)
    
    RegCloseKey hKey
    If lRes <> ERROR_SUCCESS Then _
      Err.Raise vbObjectError + 2, "frmAssMgr.GetPathForCfg", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
    
      
    GetPathForCfg = sTmp2
      
  End If
  
End Function
