VERSION 5.00
Object = "{7E00A3A2-8F5C-11D2-BAA4-04F205C10000}#1.0#0"; "VsVIEW6.ocx"
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "formx.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Begin VB.Form frxModelMon 
   Caption         =   "Модель"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frxModelMonitor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmStdMDIChild"
   ScaleHeight     =   4815
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FormXLib.FormX frmxChild 
      Left            =   1800
      Top             =   0
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
      ClientEdge      =   0
      UseCaptionButton=   0
      UseCaptionDblClick=   0
      CanDragDocked   =   0
      AutoFrameMDIChild=   1
      AutoFrameFloating=   0
      AutoFrameDockTop=   0
      AutoFrameDockLeft=   0
      AutoFrameDockBottom=   0
      AutoFrameDockRight=   0
   End
   Begin VSVIEW6Ctl.VSViewPort VSViewPort1 
      Height          =   3720
      Left            =   525
      TabIndex        =   0
      Top             =   855
      Width           =   5535
      _cx             =   9763
      _cy             =   6562
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483633
      AutoScroll      =   -1  'True
      VirtualWidth    =   1000
      VirtualHeight   =   1000
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   -1  'True
      MouseScroll     =   -1  'True
      ProportionalBars=   -1  'True
      Begin Threed.SSPanel SSPanel1 
         Height          =   3360
         Left            =   165
         TabIndex        =   1
         Top             =   225
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5927
         _Version        =   131074
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Begin VsOcxLib.VideoSoftIndexTab vsTabModel 
            Height          =   2985
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4905
            _Version        =   327680
            _ExtentX        =   8652
            _ExtentY        =   5265
            _StockProps     =   102
            Caption         =   "Модель|Прогон|Калибровка"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            ForeColor       =   0
            Style           =   2
            Position        =   2
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            ShowFocusRect   =   0   'False
            BackSheets      =   1
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            FrontTabForeColor=   16711680
            New3D           =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frxModelMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_def_MinWidth = 7000
Private Const m_def_MinHeight = 4000


Private m_mg_GertNet As MGertNet

Public Property Get MeX() As FormX
    ' Property MeX is the FormX equivalent of the keyword Me
    Set MeX = frmxChild
    ' N.B. A more general definition for MeX that does not hardwire the name of the embedded FormX control
    '      is as follows.  The more general definition is less efficient because it calls a helper function.
    '   Public Property Get MeX() as FormX
    '       Set MeX = otxFormX(Me)
    '   End Property
    '
    '   The general definition is convenient if you paste the definition into form modules
    '   that use different names for the embedded FormX control.
End Property

Public Property Get AutoFrame() As CollectionX
    Set AutoFrame = MeX.AutoFrame
End Property

Public Property Set AutoFrame(ByVal NewValue As CollectionX)
    Set MeX.AutoFrame = NewValue
End Property

Public Property Get CanDragDocked() As Variant
    CanDragDocked = MeX.CanDragDocked
End Property

Public Property Let CanDragDocked(ByVal NewValue As Variant)
    MeX.CanDragDocked = NewValue
End Property

Public Property Get ClientEdge() As Variant
    ClientEdge = MeX.ClientEdge
End Property

Public Property Let ClientEdge(ByVal NewValue As Variant)
    MeX.ClientEdge = NewValue
End Property

Public Sub Dock(Optional ByVal Edge, Optional ByVal Panel, Optional ByVal Pos, Optional ByVal Length, Optional ByVal Options, Optional ByVal Attribs)
    MeX.Dock Edge, Panel, Pos, Length, Options, Attribs
End Sub

Public Sub DoMDI(Optional ByVal Left, Optional ByVal Top, Optional ByVal Width, Optional ByVal Height, Optional ByVal Options, Optional ByVal Attribs)
    MeX.DoMDI Left, Top, Width, Height, Options, Attribs
End Sub

Public Property Get Edge() As FormXLib.otxDockEdge
    Edge = MeX.Edge
End Property

Public Property Let Edge(ByVal NewValue As FormXLib.otxDockEdge)
    MeX.Edge = NewValue
End Property

Public Sub Float(Optional ByVal Left, Optional ByVal Top, Optional ByVal Width, Optional ByVal Height, Optional ByVal Options, Optional ByVal Attribs)
    MeX.Float Left, Top, Width, Height, Options, Attribs
End Sub

Public Property Get FrameType() As otxChildFormFrameType
    FrameType = MeX.FrameType
End Property

Public Property Let FrameType(ByVal NewValue As otxChildFormFrameType)
    MeX.FrameType = NewValue
End Property

Public Property Get MDIForm() As Object
    Set MDIForm = MeX.MDIForm
End Property

Public Property Get MDIFormX() As Object
    Set MDIFormX = MeX.MDIFormX
End Property

Public Property Get UseCaptionButton() As Variant
    UseCaptionButton = MeX.UseCaptionButton
End Property

Public Property Let UseCaptionButton(ByVal NewValue As Variant)
    MeX.UseCaptionButton = NewValue
End Property

Public Property Get UseCaptionDblClick() As Variant
    UseCaptionDblClick = MeX.UseCaptionDblClick
End Property

Public Property Let UseCaptionDblClick(ByVal NewValue As Variant)
    MeX.UseCaptionDblClick = NewValue
End Property

Public Property Get UseCaptionRtDblClick() As Variant
    UseCaptionRtDblClick = MeX.UseCaptionRtDblClick
End Property

Public Property Let UseCaptionRtDblClick(ByVal NewValue As Variant)
    MeX.UseCaptionRtDblClick = NewValue
End Property

Private Sub Form_Activate()
  With frmMain.OtxCommandManager1.ToolBars("Файл").CommandUIs("SF").Command
      .Checked = otxUnpressed
  End With
  
  With frmMain.OtxCommandManager1.ToolBars("Файл").CommandUIs("Model").Command
      .Checked = otxPressed
  End With
End Sub

Private Sub Form_Load()
    ' *** OT/X DockingForms Usage Requirement ***
    '   The standard form property MDIChild must be false.
    '   The property FrameType on the embedded FormX control must be used to make
    '   the form an MDIChild.  See the Usage Requirements section in the help file
    '   for more details.  Setting the MDIChild property to true makes the VB form
    '   incompatible with the FormX control.  Results in this case are undefined.
    Debug.Assert Me.MDIChild = False
    
    'vsTabModel.Move 0, 0
        
    
    VSViewPort1.VirtualWidth = m_def_MinWidth
    VSViewPort1.VirtualHeight = m_def_MinHeight
    SSPanel1.Move 0, 0, m_def_MinWidth, m_def_MinHeight
          
    vsTabModel.AutoSwitch = False
    
    Dim fExt1 As VBControlExtender
    Set fExt1 = Controls.Add("Hazard.CfrmMEditor", "frmMEditor", vsTabModel)
    Set frmMEditor = fExt1
    
    Dim fExt2 As VBControlExtender
    Set fExt2 = Controls.Add("Hazard.CfrmMon", "frmMon", vsTabModel)
    Set frmMon = fExt2
    
    Dim fExt3 As VBControlExtender
    Set fExt3 = Controls.Add("Hazard.CfrmCalibrate", "frmCalibrate", vsTabModel)
    Set frmCalibrate = fExt3
    
'    Dim fExt4 As VBControlExtender
'    Set fExt4 = Controls.Add("Hazard.CfrmProcs", "frmProcs", vsTabModel)
'    Set frmProcs = fExt4
    
'    Load frmProcs
'    MkFChild frmProcs.hwnd
'    SetParent frmProcs.hwnd, vsTabModel.hwnd
'
'    Load frmCalibrate
'    MkFChild frmCalibrate.hwnd
'    SetParent frmCalibrate.hwnd, vsTabModel.hwnd
'
'    Load frmMon
'    MkFChild frmMon.hwnd
'    SetParent frmMon.hwnd, vsTabModel.hwnd
'
'    Load frmMEditor
'    MkFChild frmMEditor.hwnd
'    SetParent frmMEditor.hwnd, vsTabModel.hwnd
    
    fExt1.Move 0, 0
    fExt2.Move 100, 100
    fExt3.Move 200, 200
    'fExt4.Move 300, 300
    
    DoEvents
    UpdateWindow hWnd
    
    vsTabModel.AutoSwitch = True
    
    vsTabModel.Visible = True

    'fExt4.Visible = True
    fExt3.Visible = True
    fExt2.Visible = True
    fExt1.Visible = True

    
    Dim iCancel As Integer
    vsTabModel_Switch -1, 0, iCancel
    
    DoEvents
    UpdateWindow hWnd
        
End Sub

Private Sub MkFChild(ByVal hw As Long)
  SetWindowLong hw, GWL_STYLE, (GetWindowLong(hw, GWL_STYLE) Or WS_CHILD) And (Not WS_POPUP)
End Sub

Private Sub Form_Resize()
  VSViewPort1.Move 0, 0, ScaleWidth, ScaleHeight
  
  Dim lW As Long, lH As Long
  Dim sMW As Single, sMH As Single, fCurr As Object
  Set fCurr = FormByIdx(vsTabModel.CurrTab)
  sMW = fCurr.MinW: sMH = fCurr.MinH
  lW = IIf(VSViewPort1.Width > sMW, VSViewPort1.Width, sMW)
  lH = IIf(VSViewPort1.Height > sMH, VSViewPort1.Height, sMH)
  'vsTabModel.Width = lW: vsTabModel.Height = lH
  'SSPanel1.Move 0, 0, lW, lH
  SSPanel1.Width = lW: SSPanel1.Height = lH
  vsTabModel.Move 0, 0, SSPanel1.Width, SSPanel1.Height
  
End Sub

Private Sub Form_Terminate()
  Set m_mg_GertNet = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not haApp Is Nothing Then
    If haApp.IsCalcRun Then _
      Cancel = 1: MsgBox "При активном процессе расчёта закрыть нельзя", vbInformation Or vbOKOnly, "Предупреждение"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

  With frmMain.OtxCommandManager1.ToolBars("Файл").CommandUIs("Model").Command
      .Checked = otxUnpressed
  End With
    
  vsTabModel.AutoSwitch = False
    
'  SetParent frmMEditor.hwnd, 0
'  Unload frmMEditor
'
'  SetParent frmMon.hwnd, 0
'  Unload frmMon
'
'  SetParent frmCalibrate.hwnd, 0
'  Unload frmCalibrate
'
'  SetParent frmProcs.hwnd, 0
'  Unload frmProcs
          
'  Set frmProcs = Nothing
'  Set frmCalibrate = Nothing
'  Set frmMon = Nothing
'  Set frmMEditor = Nothing

  frmMEditor.HandsOffModel
  frmMon.HandsOffModel
  frmCalibrate.HandsOffModel
  'frmProcs.HandsOffModel
  
  HandsOffModel_Internal
  
  Hide
  DoEvents
  
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Private Sub HandsOffModel_Internal()
  Set m_mg_GertNet = Nothing
End Sub


Public Sub HandsOffModel()
  'frxRep.vsModel.EndDoc
  
  frmMEditor.HandsOffModel
  frmMon.HandsOffModel
  frmCalibrate.HandsOffModel
  'frmProcs.HandsOffModel
  
  HandsOffModel_Internal
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  On Error GoTo ERR_H
  
  Set m_mg_GertNet = mgn
  frmMEditor.BindModel mgn
  frmMon.BindModel mgn
  frmCalibrate.BindModel mgn
  'frmProcs.BindModel mgn
  
  'frxRep.vsModel.StartDoc
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Friend Function FormByIdx(ByVal lIdx As Long) As Object
  Select Case lIdx
    Case 0
      Set FormByIdx = frmMEditor
    Case 1
      Set FormByIdx = frmMon
    Case 2
      Set FormByIdx = frmCalibrate
    'Case 3
     ' Set FormByIdx = frmProcs
    Case Else
  End Select
End Function

Private Sub vsTabModel_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
  If OldTab <> -1 Then
    If Not FormByIdx(OldTab).OnSwitchTo(False) Then
      Cancel = 1
      Exit Sub
    End If
  End If
  
  With FormByIdx(NewTab)
    Caption = .Caption
    .OnSwitchTo True
    VSViewPort1.VirtualWidth = .MinW
    VSViewPort1.VirtualHeight = .MinH
  End With
  Form_Resize
End Sub

Public Sub GetExtraWH(ByRef sW As Single, ByRef sH As Single)
  Dim bFlVS As Boolean, bFlHS As Boolean
  bFlVS = VSViewPort1.VirtualHeight < VSViewPort1.Height
  bFlHS = VSViewPort1.VirtualWidth < VSViewPort1.Width
  
  If bFlVS Or Not bFlHS Then
    sW = 0
  Else
    sW = ScaleX(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
  End If
  
  If bFlHS Or Not bFlVS Then
    sH = 0
  Else
    sH = ScaleY(GetSystemMetrics(SM_CYHSCROLL), vbPixels, vbTwips)
  End If
End Sub
