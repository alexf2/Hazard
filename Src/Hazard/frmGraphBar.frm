VERSION 5.00
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "formx.ocx"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "graph32.ocx"
Begin VB.Form frmGraphBar 
   Caption         =   "Функция опасности"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "frmDocked"
   ScaleHeight     =   4050
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5430
      _Version        =   327680
      _ExtentX        =   9578
      _ExtentY        =   7144
      _StockProps     =   70
      ConvInfo        =   1418783674
      Align           =   5
      AutoSizeChildren=   1
      BorderWidth     =   0
      ChildSpacing    =   0
      BevelOuter      =   0
      Begin GraphLib.Graph Graph1 
         Height          =   4050
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5430
         _Version        =   65536
         _ExtentX        =   9578
         _ExtentY        =   7144
         _StockProps     =   96
         AutoInc         =   0
         Background      =   0
         DrawMode        =   3
         Foreground      =   2
         GraphStyle      =   4
         GraphType       =   6
         LegendStyle     =   1
         NumSets         =   2
         Palette         =   1
         PrintStyle      =   1
         RandomData      =   0
         ColorData       =   2
         ColorData[0]    =   3
         ColorData[1]    =   10
         ExtraData       =   0
         ExtraData[]     =   0
         FontFamily      =   4
         FontSize        =   4
         FontSize[0]     =   200
         FontSize[1]     =   150
         FontSize[2]     =   100
         FontSize[3]     =   100
         FontStyle       =   4
         GraphData       =   2
         GraphData[]     =   5
         GraphData[0,0]  =   0
         GraphData[0,1]  =   0
         GraphData[0,2]  =   0
         GraphData[0,3]  =   0
         GraphData[0,4]  =   0
         GraphData[1,0]  =   0
         GraphData[1,1]  =   0
         GraphData[1,2]  =   0
         GraphData[1,3]  =   0
         GraphData[1,4]  =   0
         LabelText       =   0
         LegendText      =   0
         PatternData     =   0
         SymbolData      =   2
         SymbolData[0]   =   3
         SymbolData[1]   =   1
         XPosData        =   0
         XPosData[]      =   0
      End
   End
   Begin FormXLib.FormX frmxChild 
      Left            =   1800
      Top             =   0
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
      FrameType       =   4
      AutoFrameMDIChild=   0
      AutoFrameFloating=   1
      AutoFrameDockTop=   0
      AutoFrameDockLeft=   0
      AutoFrameDockBottom=   1
      AutoFrameDockRight=   0
   End
End
Attribute VB_Name = "frmGraphBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Form_Load()
    ' *** OT/X DockingForms Usage Requirement ***
    '   The standard form property MDIChild must be false.
    '   The property FrameType on the embedded FormX control must be used to make
    '   the form an MDIChild.  See the Usage Requirements section in the help file
    '   for more details.  Setting the MDIChild property to true makes the VB form
    '   incompatible with the FormX control.  Results in this case are undefined.
    Debug.Assert Me.MDIChild = False
End Sub

