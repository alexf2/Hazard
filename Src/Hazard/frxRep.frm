VERSION 5.00
Object = "{7E00A3A2-8F5C-11D2-BAA4-04F205C10000}#1.0#0"; "VsVIEW6.ocx"
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "formx.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Begin VB.Form frxRep 
   Caption         =   "Отчёты"
   ClientHeight    =   3945
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6960
   Icon            =   "frxRep.frx":0000
   LinkTopic       =   "frmDocked"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand sscZoom 
      Height          =   270
      Left            =   6555
      TabIndex        =   14
      ToolTipText     =   "Масштаб"
      Top             =   3645
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   476
      _Version        =   131074
      PictureMaskColor=   255
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Windowless      =   0   'False
      Picture         =   "frxRep.frx":000C
      AutoSize        =   2
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand sscPrint 
      Height          =   270
      Left            =   6195
      TabIndex        =   13
      ToolTipText     =   "Печатать отчёт"
      Top             =   3660
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   476
      _Version        =   131074
      PictureMaskColor=   255
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Windowless      =   0   'False
      Picture         =   "frxRep.frx":010E
      AutoSize        =   2
      ButtonStyle     =   4
      BevelWidth      =   0
   End
   Begin Threed.SSPanel lblZoom 
      Height          =   210
      Index           =   0
      Left            =   4035
      TabIndex        =   10
      Top             =   3695
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   370
      _Version        =   131074
      ForeColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Масштаб:"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   1
   End
   Begin Threed.SSPanel lblPage 
      Height          =   165
      Index           =   1
      Left            =   3270
      TabIndex        =   11
      Top             =   3718
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   291
      _Version        =   131074
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      ClipControls    =   0   'False
      BevelOuter      =   0
   End
   Begin Threed.SSPanel lblZoom 
      Height          =   165
      Index           =   1
      Left            =   5190
      TabIndex        =   12
      Top             =   3718
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   291
      _Version        =   131074
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "100"
      ClipControls    =   0   'False
      BevelOuter      =   0
   End
   Begin Threed.SSPanel lblPage 
      Height          =   210
      Index           =   0
      Left            =   2085
      TabIndex        =   9
      Top             =   3695
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   370
      _Version        =   131074
      ForeColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Страница:"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   1
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   0
      Left            =   2775
      TabIndex        =   8
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   1
      Left            =   3045
      TabIndex        =   7
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   2
      Left            =   3675
      TabIndex        =   6
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdPage 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   3
      Left            =   3930
      TabIndex        =   5
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   0
      Left            =   4710
      TabIndex        =   4
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   1
      Left            =   4965
      TabIndex        =   3
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   2
      Left            =   5610
      TabIndex        =   2
      Top             =   3675
      Width           =   250
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Index           =   3
      Left            =   5865
      TabIndex        =   1
      Top             =   3675
      Width           =   250
   End
   Begin VsOcxLib.VideoSoftIndexTab tabPrint 
      Height          =   3510
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   6510
      _Version        =   327680
      _ExtentX        =   11483
      _ExtentY        =   6191
      _StockProps     =   102
      Caption         =   "&Модель|&Комплекс мер"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Style           =   0
      Position        =   1
      AutoScroll      =   -1  'True
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      CaptionStyle    =   4
      New3D           =   0   'False
      Begin VSVIEW6Ctl.VSPrinter vpSP 
         Height          =   3180
         Left            =   75015
         TabIndex        =   16
         ToolTipText     =   "Отчёты по комплексам мер"
         Top             =   15
         Width           =   6480
         _cx             =   11430
         _cy             =   5609
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _ConvInfo       =   1
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   400
         MarginTop       =   1440
         MarginRight     =   400
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   1
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   60
         SmallChangeVert =   60
         Track           =   -1  'True
         ProportionalBars=   -1  'True
         Zoom            =   16.0427807486631
         ZoomMode        =   0
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   5
         MouseZoom       =   2
         MouseScroll     =   -1  'True
         MousePage       =   -1  'True
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         HTMLStyle       =   1
      End
      Begin VSVIEW6Ctl.VSPrinter vsModel 
         Height          =   3180
         Left            =   15
         TabIndex        =   15
         ToolTipText     =   "Отчёты по модели"
         Top             =   15
         Width           =   6480
         _cx             =   11430
         _cy             =   5609
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _ConvInfo       =   1
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   400
         MarginTop       =   1440
         MarginRight     =   400
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   1
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   60
         SmallChangeVert =   60
         Track           =   -1  'True
         ProportionalBars=   -1  'True
         Zoom            =   16.0427807486631
         ZoomMode        =   0
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   5
         MouseZoom       =   2
         MouseScroll     =   -1  'True
         MousePage       =   -1  'True
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         HTMLStyle       =   1
      End
   End
   Begin FormXLib.FormX frmxChild 
      Left            =   1845
      Top             =   0
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
      FrameType       =   4
      ClientEdge      =   0
      CoolLookBorder  =   1
      AutoFrameMDIChild=   0
      AutoFrameFloating=   1
      AutoFrameDockTop=   0
      AutoFrameDockLeft=   1
      AutoFrameDockBottom=   1
      AutoFrameDockRight=   1
   End
End
Attribute VB_Name = "frxRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_def_XGap = 50#



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

Friend Function CurrPrinter() As VSPrinter
  Select Case tabPrint.CurrTab
    Case 0
      Set CurrPrinter = vsModel
    Case 1
      Set CurrPrinter = vpSP
  End Select
End Function

Private Sub cmdPage_Click(Index As Integer)
  Dim vsp As VSPrinter: Set vsp = CurrPrinter
  Select Case Index
    Case 0
        vsp.PreviewPage = 1
    Case 1
        vsp.PreviewPage = vsp.PreviewPage - 1
    Case 2
        vsp.PreviewPage = vsp.PreviewPage + 1
    Case 3
        vsp.PreviewPage = vsp.PageCount
  End Select
    UpdateControlBar
End Sub

Public Sub UpdateControlBar()
  Dim vsp As VSPrinter: Set vsp = CurrPrinter
  
  ' update labels
  lblPage(1) = Format(vsp.PreviewPage)
  lblZoom(1) = Format(vsp.Zoom, "0")
      
  ' enable/disable buttons
  cmdPage(0).Enabled = (vsp.PreviewPage > 1)
  cmdPage(1).Enabled = (vsp.PreviewPage > 1)
  cmdPage(2).Enabled = (vsp.PreviewPage < vsp.PageCount)
  cmdPage(3).Enabled = (vsp.PreviewPage < vsp.PageCount)
  cmdZoom(0).Enabled = (vsp.Zoom > vsp.ZoomMin)
  cmdZoom(1).Enabled = (vsp.Zoom > vsp.ZoomMin)
  cmdZoom(2).Enabled = (vsp.Zoom < vsp.ZoomMax)
  cmdZoom(3).Enabled = (vsp.Zoom < vsp.ZoomMax)
  
  ' variable zoom rate
  If vsp.Zoom >= 50 Then
      vsp.ZoomStep = 10
  Else
      vsp.ZoomStep = 5
  End If
End Sub



Private Sub cmdZoom_Click(Index As Integer)
  Dim vsp As VSPrinter: Set vsp = CurrPrinter
  Dim z%
  Select Case Index
    Case 0
        z = vsp.ZoomMin
    Case 1
        z = vsp.Zoom - vsp.ZoomStep
    Case 2
        z = vsp.Zoom + vsp.ZoomStep
    Case 3
        z = vsp.ZoomMax
  End Select
  If z < vsp.ZoomMin Then z = vsp.ZoomMin
  If z > vsp.ZoomMax Then z = vsp.ZoomMax
  vsp.Zoom = z
  UpdateControlBar
End Sub

Private Sub Form_Load()
    ' *** OT/X DockingForms Usage Requirement ***
    '   The standard form property MDIChild must be false.
    '   The property FrameType on the embedded FormX control must be used to make
    '   the form an MDIChild.  See the Usage Requirements section in the help file
    '   for more details.  Setting the MDIChild property to true makes the VB form
    '   incompatible with the FormX control.  Results in this case are undefined.
    Debug.Assert Me.MDIChild = False
    
    Load frmPrtShadow
    UpdateControlBar
End Sub

Private Sub Form_Resize()
  
  tabPrint.Move 0, 0, ScaleWidth, ScaleHeight
    
  Dim sX As Single, sY As Single, sTY As Single
  sX = 2400
  sTY = tabPrint.Height - tabPrint.ClientHeight
  'sY = tabPrint.Top + tabPrint.Height - 2# * sTY / 3#
  sY = ScaleHeight - 250
    
  sTY = (cmdPage(0).Height - lblPage(0).Height) / 2#
  
  lblPage(0).Move sX, sY: sX = sX + lblPage(0).Width + m_def_XGap
  cmdPage(0).Move sX, sY: sX = sX + cmdPage(0).Width + m_def_XGap
  cmdPage(1).Move sX, sY: sX = sX + cmdPage(1).Width
  lblPage(1).Move sX, sY + sTY: sX = sX + lblPage(1).Width
  cmdPage(2).Move sX, sY: sX = sX + cmdPage(2).Width + m_def_XGap
  cmdPage(3).Move sX, sY: sX = sX + cmdPage(3).Width + m_def_XGap
  
  lblZoom(0).Move sX, sY: sX = sX + lblZoom(0).Width + m_def_XGap
  cmdZoom(0).Move sX, sY: sX = sX + cmdZoom(0).Width + m_def_XGap
  cmdZoom(1).Move sX, sY: sX = sX + cmdZoom(1).Width
  lblZoom(1).Move sX, sY + sTY: sX = sX + lblZoom(1).Width
  cmdZoom(2).Move sX, sY: sX = sX + cmdZoom(2).Width + m_def_XGap
  cmdZoom(3).Move sX, sY: sX = sX + cmdZoom(3).Width + m_def_XGap
  
  'cmdPrint(0).Move sX, sY: sX = sX + cmdPrint(0).Width + m_def_XGap
  sscPrint.Move sX, sY - 20: sX = sX + sscPrint.Width + m_def_XGap
  sscZoom.Move sX, sY - 20: sX = sX + sscZoom.Width + m_def_XGap
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload frmPrtShadow
End Sub

Private Sub sscPrint_Click()
  PopupMenu frmPrtShadow.mnuPrint
End Sub

Private Sub sscZoom_Click()
  PopupMenu frmPrtShadow.mnuZoom
End Sub


Private Sub tabPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then _
    PopupMenu frmPrtShadow.mnuRpt
End Sub

Private Sub tabPrint_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
  UpdateControlBar
End Sub

Private Sub vpSP_EndDoc()
  'MakePageLabels vsSP
End Sub

Private Sub vpSP_EndPage()
  MakePageLabels vpSP, vpSP.CurrentPage, vpSP.CurrentPage, "Комплекс мер"
End Sub

Private Sub vpSP_MousePage(NewPage As Integer)
  If NewPage <> 0 Then UpdateControlBar
End Sub

Private Sub vpSP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton And Shift = 0 Then
    Dim x2 As Single, y2 As Single
    x2 = x: y2 = y
    With vpSP
      .ClientToPage x2, y2
      If x2 < 0 Or x2 > .PageWidth Or _
         y2 < 0 Or y2 > .PageHeight Then _
        PopupMenu frmPrtShadow.mnuRpt
    End With
  End If
End Sub

Private Sub vpSP_MouseZoom(NewZoom As Integer)
  If NewZoom <> 0 Then UpdateControlBar
End Sub

Private Sub vsModel_EndDoc()
  'MakePageLabels vsModel
End Sub

Friend Sub MakePageLabels(ByVal vp As VSPrinter, ByVal lP1 As Long, ByVal lP2 As Long, ByVal sHeader As String)

  Dim sFName As String, fSize As Long, fBold As Boolean
  Dim fItalic As Boolean, fTextAlign As Long
  Dim fCurrentX As Long, fCurrentY As Long, fColumns As Long
  Dim fLS As Long

  Dim i As Long
  With vp
       
    sFName = .FontName
    fSize = .FontSize
    fBold = .FontBold
    fItalic = .FontItalic
    fTextAlign = .TextAlign
    fCurrentX = .CurrentX
    fCurrentY = .CurrentY
    fColumns = .Columns
    fLS = .LineSpacing
    
    ' set up to print
    .Font.Charset = 204
    .Columns = 1
    .LineSpacing = 100
    
    ' loop through all pages
    For i = lP1 To lP2
        
        .FontName = "Arial"
        .FontSize = 12
        .FontBold = True
        .FontItalic = True
        .TextAlign = taRightTop
    
        ' start the overlay on the current page
        '.StartOverlay i
        
        ' draw the footer
        .CurrentX = .MarginLeft
        .CurrentY = .PageHeight - .MarginBottom + 100
        vp = "Страница " & Format(i) ' & " из " & Format(.PageCount)
        
        .CurrentX = .MarginLeft
        .CurrentY = .PageHeight - .MarginBottom + 100
        .FontName = "Times New Roman"
        .FontSize = 9
        .TextAlign = taLeftTop
        vp = "Hazard 3.0 (c)AlexCorp.&GA 1998-2000"
        
        '.FontName = "Arial"
        .FontSize = 12
        .TextAlign = taLeftBottom
        '.TextAlign = taLeftTop
        .CurrentX = .MarginLeft
        .CurrentY = .MarginTop - 370
        '.CurrentY = .MarginHeader
        vp = sHeader
        
        .FontSize = 10
        .FontName = "Courier New"
        .TextAlign = taRightBottom
        .CurrentX = .PageWidth - .MarginRight
        .CurrentY = .MarginTop - 570
        .FontItalic = False
        vp = Format(Now, "d/mmm/yyyy h:mm")
        
        
        ' finish the overlay
        '.EndOverlay
    Next i
    .LineSpacing = fLS
    .FontName = sFName
    .FontSize = fSize
    .FontBold = fBold
    .FontItalic = fItalic
    .TextAlign = fTextAlign
    .CurrentX = fCurrentX
    .CurrentY = fCurrentY
    .Columns = fColumns
  End With
End Sub


Private Sub vsModel_EndPage()
  MakePageLabels vsModel, vsModel.CurrentPage, vsModel.CurrentPage, "Модель"
End Sub

Private Sub vsModel_MousePage(NewPage As Integer)
  If NewPage <> 0 Then
    UpdateControlBar
  End If
End Sub

Private Sub vsModel_MouseScroll()
  
'  Dim sX As Single, sY As Single
'
'  sY = vsModel.PageHeight * (vsModel.Zoom / 100) - vsModel.Height - 300
'  If sY > 0 And vsModel.ScrollTop > sY + 10 Then
'    If vsModel.PreviewPage < vsModel.PageCount Then _
'      vsModel.PreviewPage = vsModel.PreviewPage + 1: vsModel.ScrollTop = 0
'  ElseIf sY > 0 And vsModel.ScrollTop <= 0 Then
'    If vsModel.PreviewPage > 1 Then _
'      vsModel.PreviewPage = vsModel.PreviewPage - 1: vsModel.ScrollTop = sY
'  End If
End Sub

Private Sub vsModel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton And Shift = 0 Then
    Dim x2 As Single, y2 As Single
    x2 = x: y2 = y
    With vsModel
      .ClientToPage x2, y2
      If x2 < 0 Or x2 > .PageWidth Or _
         y2 < 0 Or y2 > .PageHeight Then _
        PopupMenu frmPrtShadow.mnuRpt
    End With
  End If
End Sub

Private Sub vsModel_MouseZoom(NewZoom As Integer)
  If NewZoom <> 0 Then UpdateControlBar
End Sub

