VERSION 5.00
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "formx.ocx"
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#104.0#0"; "AlexOCX.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmCancel 
   BorderStyle     =   0  'None
   Caption         =   "Процессы"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frxCancel.frx":0000
   LinkTopic       =   "frmDockedOnly"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1470
   Begin FormXLib.FormX frmxChild 
      Left            =   1050
      Top             =   615
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
      FrameType       =   4
      ClientEdge      =   0
      UseCaptionDblClick=   0
      CanDragDocked   =   0
      CoolLookBorder  =   1
      AutoFrameMDIChild=   0
      AutoFrameFloating=   1
      AutoFrameDockLeft=   0
      AutoFrameDockBottom=   0
      AutoFrameDockRight=   0
   End
   Begin Threed.SSPanel ssShortName 
      Height          =   210
      Left            =   135
      TabIndex        =   0
      Top             =   375
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   370
      _Version        =   131074
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   0
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
      Caption         =   "П"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   210
      Left            =   1170
      TabIndex        =   1
      Top             =   375
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   370
      _Version        =   131074
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   0
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
      Caption         =   "О"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   210
      Left            =   660
      TabIndex        =   2
      Top             =   375
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   370
      _Version        =   131074
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   0
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
      Caption         =   "Р"
      ClipControls    =   0   'False
      BevelOuter      =   0
      AutoSize        =   1
   End
   Begin AlexOCX.Bulb blbRang 
      Height          =   345
      Left            =   562
      ToolTipText     =   "Прервать ранжировку мер"
      Top             =   30
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      DrawMode        =   1
      AutoResize      =   -1  'True
      CantWidth       =   2
      TimeOn          =   500
   End
   Begin AlexOCX.Bulb blbOpt 
      Height          =   345
      Left            =   1080
      ToolTipText     =   "Прервать оптимизацию"
      Top             =   30
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      DrawMode        =   1
      AutoResize      =   -1  'True
      CantWidth       =   2
      TimeOn          =   500
   End
   Begin AlexOCX.Bulb blbRunOnce 
      Height          =   345
      Left            =   45
      ToolTipText     =   "Прервать прогон модели"
      Top             =   30
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      DrawMode        =   1
      AutoResize      =   -1  'True
      CantWidth       =   2
      TimeOn          =   500
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_LockResize As Boolean

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

Private Sub blbOpt_Click(aocxSender As AlexOCX.Bulb)
  haApp.CancelOpt
End Sub

Private Sub blbRang_Click(aocxSender As AlexOCX.Bulb)
  haApp.CancelRang
End Sub

Private Sub blbRunOnce_Click(aocxSender As AlexOCX.Bulb)
  haApp.CancelCalc
End Sub

Private Sub Form_Load()
    ' *** OT/X DockingForms Usage Requirement ***
    '   The standard form property MDIChild must be false.
    '   The property FrameType on the embedded FormX control must be used to make
    '   the form an MDIChild.  See the Usage Requirements section in the help file
    '   for more details.  Setting the MDIChild property to true makes the VB form
    '   incompatible with the FormX control.  Results in this case are undefined.
    Debug.Assert Me.MDIChild = False
    
'    Dim lMaxX As Long, lMaxY As Long
'    Dim vbe As VBControlExtender
'    lMaxX = 0: lMaxY = 0
'    For Each vbe In Controls
'      If Not (TypeOf vbe Is FormX) Then
'        If lMaxX < vbe.Left + vbe.Width Then lMaxX = vbe.Left + vbe.Width
'        If lMaxY < vbe.Top + vbe.Height Then lMaxY = vbe.Top + vbe.Height
'      End If
'    Next vbe
'    If Width <> lMaxX Then Width = lMaxX
'    If Height <> lMaxY Then Height = lMaxY
  m_b_LockResize = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error GoTo ERR_H
  m_b_LockResize = True
  Dim lX As Long, lY As Long
  GetSz lX, lY
  
  If Height > lY Then Height = lY
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub GetSz(lMaxX As Long, lMaxY As Long)
    Dim vbe As VBControlExtender
    lMaxX = 0: lMaxY = 0
    For Each vbe In Controls
      If Not (TypeOf vbe Is FormX) Then
        If lMaxX < vbe.Left + vbe.Width Then lMaxX = vbe.Left + vbe.Width
        If lMaxY < vbe.Top + vbe.Height Then lMaxY = vbe.Top + vbe.Height
      End If
    Next vbe
End Sub
