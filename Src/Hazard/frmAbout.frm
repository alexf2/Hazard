VERSION 5.00
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#105.0#0"; "AlexOCX.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   873
      _Version        =   131074
      Font3D          =   5
      MarqueeStyle    =   1
      ForeColor       =   128
      BackColor       =   12632256
      MarqueeDelay    =   40
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "The Hazard 2000, known as Hazard 3.0"
      ClipControls    =   0   'False
      BevelOuter      =   0
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   2280
      Left            =   195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
      _Version        =   327680
      _ExtentX        =   2990
      _ExtentY        =   4022
      _StockProps     =   70
      ConvInfo        =   1418783674
      ChildSpacing    =   0
      BevelInnerWidth =   0
   End
   Begin Threed.SSCommand ssOk 
      Default         =   -1  'True
      Height          =   390
      Left            =   2190
      TabIndex        =   6
      Top             =   3330
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   688
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Windowless      =   0   'False
      Picture         =   "frmAbout.frx":000C
      Caption         =   "Да"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   2670
      Left            =   2055
      TabIndex        =   5
      Top             =   540
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   4710
      _Version        =   131074
      MarqueeDirection=   3
      MarqueeStyle    =   1
      ForeColor       =   12619904
      BackColor       =   12632256
      MarqueeDelay    =   200
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   $"frmAbout.frx":0176
      ClipControls    =   0   'False
      BevelOuter      =   0
   End
   Begin AlexOCX.AngularLabel AngularLabel3 
      Height          =   270
      Left            =   3000
      TabIndex        =   4
      Top             =   2415
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   476
      ForeColor       =   128
      FontCtx         =   "frmAbout.frx":0254
      Caption         =   "1998-2000"
   End
   Begin AlexOCX.AngularLabel AngularLabel2 
      Height          =   1650
      Left            =   3480
      TabIndex        =   3
      Top             =   990
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   2910
      ForeColor       =   128
      FontCtx         =   "frmAbout.frx":02AC
      Caption         =   "(c)AlexCorp. & GA"
   End
   Begin AlexOCX.AngularLabel AngularLabel1 
      Height          =   1785
      Left            =   2190
      TabIndex        =   2
      Top             =   930
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   3149
      ForeColor       =   128
      FontCtx         =   "frmAbout.frx":0304
      Caption         =   "Version 3.0 beta"
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    Hide
  ElseIf KeyAscii = vbKeyReturn Then
    ssOk.SetFocus
    ssOk_Click
  End If
End Sub


Private Sub Form_Load()
  Set VideoSoftElastic1.Picture = LoadResPicture(102, vbResBitmap)
End Sub

Private Sub ssOk_Click()
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

