VERSION 5.00
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#104.0#0"; "AlexOCX.ocx"
Object = "{CC52DF59-28C5-11D4-8F1B-00504E02C39D}#67.0#0"; "GNControls.ocx"
Begin VB.Form Form_TstOCX 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin GNControls.CtlRepeater CtlRepeater1 
      Height          =   2085
      Left            =   0
      TabIndex        =   2
      Top             =   2415
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   3678
      TypeMask        =   "M"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   4050
      TabIndex        =   1
      Top             =   510
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1905
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin AlexOCX.Bulb Bulb1 
      Height          =   345
      Left            =   480
      Top             =   465
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      ColorLight      =   14737632
      DrawMode        =   1
      AutoResize      =   -1  'True
      TimeOn          =   500
   End
End
Attribute VB_Name = "Form_TstOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cr As CtlRepeater
Attribute cr.VB_VarHelpID = -1

Private Sub Bulb2_Click(ByVal aocxSender As AlexOCX.Bulb)

End Sub

Private Sub Command1_Click()
  Bulb1.PulseStart
  'CtlRepeater1.Width = CtlRepeater1.Width + 10
  
End Sub

Private Sub Command2_Click()
  Bulb1.PulseStop
End Sub

Private Sub cr_XClickAssEnum(ByVal feSender As Factor)
  MsgBox feSender.Name, vbOKOnly, "Enum"
End Sub

Private Sub cr_XClickSetupValue(ByVal feSender As Factor)
  MsgBox feSender.Name, vbOKOnly, "Value"
End Sub

Private Sub cr_XFactorChanged(ByVal fac As Factor)
  MsgBox fac.Name, vbOKOnly, "Factor changed"
End Sub

Private Sub Form_Load()
  Dim mgn As New GERTNETLib.MGertNet
  Dim ips As IPersistStorage
  Set ips = mgn

  Dim iStg As IStorage
  Set iStg = OpenCF("g:\tst_percent.hzd", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ips.Load iStg

  'Set FacEdit1.GNet = mgn
  'Set FacEdit1.HFactor = mgn.Factors("M01")

  Set CtlRepeater1.GNet = mgn
  'Set CtlRepeater2.GNet = mgn
  'Set cr = CtlRepeater1
End Sub

Private Sub Form_Resize()
  'vsIndexTab1.Left = 0
  'vsIndexTab1.Width = ScaleWidth
  'vsIndexTab1.Height = ScaleHeight - vsIndexTab1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set CtlRepeater1.GNet = Nothing
  'Set CtlRepeater2.GNet = Nothing
  Set cr = Nothing
  'Set FacEdit1.GNet = Nothing
  'Set FacEdit1.HFactor = Nothing
End Sub
