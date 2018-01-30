VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmOptimize 
      Caption         =   "Optimize"
      Height          =   375
      Left            =   5445
      TabIndex        =   16
      Top             =   4920
      Width           =   945
   End
   Begin VB.CommandButton cmStaRng 
      Caption         =   "Start"
      Height          =   375
      Left            =   555
      TabIndex        =   15
      Top             =   4920
      Width           =   945
   End
   Begin VB.ComboBox cbnMode 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4875
      List            =   "Form1.frx":0002
      TabIndex        =   11
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtRndInit 
      Height          =   345
      Left            =   3060
      TabIndex        =   9
      Text            =   "1"
      Top             =   4215
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   555
      TabIndex        =   8
      Top             =   4140
      Width           =   945
   End
   Begin VB.CommandButton cmdSta 
      Caption         =   "Start"
      Height          =   375
      Left            =   555
      TabIndex        =   7
      Top             =   3570
      Width           =   945
   End
   Begin VB.TextBox txtN 
      Height          =   345
      Left            =   4710
      TabIndex        =   4
      Text            =   "1000"
      Top             =   3675
      Width           =   855
   End
   Begin VB.TextBox txtK 
      Height          =   345
      Left            =   2895
      TabIndex        =   3
      Text            =   "100"
      Top             =   3675
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   3045
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   6360
   End
   Begin MSComctlLib.ProgressBar pbRng 
      Height          =   360
      Left            =   90
      TabIndex        =   13
      Top             =   5355
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbRng 
      Alignment       =   2  'Center
      Caption         =   "%"
      Height          =   240
      Left            =   2670
      TabIndex        =   14
      Top             =   5070
      Width           =   1530
   End
   Begin VB.Label Label5 
      Caption         =   "Mode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4260
      TabIndex        =   12
      Top             =   4290
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Rnd init:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2220
      TabIndex        =   10
      Top             =   4260
      Width           =   810
   End
   Begin VB.Label Label3 
      Caption         =   "N:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4425
      TabIndex        =   6
      Top             =   3750
      Width           =   210
   End
   Begin VB.Label Label2 
      Caption         =   "K:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2595
      TabIndex        =   5
      Top             =   3750
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "%"
      Height          =   240
      Left            =   2663
      TabIndex        =   2
      Top             =   2745
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mgn As New GERTNETLib.MGertNet
Private WithEvents mgnEv As GERTNETLib.MGertNet
Attribute mgnEv.VB_VarHelpID = -1
Private bFlShow As Boolean

Private collsftGlobal As CollSF

Private cc As Integer

Private Sub EnblControls(ByVal bFl As Boolean)
  
  txtK.Enabled = bFl
  
  txtN.Enabled = bFl
  
  txtRndInit.Enabled = bFl
  cbnMode.Enabled = bFl
  cmdSta.Enabled = bFl
  cmdCancel.Enabled = Not bFl
End Sub

Private Sub cmdCancel_Click()
  mgn.Cancel
End Sub

Private Sub cmdSta_Click()
  On Error GoTo ERR_H
  
  bFlShow = True
  
  Dim bb As Boolean
  bb = mgn.IsRunning
  
  cc = 1
  ProgressBar1.Value = 0
  Label1.Caption = "0%"
  
  With mgn
    Dim mTyp As ModelType
    mTyp = IIf(cbnMode.ListIndex = 0, MT_Imitate, IIf(cbnMode.ListIndex = 1, MT_ImitateIndistinct, MT_Analytical))
    If mTyp <> .RunMode Then .RunMode = mTyp
    'MT_Imitate
    'MT_Analytical
    'MT_ImitateIndistinct
    '.CalibrateModel 0.02, 0.002, 0.0002, 0.00002
    
    .Run CLng(Val(txtK.Text)), CLng(Val(txtN.Text)), False, CLng(Val(txtRndInit.Text))
    EnblControls False
  End With
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub cmOptimize_Click()
  
'  Set collsftGlobal = tCollSFT
  
  If Not collsftGlobal.IsValidDelta Then
    MsgBox "Need calc delta", vbOKOnly, "Error"
  Else
    mgn.OptimizeSelectSF OT_FixMoney_MinQ, collsftGlobal, 290, 0.0005, 2
    EnblControls False
  End If
  
  'tCollSFT.Sorting = CSFS_Price
  'DumpCSF tCollSFT
  
End Sub

Private Sub cmStaRng_Click()
  On Error GoTo ERR_H
    
'  Dim vvtd
'  vvtd = mgn.TimeWork2
  
'  Dim spr As String
'  Dim iVal As Long, sVal As String, dVal As Double, shVal As Integer
'  iVal = 12: sVal = "ss": dVal = 127.1278: shVal = Asc("B")
'  spr = Sprintf("TEST: %u; [%c] %s; %5.2f - [%c]", , iVal, shVal, sVal, dVal, shVal)
'  Exit Sub
    
  bFlShow = False
    
  pbRng.Value = 0
  lbRng.Caption = "0%"
  
  With mgn
    Dim mTyp As ModelType
    mTyp = IIf(cbnMode.ListIndex = 0, MT_Imitate, IIf(cbnMode.ListIndex = 1, MT_ImitateIndistinct, MT_Analytical))
    If mTyp <> .RunMode Then .RunMode = mTyp
   'kkkkkkkkkkkkk
    TstRanging
    EnblControls False
  End With
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub Form_Load()
  
  Set mgnEv = mgn
  
  cbnMode.AddItem "Imitate"
  cbnMode.AddItem "Imitate indistinct"
  cbnMode.AddItem "Analytical"
  cbnMode.ListIndex = 0
  
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
  ProgressBar1.Value = 0

  Dim ips As IPersistStorage
  Set ips = mgn
  
  Dim iStg As IStorage
  Set iStg = OpenCF("g:\default_stg.atg", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ips.Load iStg
  
  mgn.NotifyStep = 10
End Sub

Private Sub mgnEv_OnCancel(ByVal dtTime As Date)
  EnblControls True
  Text1.Text = Text1.Text & "Canceled" & vbCrLf
End Sub

Private Sub mgnEv_OnCancel3(ByVal dtTime As Date)
  EnblControls True
  Text1.Text = Text1.Text & "Canceled" & vbCrLf
End Sub

Private Sub mgnEv_OnEndOfWork(ByVal dtTime As Date)
  'If Not bFlShow Then Exit Sub
      
'  Dim sRep As String
'  EnblControls True
'  sRep = CStr(mgn.InTest1) & "; " & CStr(mgn.InTest2) & "; " & _
'    CStr(mgn.InTest3) & "; " & CStr(mgn.InTest4) & "; " & vbCrLf
'  Text1.Text = Text1.Text & sRep
      
  Dim sRep As String
  EnblControls True
  sRep = "<Work time: " & Format(dtTime, "Long Time") & ">" & vbCrLf & _
    GetPDescr() & _
    "--------------------------" & vbCrLf
  Text1.Text = Text1.Text & sRep
  
End Sub

Private Sub mgnEv_OnEndOfWork2(ByVal dtTime As Date)
  EnblControls True
  DumpCSF collsftGlobal
End Sub

Private Sub mgnEv_OnEndOfWork3(ByVal dtTime As Date)
  EnblControls True
  LoadOptResult mgn
End Sub

Private Sub mgnEv_OnErrorCalc(ByVal bsErrMsg As String)
  MsgBox bsErrMsg, vbOKOnly, "Error"
End Sub

Private Sub mgnEv_OnStepOfWork(ByVal shPercent As Integer)
    
  If cc = 1 Then
    cc = cc + 1
    Dim bb As Boolean
    bb = mgn.IsRunning
    'mgn.Run 2, 10, False, -1
  End If
  Label1.Caption = CStr(shPercent) & "%"
  ProgressBar1.Value = shPercent
End Sub

Private Sub mgnEv_OnStepOfWork2(ByVal shPercent As Integer)
  lbRng.Caption = CStr(shPercent) & "%"
  pbRng.Value = shPercent
End Sub

Private Sub mgnEv_OnStepOfWork3(ByVal shPercent As Integer)
  lbRng.Caption = CStr(shPercent) & "% (optimiz.)"
  pbRng.Value = shPercent
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'ProgressBar1.
End Sub

Private Function GetPDescr() As String
  Dim lMin As Long, lMax As Long, dMx As Double, dDx As Double
  Dim l As Long
  Dim sarr()
  sarr = Array("Гомеостазис", "Нарушение ==", "Адаптация D", _
    "Оп. ситуация", "Адаптация F", "Кр. ситуация", _
    "Адаптация H", "Происшествие")
    
  GetPDescr = "Состояние  Вероятность  Ср.кв. откл.  Min  Max  Индекс опасности" & vbCrLf
  For l = 66 To 73
    mgn.GetInfSampleK l, lMin, lMax, dMx, dDx
    Dim s1 As String, s2 As String, s3 As String, s4 As String
    s1 = Format(sarr(l - 66), "!@@@@@@@@@@@@")
    s2 = Format(dMx, "#0.0#########;;0")
    s3 = Format(dDx, "#0.0#########;;0")
    GetPDescr = GetPDescr & s1 & "  " & s2 & "  " & s3 & "  " & CStr(lMin) & "  " & CStr(lMax) & "  "
    If Not mgn.RunMode = MT_Analytical Then
      mgn.GetInfSampleKIdx l, lMin, lMax, dMx, dDx
      s4 = Format(dMx, "#0.0#########;;0")
      GetPDescr = GetPDescr & s4 & vbCrLf
    Else
      GetPDescr = GetPDescr & vbCrLf
    End If
  Next l
End Function

Private Sub TstRanging()
  Dim tCollSFT As New CollSF
  Set collsftGlobal = tCollSFT
  Dim ips As IPersistStorage
  Set ips = tCollSFT
  
  Dim iStg As IStorage
  Set iStg = OpenCF("g:\default_stg.atg", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  Dim iStg2 As IStorage
  Set iStg2 = OpenStorageInt(iStg, "CollSF_Main", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ips.Load iStg2
  
  tCollSFT.Sorting = CSFS_Price
  DumpCSF tCollSFT
  mgn.CalcDeltaQ tCollSFT, 20, 10000
  
End Sub

Private Sub DumpCSF(ByVal coll As CollSF)
  Dim vvt
  Debug.Print "ID  ", "Cost  ", "DelatQ  ", "Name  "
  For Each vvt In coll
    Dim sf As SafetyPrecaution
    Set sf = vvt
    With sf
      Debug.Print .ID, .Cost, .DeltaQ, .Name
      DumpFChange .FChanges
    End With
  Next vvt
  Debug.Print
End Sub

Private Sub DumpFChange(ByVal coll As ICollFChange)
  Dim vvt
  Debug.Print " ", "ID  ", "NameF  ", "DelatQ  "
  For Each vvt In coll
    Dim sf As FChange
    Set sf = vvt
    With sf
      Debug.Print " ", .ID, .NameFactor, .Delta
    End With
  Next vvt
  Debug.Print
End Sub

Private Sub LoadOptResult(ByVal m As GERTNETLib.MGertNet)
  Dim coll() As CollSF
  Dim vvt
  
  vvt = m.OptimizResultsGetAndClear
  coll = vvt
  Debug.Print "------The optimization:------", vbCrLf
  Dim i As Long
  For i = LBound(coll) To UBound(coll)
    DumpCSF coll(i)
  Next i
  Debug.Print vbCrLf

End Sub
