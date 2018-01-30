VERSION 5.00
Begin VB.Form AWMTst 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "AWMTst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim swm As New AWM
  Dim ifa As IFactorAssign
  Set ifa = swm
  On Error GoTo ERR_H
  
  ifa.Configure Me, hWnd, "Конф1", "f:\hzd backup\inst\examples\def_ammonia.cfg"
  Exit Sub

  Dim mg As New MGertNet
  Dim ips As IPersistStorage
  Dim iStg As IStorage
  Set iStg = OpenCF("f:\hzd backup\inst\examples\default.hzd", STGM_DIRECT1 Or STGM_READ1 Or STGM_SHARE_EXCLUSIVE1)
  Set ips = mg
  ips.Load iStg: Set iStg = Nothing
  ifa.AssignFactor Me, hWnd, mg, mg.Factors("T03"), "Конф 1", "g:\cg.stg"
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "TT"
  Err.Clear
End Sub
