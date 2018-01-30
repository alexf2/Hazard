VERSION 5.00
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   1290
      Top             =   870
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
      WindowText1     =   "Form1"
      WindowKey1      =   "Form1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public haApp As HazardApp
'Public MyCallb As ICallBack

Private Sub Form_Load()
  Set haApp = CreateObject("Hazard.HazardApp")
  'Set MyCallb = New MyCallBack2
  haApp.OpenDoc "f:\\Hzd BackUp\\Inst\\Examples\\default.hzd", True
  haApp.GertNetMainDsp.RunMode = MT_Analytical
  
  
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnEndOfWork", Value:=NM_OnEndOfWork, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnStepOfWork", Value:=NM_OnStepOfWork, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnCancel", Value:=NM_OnCancel, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnEndOfWork2", Value:=NM_OnEndOfWork2, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnStepOfWork2", Value:=NM_OnStepOfWork2, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnEndOfWork3", Value:=NM_OnEndOfWork3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnStepOfWork3", Value:=NM_OnStepOfWork3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnCancel3", Value:=NM_OnCancel3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("Form1").Messages.Add _
    Key:="NM_OnErrorCalc", Value:=NM_OnErrorCalc, ProcessDefaultProc:=False

  GTMsgHook1.Subclass = True
  'lOldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
  
  haApp.GertNetMainDsp.RunMode = MT_AnalyticalIndistinct
  haApp.GertNetMainDsp.NotifyWnd = hwnd
  haApp.StartCalc 100, 200, -1, TP_BELOW_NORMAL, Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  GTMsgHook1.Subclass = False
'  If lOldWndProc <> 0 Then
'    SetWindowLong hwnd, GWL_WNDPROC, lOldWndProc
'    lOldWndProc = 0
'  End If
  
  haApp.Quit
  Set haApp = Nothing
  Set objCli = Nothing
End Sub

Private Sub GTMsgHook1_Message(ByVal hwnd As Long, ByVal MsgID As Long, ByVal wParam As Long, ByVal lParam As Long)
  Select Case MsgID
    Case NM_OnEndOfWork
      Dim l1 As Long, l2 As Long, m As Double, d As Double
      haApp.GN_Run.GetInfSampleK 73, l1, l2, m, d
      MsgBox "m = " & m
    Case NM_OnStepOfWork
    Case NM_OnCancel
    Case NM_OnEndOfWork2
    Case NM_OnStepOfWork2
    Case NM_OnEndOfWork3
    Case NM_OnStepOfWork3
    Case NM_OnCancel3
    Case NM_OnErrorCalc
  End Select
End Sub
