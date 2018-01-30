VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyCallBack2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICallBack

Private Sub ICallBack_Cancel(ByVal dt As Date, ByVal gn As IMGertNet)

End Sub

Private Sub ICallBack_Cancel3(ByVal dt As Date, ByVal gn As GERTNETLib.IMGertNet)

End Sub

Private Sub ICallBack_EndOfWork(ByVal dt As Date, ByVal gn As GERTNETLib.IMGertNet)
  Dim l1 As Long, l2 As Long, m As Double, d As Double
  gn.GetInfSampleK 73, l1, l2, m, d
  MsgBox "m = " & m
End Sub

Private Sub ICallBack_EndOfWork2(ByVal dt As Date, ByVal gn As GERTNETLib.IMGertNet)

End Sub

Private Sub ICallBack_EndOfWork3(ByVal dt As Date, ByVal gn As GERTNETLib.IMGertNet)

End Sub

Private Sub ICallBack_ErrorCalc(ByVal sMsg As String, ByVal gn As GERTNETLib.IMGertNet)

End Sub

Private Sub ICallBack_StepOfWork(ByVal iPercent As Integer)
  Dim i As Long
  i = 1
End Sub

Private Sub ICallBack_StepOfWork2(ByVal iPercent As Integer)

End Sub

Private Sub ICallBack_StepOfWork3(ByVal iPercent As Integer)

End Sub

