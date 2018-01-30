VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub EndOfWork(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub EndOfWork2(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub EndOfWork3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub

Public Sub StepOfWork(ByVal iPercent As Integer)
End Sub
Public Sub StepOfWork2(ByVal iPercent As Integer)
End Sub
Public Sub StepOfWork3(ByVal iPercent As Integer)
End Sub

Public Sub Cancel(ByVal dt As Date, ByVal gn As MGertNet)
End Sub
Public Sub Cancel3(ByVal dt As Date, ByVal gn As MGertNet)
End Sub

Public Sub ErrorCalc(ByVal sMsg As String, ByVal gn As MGertNet)
End Sub

