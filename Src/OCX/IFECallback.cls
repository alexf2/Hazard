VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFECallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ClickAssEnum(ByVal feSender As Factor)
End Sub

Public Sub ClickPr(ByVal feSender As Factor)
End Sub

Public Sub ClickAskVal(ByVal feSender As Factor)
End Sub

Public Sub ClickSetupValue(ByVal feSender As Factor)
End Sub

Public Sub FactorChanged(ByVal fac As Factor)
End Sub

Public Sub FactorFocus(ByVal facEd As IFacED)
End Sub

