VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlgCfgEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Name As String
Public FileName As String
Public OldName As Variant

Public Dirty As AdvDirty
Public Mounted As Boolean

Private Sub Class_Initialize()
  Dirty = AD_None
  Mounted = False
  OldName = Null
End Sub

Public Sub Modify()
  Select Case Dirty
    Case Ad_Updated
    Case AD_None
      Dirty = Ad_Updated
    Case Ad_Deleted
    Case Ad_New
  End Select
End Sub

