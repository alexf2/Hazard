VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RailWay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFactorAssign

Private m_lMRes As Long

Public Property Get IFactorAssign_ModalResult() As Long
  IFactorAssign_ModalResult = m_lMRes
End Property

Public Sub IFactorAssign_About(ByVal frm As Object, ByVal lHwnd As Long)
  frmAboutRailway.Show vbModal, frm
End Sub

Public Sub IFactorAssign_AssignFactor(ByVal frm As Object, ByVal lHwnd As Long, ByVal mg As MGertNet, ByVal f As Factor, ByVal sCfgName As String, ByVal sPath As String)
  Dim vRes As VbMsgBoxResult
  vRes = MsgBox(f.Name & " = " & f.Value, vbQuestion Or vbYesNo, "Железная дорога: " & sCfgName)
  If vRes = vbYes Then f.Value = 9
  m_lMRes = 1
End Sub

Public Property Get IFactorAssign_Description() As String
  IFactorAssign_Description = "Железная дорога"
End Property


Public Sub IFactorAssign_Configure(ByVal frm As Object, ByVal lHwnd As Long, ByVal sConfigName As String, ByVal sFileName As String)
  MsgBox "Конфигурация '" & sConfigName & "'. Файл: " & sFileName, vbOKOnly, "RailWay"
End Sub

