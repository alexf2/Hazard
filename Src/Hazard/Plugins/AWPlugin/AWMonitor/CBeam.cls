VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private m_obj As Object
Private m_l_MP As Long

Public Sub Beam(ByVal obj As Object)
  Set m_obj = obj
  SetCapture obj.hWnd
  m_l_MP = obj.MousePointer
  obj.MousePointer = vbHourglass
End Sub

Public Sub Off()
  If Not (m_obj Is Nothing) Then
    ReleaseCapture
    m_obj.MousePointer = m_l_MP
    Set m_obj = Nothing
  End If
End Sub

Private Sub Class_Terminate()
  Off
End Sub

