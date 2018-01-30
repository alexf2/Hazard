VERSION 5.00
Begin VB.UserControl CfrmProcs 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   6045
   ScaleWidth      =   7200
   Begin VB.Label Label1 
      Caption         =   "Procs"
      Height          =   900
      Left            =   960
      TabIndex        =   0
      Top             =   1170
      Width           =   2100
   End
End
Attribute VB_Name = "CfrmProcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mg_GertNet As MGertNet

Public Caption As String

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get MousePointer() As Long
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal lP As Long)
  UserControl.MousePointer = lP
End Property


Public Property Get MinW() As Single
  MinW = 5000
End Property

Public Property Get MinH() As Single
  MinH = 4000
End Property

Public Function OnSwitchTo(ByVal bFlShow As Boolean) As Boolean
  OnSwitchTo = True
  If bFlShow Then
  End If
End Function


Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_mg_GertNet = Nothing
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
End Sub


Private Sub UserControl_Initialize()
  Caption = "Процедуры"
End Sub

Private Sub UserControl_Terminate()
  HandsOffModel
End Sub
