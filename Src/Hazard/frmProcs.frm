VERSION 5.00
Begin VB.Form frmProcs 
   BorderStyle     =   0  'None
   Caption         =   "Процедуры"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Procs"
      Height          =   900
      Left            =   1905
      TabIndex        =   0
      Top             =   1260
      Width           =   2100
   End
End
Attribute VB_Name = "frmProcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_mg_GertNet As MGertNet

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


Private Sub Form_Unload(Cancel As Integer)
  HandsOffModel
End Sub

Public Property Get IsBoundModel() As Boolean
  IsBoundModel = Not (m_mg_GertNet Is Nothing)
End Property

Public Sub HandsOffModel()
  Set m_mg_GertNet = Nothing
End Sub

Public Sub BindModel(ByVal mgn As MGertNet)
  Set m_mg_GertNet = mgn
End Sub

