VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmPlugin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Модуль"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   795
   ControlBox      =   0   'False
   Icon            =   "frmPlugin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   795
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand ssCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   217
      TabIndex        =   0
      ToolTipText     =   "Выгрузить модуль"
      Top             =   60
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmPlugin.frx":000C
      AutoSize        =   2
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "frmPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_tIfass As IFactorAssign

Private Sub Form_Unload(Cancel As Integer)
  Set m_tIfass = Nothing
End Sub

Private Sub ssCancel_Click()
  Unload Me
End Sub
