VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand ssClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   1635
      TabIndex        =   2
      ToolTipText     =   "Вносит изменения в оценки факторов опасности модели"
      Top             =   1822
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      _Version        =   131074
      PictureFrames   =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmAbout.frx":000C
      Caption         =   "Закрыть"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "(c)AlexCorp. 2000"
      Height          =   240
      Left            =   1080
      TabIndex        =   1
      Top             =   1140
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Универсальный экспертный модуль оценки значений факторов опасности методом средневзвешенного"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   375
      TabIndex        =   0
      Top             =   337
      Width           =   3675
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ssClose_Click()
  Unload Me
End Sub
