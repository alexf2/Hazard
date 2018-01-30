VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmNewCfg 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Новая конфигурация"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmNewCfg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFBF0&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   4485
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   3149
      _Version        =   131074
      BackColor       =   12632256
      PictureBackgroundStyle=   2
      ClipControls    =   0   'False
      BorderWidth     =   0
      BevelOuter      =   0
      Begin Threed.SSPanel sspI1 
         Height          =   210
         Left            =   105
         TabIndex        =   4
         Top             =   255
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   370
         _Version        =   131074
         Font3D          =   3
         ForeColor       =   0
         BackStyle       =   1
         Caption         =   "Имя конфигурации"
         ClipControls    =   0   'False
         BevelOuter      =   0
         AutoSize        =   2
         Alignment       =   0
      End
      Begin Threed.SSCommand ssCancel 
         CausesValidation=   0   'False
         Height          =   525
         Left            =   165
         TabIndex        =   2
         Top             =   1140
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmNewCfg.frx":000C
         Caption         =   "Отмена"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssOk 
         Height          =   525
         Left            =   2730
         TabIndex        =   1
         Top             =   1140
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   926
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmNewCfg.frx":0176
         Caption         =   "Подтверждение"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmNewCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

