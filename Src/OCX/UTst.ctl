VERSION 5.00
Object = "{CBC5C1A6-2402-11D4-8F15-00504E02C39D}#104.0#0"; "AlexOCX.ocx"
Object = "{CC52DF59-28C5-11D4-8F1B-00504E02C39D}#67.0#0"; "GNControls.ocx"
Begin VB.UserControl UTst 
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   ScaleHeight     =   5625
   ScaleWidth      =   6285
   Begin GNControls.CtlRepeater CtlRepeater1 
      Height          =   3600
      Left            =   165
      TabIndex        =   0
      Top             =   1860
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   6350
   End
   Begin AlexOCX.Bulb Bulb1 
      Height          =   420
      Left            =   480
      Top             =   975
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   741
   End
End
Attribute VB_Name = "UTst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get PCtlRepeater1() As CtlRepeater
  Set PCtlRepeater1 = CtlRepeater1
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property
