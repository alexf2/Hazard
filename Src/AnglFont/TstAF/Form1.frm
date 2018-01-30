VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   1350
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim byteArr() As Byte
  Dim vvt
  Open "G:\pb.tst" For Binary As #1
  Get #1, , vvt
  Close #1
  
  byteArr = vvt
  Set pb = New PropertyBag
  pb.Contents = byteArr
    
  Set ff = pb.ReadProperty("pp")
  Debug.Print ff.Angle
End Sub

Private Sub Form_Load()
  
  Dim ff As New AngledFont
  Dim pb As New PropertyBag
  ff.Angle = 45
    
  pb.WriteProperty "pp", pb
  Dim vvt
  vvt = pb.Contents
  
  Open "G:\pb.tst" For Binary As #1
  Put #1, , vvt
  Close #1

End Sub
