VERSION 5.00
Begin VB.Form frmAboutRailway 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "� ���������"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "�������"
      Height          =   435
      Left            =   1726
      TabIndex        =   0
      Top             =   1793
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(c)AlexCorp. 2000 ���������� ������ ������ �������� �������� ��������� ��� �������� �����"
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
      Left            =   510
      TabIndex        =   1
      Top             =   375
      Width           =   3675
   End
End
Attribute VB_Name = "frmAboutRailway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub
