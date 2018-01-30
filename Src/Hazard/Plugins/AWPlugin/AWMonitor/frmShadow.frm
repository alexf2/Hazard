VERSION 5.00
Begin VB.Form frmShadow 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGOSTEdit 
      Caption         =   "�������� ������"
      Begin VB.Menu mnuItem 
         Caption         =   "Item"
         Enabled         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddTopic 
         Caption         =   "�������� ������"
      End
      Begin VB.Menu mnuAddGOST 
         Caption         =   "�������� ��������"
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "�������������	F2"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "�������������	F3"
      End
      Begin VB.Menu s11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "�������	Del"
      End
   End
   Begin VB.Menu mnuTblEdit 
      Caption         =   "�������� ������"
      Begin VB.Menu mnuInsBefore 
         Caption         =   "�������� ��"
      End
      Begin VB.Menu mnuInsAfter 
         Caption         =   "�������� �����"
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "�������	Del"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColCap 
         Caption         =   "�������� ��� ���������"
      End
      Begin VB.Menu mnuUniformGrad 
         Caption         =   "����������� ����"
      End
      Begin VB.Menu mnuUniformGradRevers 
         Caption         =   "�������� ����������� ����"
      End
   End
   Begin VB.Menu mnuFacEdit 
      Caption         =   "�������� ��������"
      Begin VB.Menu mnuCreFac 
         Caption         =   "������� �����"
      End
      Begin VB.Menu mnuRenameFac 
         Caption         =   "�������������	F2"
      End
      Begin VB.Menu mnuEditFac 
         Caption         =   "�������������	F3"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveFac 
         Caption         =   "�������	Del"
      End
   End
   Begin VB.Menu mnuTabEditor 
      Caption         =   "�������� ������"
      Begin VB.Menu mnuTAddNew 
         Caption         =   "�������� ����� ����"
      End
      Begin VB.Menu mnuTRename 
         Caption         =   "�������������	F2"
      End
      Begin VB.Menu mnuTAssGOST 
         Caption         =   "��������� ��������"
      End
      Begin VB.Menu mnuTAssTable 
         Caption         =   "��������� �������"
      End
      Begin VB.Menu mnuTAssEmpty 
         Caption         =   "��������"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTRefresh 
         Caption         =   "�������� ���������	F5"
      End
      Begin VB.Menu mnuTSave 
         Caption         =   "���������"
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTAssWeights 
         Caption         =   "���������� �����"
      End
      Begin VB.Menu mnuTUniform 
         Caption         =   "����������� ����"
      End
      Begin VB.Menu mnuTCheckW 
         Caption         =   "��������� ����"
      End
      Begin VB.Menu mnuTNormalize 
         Caption         =   "������������� ����	F10"
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTRemove 
         Caption         =   "������� 	Del"
      End
   End
End
Attribute VB_Name = "frmShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAddGOST_Click()
  frmGostEdit.PostAddGOST
End Sub

Private Sub mnuAddTopic_Click()
  frmGostEdit.PostAddTopic
End Sub


Private Sub mnuEdit_Click()
  frmGostEdit.PostCmEdit
End Sub


Private Sub mnuRemove_Click()
  frmGostEdit.PostCmRemove
End Sub


Private Sub mnuRename_Click()
  frmGostEdit.PostCmRename
End Sub

Private Sub mnuColCap_Click()
  frmGOST.PostColCap
End Sub

Private Sub mnuDelete_Click()
  frmGOST.PostDelete
End Sub

Private Sub mnuInsAfter_Click()
  frmGOST.PostInsAfter
End Sub

Private Sub mnuInsBefore_Click()
  frmGOST.PostInsBefore
End Sub

Private Sub mnuTAddNew_Click()
  frmEditFT.PostTAddNew
End Sub

Private Sub mnuTAssEmpty_Click()
  frmEditFT.PostTAssEmpty
End Sub

Private Sub mnuTAssGOST_Click()
  frmEditFT.PostTAssGOST
End Sub

Private Sub mnuTAssTable_Click()
  frmEditFT.PostTAssTable
End Sub

Private Sub mnuTAssWeights_Click()
  frmEditFT.PostTAssWeights
End Sub

Private Sub mnuTCheckW_Click()
  frmEditFT.PostTCheckW
End Sub

Private Sub mnuTNormalize_Click()
  frmEditFT.PostTNormalize
End Sub

Private Sub mnuTRefresh_Click()
  frmEditFT.PostTRefresh
End Sub

Private Sub mnuTRemove_Click()
  frmEditFT.PostTRemove
End Sub

Private Sub mnuTRename_Click()
  frmEditFT.PostTRename
End Sub

Private Sub mnuTSave_Click()
  frmEditFT.PostTSave
End Sub

Private Sub mnuTUniform_Click()
  frmEditFT.PostTUniform
End Sub

Private Sub mnuUniformGrad_Click()
  frmGOST.PostUniformGrad
End Sub

Private Sub mnuCreFac_Click()
  frmFactors.PostCreFac
End Sub

Private Sub mnuEditFac_Click()
  frmFactors.PostEditFac
End Sub

Private Sub mnuRemoveFac_Click()
  frmFactors.PostRemoveFac
End Sub

Private Sub mnuRenameFac_Click()
  frmFactors.PostRenameFac
End Sub

Private Sub mnuUniformGradRevers_Click()
  frmGOST.PostUniformGradRevers
End Sub
