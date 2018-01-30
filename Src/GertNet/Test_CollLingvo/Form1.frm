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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clLingvo As GERTNETLib.CollLingvo

Private Sub Form_Load()
  On Error GoTo ERRH
  
  'GoTo TST01
  
  'Test_CollLingvo
  'Test_CollFac
  Test_GNet
  
  
  
  '---------------------
'TST01:
  
  Exit Sub
ERRH:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub DumpColl(ByVal refCL As GERTNETLib.CollLingvo)
  Debug.Print
  Dim vvt As Variant
  Dim lv As GERTNETLib.LingvoEnum
  For Each vvt In refCL
    Set lv = vvt
    Dim i As Long
    For i = 0 To lv.Count - 1
      Dim lk As ILongKey
      Set lk = lv
      With lv
        Debug.Print lk.Get, .Mark(i), .Quality(i)
      End With
    Next i
    Debug.Print
  Next vvt
End Sub

Private Sub DumpFac(ByVal refCL As GERTNETLib.CollFac)
  Debug.Print
  Dim vvt As Variant
  Dim lv As GERTNETLib.Factor
  For Each vvt In refCL
          
      Set lv = vvt
      Dim lk As IBSTRKey
      Set lk = lv
      With lv
        Debug.Print lk.Get(), .IDEnum, .Name, .Value
      End With
        
  Next vvt
End Sub

Private Sub Test_CollLingvo()
  Set clLingvo = New GERTNETLib.CollLingvo
  Dim ipersCL As IPersistStorage
  Set ipersCL = clLingvo
  ipersCL.InitNew Nothing
  
  Dim lInd As Long
  lInd = clLingvo.LookNextIndex
  Debug.Print lInd
  
  clLingvo.Add New GERTNETLib.LingvoEnum, -1
  clLingvo.Add New GERTNETLib.LingvoEnum, -1
  clLingvo.Add New GERTNETLib.LingvoEnum, -1
  lInd = clLingvo.Add(New GERTNETLib.LingvoEnum, -1)
  lInd = clLingvo.Add(New GERTNETLib.LingvoEnum, -1)
  
  lInd = clLingvo.LookNextIndex
  Debug.Print lInd
  lInd = clLingvo.Count
  Debug.Print lInd
  
  Dim leEnum(1 To 6) As GERTNETLib.LingvoEnum
  Set leEnum(1) = clLingvo.Item(1)
  Set leEnum(2) = clLingvo.Item(2)
  Set leEnum(3) = clLingvo.Item(3)
  Set leEnum(4) = clLingvo.Item(4)
  Set leEnum(5) = clLingvo.Item(5)
  'Set leEnum(6) = clLingvo.Item(6)
  
  leEnum(2).Mark(5) = leEnum(2).Mark(5) + "++"
  leEnum(3).Mark(1) = leEnum(3).Mark(1) + "--"
  
  Debug.Print
  lInd = clLingvo.GetIndexOfItem(leEnum(4))
  Debug.Print lInd
  lInd = clLingvo.GetIndexOfItem(leEnum(3))
  Debug.Print lInd
  lInd = clLingvo.GetIndexOfItem(leEnum(1))
  Debug.Print lInd
  
  DumpColl clLingvo
  
  clLingvo.Remove 2
  Set leEnum(2) = Nothing
  Debug.Print "After removing"
  DumpColl clLingvo
  
  Dim vv As Variant
  Set vv = clLingvo(3)
  
  'Debug.Print clLingvo.Item(2).Mark(1)
  'clLingvo.Remove 2
  
  Dim stg As IStorage
  Set stg = CreateCF("g:\tst_stg", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  ipersCL.Save stg, True
  
  leEnum(4).Mark(1) = leEnum(4).Mark(1) + "&&"
  
  ipersCL.Save stg, True
  
  Set clLingvo = Nothing
  
  Set clLingvo = New GERTNETLib.CollLingvo
  Set ipersCL = clLingvo
  Set stg = Nothing
  Set stg = OpenCF("g:\tst_stg", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ipersCL.Load stg
  
  DumpColl clLingvo
  
  clLingvo.Add New GERTNETLib.LingvoEnum, -1
  clLingvo(6).Mark(1) = clLingvo(6).Mark(1) + " _ABC_ "
  ipersCL.Save stg, True
  
  DumpColl clLingvo
End Sub
Private Sub Test_CollFac()
  Dim clFac As New GERTNETLib.CollFac
  Dim ipersCL As IPersistStorage
  Set ipersCL = clFac
  ipersCL.InitNew Nothing
  
  clFac.Add New Factor, "M01"
  clFac.Add New Factor, "M02"
  clFac.Add New Factor, "M03"
  clFac.Add New Factor, "M04"
  clFac.Add New Factor, "M05"
  
  clFac("M01").Name = "PP1"
  clFac("M02").Name = "PP2"
  clFac("M03").Name = "PP3"
  clFac("M04").Name = "PP4"
  clFac("M05").Name = "PP5"
  'clFac("M05_").Name = "PP5"
  'clFac.Remove "M01_"
  DumpFac clFac
End Sub

Private Sub Test_GNet()
  Dim mgn As New GERTNETLib.MGertNet
  Dim ipersCL As IPersistStorage
  Set ipersCL = mgn
  ipersCL.InitNew Nothing
  
  Dim mm As Long
  mm = mgn.Enumerators.Add(New LingvoEnum, -1)
  mgn.Factors.Add New Factor, "M01"
  mgn.Factors("M01").IDEnum = mm
  
  Dim stg As IStorage
  Set stg = CreateCF("g:\tst_stg", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  ipersCL.Save stg, True
  
  Set mgn = Nothing
  Set stg = Nothing
  Set stg = OpenCF("g:\tst_stg", STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  Set mgn = New GERTNETLib.MGertNet
  Set ipersCL = mgn
  ipersCL.Load stg
  
  DumpColl mgn.Enumerators
  DumpFac mgn.Factors
  
  Dim sVN As String, sTL As String, dVal As Double
  mgn.GetFactorPresent mgn.Factors("M01"), sVN, sTL, dVal
  Debug.Print sVN, sTL, dVal
End Sub
