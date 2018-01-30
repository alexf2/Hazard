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

Private Sub Form_Load()

  On Error GoTo ERR_H

  Dim mgn As New GERTNETLib.MGertNet
  Dim ips As GERTNETLib.IPersistStorage
  Set ips = mgn
  ips.InitNew Nothing
  
  Dim leEnum As New GERTNETLib.LingvoEnum
  Dim lID As Long
  lID = mgn.Enumerators.Add(leEnum, -1)
  
  Dim fFactor As New Factor
  fFactor.IDEnum = lID
  fFactor.Name = "Factor 1"
  fFactor.TrustLvl = TL_Normal
  fFactor.Value = 5
  mgn.Factors.Add fFactor, "M01"
  
  Set fFactor = New Factor
  fFactor.IDEnum = lID
  fFactor.Name = "Factor 2"
  fFactor.TrustLvl = TL_Normal
  fFactor.Value = 7
  mgn.Factors.Add fFactor, "M02"
  
  Set fFactor = New Factor
  fFactor.IDEnum = lID
  fFactor.Name = "Factor 3"
  fFactor.TrustLvl = TL_Normal
  fFactor.Value = 2
  mgn.Factors.Add fFactor, "M03"
  
  Set fFactor = New Factor
  fFactor.IDEnum = lID
  fFactor.Name = "Factor 4"
  fFactor.TrustLvl = TL_Normal
  fFactor.Value = 4
  mgn.Factors.Add fFactor, "M04"
  
  Set fFactor = New Factor
  fFactor.IDEnum = lID
  fFactor.Name = "Factor 5"
  fFactor.TrustLvl = TL_Normal
  fFactor.Value = 5
  mgn.Factors.Add fFactor, "M05"
  
  Dim cSF As New CollSF
  Set ips = cSF
  ips.InitNew Nothing
  Dim spSafety As New SafetyPrecaution
  
  spSafety.Cost = 100
  'spSafety.DeltaQ = 1
  spSafety.Name = "Мера 1"
  spSafety.FChanges.Add New FChange, 1
  spSafety.FChanges(1).NameFactor = "M04"
  spSafety.FChanges(1).Delta = 1
  cSF.Add spSafety, -1
  spSafety.FChanges.Add New FChange, 2
  spSafety.FChanges(2).NameFactor = "M03"
  spSafety.FChanges(2).Delta = 3
  
  Set spSafety = New SafetyPrecaution
  spSafety.Cost = 1000
  'spSafety.DeltaQ = 1
  spSafety.Name = "Мера 2"
  spSafety.FChanges.Add New FChange, 1
  spSafety.FChanges(1).NameFactor = "M01"
  spSafety.FChanges(1).Delta = 2
  cSF.Add spSafety, -1
  
  Dump mgn.Factors
  mgn.ApplySF cSF
  Dump mgn.Factors
  
  mgn.Run 10, 100
  
  Dim stg As IStorage
  Set stg = CreateCF("g:\tst_stg", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  Set ips = mgn
  ips.Save stg, 1
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly, "Error"
  
End Sub

Private Sub Dump(ByVal collFac As GERTNETLib.collFac)
  Debug.Print "ID", "IDEnum", "Name", "TrustLvl", "Value"
  Dim vvt
  For Each vvt In collFac
    Dim f As Factor
    Set f = vvt
    With f
      Debug.Print .ID, .IDEnum, .Name, .TrustLvl, .Value
    End With
  Next vvt
  Debug.Print
End Sub
