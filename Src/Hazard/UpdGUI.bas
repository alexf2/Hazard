Attribute VB_Name = "UpdGUI"
Option Explicit

Public Sub UpdateGUI(ByVal hzdApp As HazardApp)
  If frmMain Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub
  UpdateGUI_MonModel hzdApp
  UpdateGUI_MonSF hzdApp
  UpdateGUI_Reports hzdApp
End Sub

Public Sub UpdateGUI_MonModel(ByVal hzdApp As HazardApp)
  If frmMain Is Nothing Then Exit Sub
  If hzdApp Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub

  If hzdApp.GertNetMain Is Nothing Then
    If Not MMNothing() Then
      frxModelMon.HandsOffModel
      DoEvents
    End If
  Else
    If MMNothing() Then
      Load frxModelMon
      DoEvents
    Else
    End If
    frxModelMon.BindModel hzdApp.GertNetMain
    DoEvents
    frxModelMon.Visible = True
    DoEvents
  End If
  
End Sub

Public Function MMNothing() As Boolean

  MMNothing = frmCalibrate Is Nothing Or _
    frmMEditor Is Nothing Or _
    frmMon Is Nothing
    'frmProcs Is Nothing

End Function

Public Function SPNothing() As Boolean

  SPNothing = frmComplSP Is Nothing Or _
    frmOptimiz Is Nothing Or _
    frmSP Is Nothing
        
End Function


Public Sub UpdateGUI_MonSF(ByVal hzdApp As HazardApp)
  If frmMain Is Nothing Then Exit Sub
  If hzdApp Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub

  If SPNothing() Then Load frxSPMonitor
  DoEvents
  
  Dim bTOff As Boolean
  bTOff = hzdApp.XCollection Is Nothing
  If Not bTOff Then bTOff = hzdApp.XCollection.Count < 1
  
  If bTOff Then
    If Not SPNothing() Then
      frxSPMonitor.HandsOffModel
      DoEvents
    End If
  Else
    If SPNothing() Then
      Load frxSPMonitor
      DoEvents
    Else
    End If
    frxSPMonitor.BindModel hzdApp.XCollection
    DoEvents
    frxSPMonitor.Visible = True
    DoEvents
  End If
  
End Sub

Public Sub EnblControls(ByVal f As Form, ByVal bFl As Boolean, Optional collExclude As CollectionX = Nothing)
  If collExclude Is Nothing Then
    Dim ctl As Control
    For Each ctl In f.Controls
      If f.Enabled <> bFl Then _
        f.Enabled = bFl
    Next ctl
    Exit Sub
  End If
    
  For Each ctl In f.Controls
    If IsNull(collExclude.Key(CStr(ObjPtr(ctl)))) Then _
      If f.Enabled <> bFl Then _
        f.Enabled = bFl
  Next ctl
End Sub

Public Sub EnblControls2(arrCtl() As Control, ByVal bFl As Boolean)
  If UBound(arrCtl) = -1 Then Exit Sub
  Dim l As Long
  For l = UBound(arrCtl) To LBound(arrCtl) Step -1
    If arrCtl(l).Enabled <> bFl Then _
      arrCtl(l).Enabled = bFl
  Next l
End Sub

Public Function IsCalcAny(ByVal mgn As MGertNet) As Boolean
  If mgn Is Nothing Then IsCalcAny = False: Exit Function
  
  If mgn.IsRunning Then
    IsCalcAny = True
  ElseIf mgn.IsRunning2 Then
    IsCalcAny = True
  ElseIf mgn.IsRunningOpt Then
    IsCalcAny = True
  Else
    IsCalcAny = False
  End If
End Function

Public Sub UpdateGUI_Reports(ByVal hzdApp As HazardApp)

  UpdateGUI_Rep1 hzdApp
  UpdateGUI_Rep2 hzdApp
    
End Sub

Public Sub UpdateGUI_Rep1(ByVal hzdApp As HazardApp)
  If frmMain Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub
  
  UpdateGUI_RepXX hzdApp.Rep1, frxRep.vsModel
End Sub

Public Sub UpdateGUI_Rep2(ByVal hzdApp As HazardApp)
  If frmMain Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub
  
  UpdateGUI_RepXX hzdApp.Rep2, frxRep.vpSP
End Sub

Public Sub UpdateGUI_RepXX(ByVal rep As CollectionX, ByVal vs As VSPrinter)
  If frmMain Is Nothing Then Exit Sub
  If Not frmMain.Visible Then Exit Sub
  
  If rep Is Nothing Then
    vs.StartDoc
    vs.EndDoc
    frxRep.UpdateControlBar
    Exit Sub
  End If
  
  Dim bm As New CBeam
  bm.Beam frmMain
  
  With vs
     
     Dim lKeyPreview As Long
     lKeyPreview = .PreviewPage
    '.Preview = True
    
    .StartDoc
              
     Dim iRep As IRepItem, l As Long
     l = 1
     For Each iRep In rep
        .Font.Charset = 204
        .FontName = "Arial"
        .FontBold = False
        .FontItalic = False
        .TextAlign = taLeftTop
        .TableBorder = tbAll
        .PageBorder = pbTopBottom
        .PenStyle = psSolid
        .BrushStyle = bsSolid
        .PenWidth = 2
        .PenColor = 0
        .BrushColor = 0
        .TextColor = 0
        .Columns = 1
    
      .FontSize = 14
      .Paragraph = CStr(l) & ".  " & iRep.Title
      .FontSize = 5
      .Paragraph = ""
      
      iRep.GenerateReport vs, l
      
      l = l + 1
     Next iRep
      
    
    .EndDoc
    .PreviewPage = IIf(lKeyPreview >= .PageCount, lKeyPreview, .PageCount)
  End With
  
  frxRep.UpdateControlBar
End Sub

