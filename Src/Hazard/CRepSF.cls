VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRepSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IRepItem

Private Type TSFData
  sFChanges() As String
  sName As String
  sNonCompat As String
  cCost As Currency
  dQ As Double
  cProfit As Currency
  dKe As Double
End Type

Private m_arr() As TSFData

Private m_s_ComplName As String
Private m_c_AveDamage As Currency
Private m_d_P0 As Double

Private Function IsValidOpt(vArr() As TSFData) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidOpt = True
  Else
    IsValidOpt = False
  End If
End Function

Private Function IsValidSA(vArr() As String) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidSA = True
  Else
    IsValidSA = False
  End If
End Function


Public Sub IRepItem_GenerateReport(ByVal vp As Object, ByVal lNumber As Long)
  With vp
    .LineSpacing = 100
    
    .Font.Charset = 204
    .FontSize = 10
    .FontName = "Courier New"
    .FontBold = False
    .FontItalic = False
    .TextAlign = taJustMiddle
    '.Paragraph = "Комплекс мер: '" & m_s_ComplName & "'"
    .Paragraph = "Средний ущерб от возможного происшествия = " & FormatCurrency(m_c_AveDamage, 2, vbTrue, vbFalse, vbTrue)
    .Paragraph = "Базовая вероятность происшествия = " & Format(m_d_P0, haApp.FmtDbl)
    
    .TextAlign = taLeftTop
              
        
     If IsValidOpt(m_arr) Then
       Dim lMCnt As Long
       For lMCnt = LBound(m_arr) To UBound(m_arr)
        .TextAlign = taLeftTop
        .FontItalic = False
        .FontSize = 12
        .FontName = "Arial"
        .FontBold = True
        .IndentLeft = 300
        .SpaceBefore = 200
        .SpaceAfter = 150
        .Paragraph = CStr(lNumber) & "." & CStr(lMCnt + 1) & " Мера '" & m_arr(lMCnt).sName & "'"
        .SpaceBefore = 0
        .SpaceAfter = 0
        .IndentLeft = 0
                        
        .FontItalic = True
        .FontSize = 10
        .FontBold = True
        
        Dim Fmt As String, s As String
        Fmt = "^+1400|^+2100|^+1400|^+1400;"
              
        s = "Затраты|Дельта Q|Выгода|Кэ;"
        
        .PenWidth = 1
        .TextAlign = taCenterTop
      
        .FontSize = 10
        .TableBorder = 7
        .TablePen = 1
        .TablePenTB = 20
        .TablePenLR = 20
            
        .LineSpacing = 100
        
        .SpaceBefore = 50
        .SpaceAfter = 50
        vp.AddTable Fmt, s, "", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), False
        .SpaceBefore = 0
        .SpaceAfter = 0
                            
        Fmt = ">+1400|<+2100|<+1400|<+1400;"
        .FontItalic = False
        '.TablePenTB = 1
        .FontBold = False
        .LineSpacing = 100
        .SpaceBefore = 50
        .SpaceAfter = 50
                
               
        Dim sBody As String
        With m_arr(lMCnt)
                    
          sBody = FormatCurrency(.cCost, 2, vbTrue, vbFalse, vbTrue) & "|" _
            & Format(.dQ, haApp.FmtDbl) & "|" & FormatCurrency(.cProfit, 2, vbTrue, vbFalse, vbTrue) & "|" _
            & Format(.dKe, haApp.FmtDbl2) & ";"
                    
          vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          
        End With

       
        If IsValidSA(m_arr(lMCnt).sFChanges) Then
                  
          .TextAlign = taLeftTop
          .IndentLeft = 350
          .SpaceBefore = 100
          
          Dim lC2 As Long, sTmp As String, lC2To As Long
          sTmp = "Воздействие на факторы: ": lC2To = UBound(m_arr(lMCnt).sFChanges)
          For lC2 = LBound(m_arr(lMCnt).sFChanges) To lC2To
            sTmp = sTmp & m_arr(lMCnt).sFChanges(lC2)
            If lC2 <> lC2To Then sTmp = sTmp & "; "
          Next lC2
          
          vp.Paragraph = sTmp
          If m_arr(lMCnt).sNonCompat <> "" Then
            .SpaceBefore = 0
            vp.Paragraph = "Несовместимые меры: " & m_arr(lMCnt).sNonCompat
          End If
          .SpaceBefore = 100
          .IndentLeft = 0
          
        End If
        
      Next lMCnt
    End If
          
   .TablePenTB = 1
   .SpaceBefore = 0
   .SpaceAfter = 0
   .Paragraph = " "
    
  End With
End Sub

Public Property Get IRepItem_Title() As String
  IRepItem_Title = "Комплекс мер '" & m_s_ComplName & "'"
End Property

'rep.InitReport haApp, m_coll_Opt, m_s_NameCompl, m_d_P0, m_c_AveDamage, m_ot_Task, vDta
Public Sub InitReport(ByVal coll As collSF, ByVal sName As String, ByVal dP0 As Double, ByVal ha As HazardApp)

  Dim srtOld As CollSFSorting
  srtOld = coll.Sorting

  m_s_ComplName = sName
  m_c_AveDamage = haApp.GertNetMain.AverageDamage
  m_d_P0 = dP0


  If coll.Count < 1 Then Exit Sub
  
  On Error GoTo ERR_H
  
  coll.Sorting = frmSP.CurrSorting
  
  ReDim m_arr(coll.Count - 1) As TSFData
  Dim l As Long: l = 1
  Dim sfc As SafetyPrecaution
  For Each sfc In coll
    'Dim sfc As SafetyPrecaution: Set sfc = coll(l)
    
    With m_arr(l - 1)
      .cCost = sfc.Cost
      .dQ = sfc.DeltaQ
      .cProfit = sfc.Profit
      .dKe = sfc.Ke
      .sName = sfc.Name
      If sfc.FChanges.Count > 0 Then
        ReDim .sFChanges(sfc.FChanges.Count - 1) As String
        Dim fc As FChange, lCnt As Long
        lCnt = 0
        For Each fc In sfc.FChanges
          .sFChanges(lCnt) = fc.NameFactor & "(" & TSign(fc.Delta) & CStr(fc.Delta) & StrFromTCH(fc.tch) & ")"
          lCnt = lCnt + 1
        Next fc
      End If
      
      .sNonCompat = ""
      
      Dim lNon As Long, larrComp() As Long
      larrComp = sfc.NonCompatibles
      For lNon = LBound(larrComp) To UBound(larrComp)
        Dim csf As SafetyPrecaution
        Set csf = coll(larrComp(lNon))
        .sNonCompat = .sNonCompat & csf.Name
        If lNon <> UBound(larrComp) Then
          .sNonCompat = .sNonCompat & "; "
        End If
      Next lNon
            
    End With
    l = l + 1
  Next sfc
  
  coll.Sorting = srtOld
  Exit Sub
ERR_H:
  
  coll.Sorting = srtOld
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub






