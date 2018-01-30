VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRepVApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IRepItem

Private Type TSPData
  sName As String
  cCost As Currency
  dQ As Double
  cProfit As Currency
  dKe As Double
End Type

Private m_arr() As TSPData

Private m_s_ComplName As String

Private m_c_AveDamage As Currency

Private m_c_CostSumm As Currency
Private m_c_ProfitSumm As Currency
Private m_d_QSumm As Double

Private m_c_ProfitReal As Currency
Private m_d_QReal As Double
Private m_d_KeReal As Double

Private m_d_New As Double
Private m_d_Old As Double

Private Type TFacData
  sNameShort As String
  sName As String
  sVal1 As String
  sVal2 As String
  iDelta As Integer
  iOver As Integer
  sTrust1 As String
  sTrust2 As String
End Type

Private m_arr_Fac() As TFacData


Private Function IsValidFac(vArr() As TFacData) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidFac = True
  Else
    IsValidFac = False
  End If
End Function

Private Function IsValidSP(vArr() As TSPData) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidSP = True
  Else
    IsValidSP = False
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
    .Paragraph = "Комплекс мер: '" & m_s_ComplName & "'"
    .Paragraph = "Средний ущерб от возможного происшествия = " & FormatCurrency(m_c_AveDamage, 2, vbTrue, vbFalse, vbTrue)
    .Paragraph = "Вероятность происшествия в исходной системе = " & Format(m_d_Old, haApp.FmtDbl)
    .Paragraph = "Вероятность происшествия после применения набора мер = " & Format(m_d_New, haApp.FmtDbl)
    .TextAlign = taLeftTop
    
        
    .FontItalic = False
    .FontSize = 12
    .FontName = "Arial"
    .FontBold = True
    .IndentLeft = 300
    .SpaceBefore = 200
    .SpaceAfter = 150
    .Paragraph = CStr(lNumber) & ".1." & " Меры"
    .SpaceBefore = 0
    .SpaceAfter = 0
    .IndentLeft = 0
                    
    .FontItalic = True
    .FontSize = 10
    .FontBold = True
    
    Dim Fmt As String, s As String
    Fmt = "^+2500|^+1400|^+2100|^+1400|^+1100;"
          
    s = "Мера|Затраты|Дельта Q|Выгода|Кэ;"
      
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
                
    Fmt = ">+2500|<+1400|<+2100|<+1400|<+1100;"
    .FontItalic = False
    .TablePenTB = 1
    .FontBold = False
    .LineSpacing = 100
    .SpaceBefore = 50
    .SpaceAfter = 50
    
    Dim k As Long
    If IsValidSP(m_arr) Then
      For k = LBound(m_arr) To UBound(m_arr)
        Dim sBody As String
        With m_arr(k)
          
          sBody = .sName & "|" & FormatCurrency(.cCost, 2, vbTrue, vbFalse, vbTrue) & "|" _
            & Format(.dQ, haApp.FmtDbl) & "|" & FormatCurrency(.cProfit, 2, vbTrue, vbFalse, vbTrue) & "|" _
            & Format(.dKe, haApp.FmtDbl2) & ";"
                    
          vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          
        End With
      Next k
    End If
    
    .TablePenTB = 20
    sBody = "Суммарные значения" & "|" & FormatCurrency(m_c_CostSumm, 2, vbTrue, vbFalse, vbTrue) & "|" _
          & Format(m_d_QSumm, haApp.FmtDbl) & "|" & FormatCurrency(m_c_ProfitSumm, 2, vbTrue, vbFalse, vbTrue) & "|" _
          & "-" & ";"
    Fmt = "^+2500|<+1400|<+2100|<+1400|<+1100;"
    .FontBold = True
    .AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
    
    .TablePenTB = 1
    .FontBold = False
    Fmt = "^+2500|^+1400|<+2100|<+1400|<+1100;"
    sBody = "Реальные значения (прогон)" & "|" & "-//-" & "|" _
          & Format(m_d_QReal, haApp.FmtDbl) & "|" & FormatCurrency(m_c_ProfitReal, 2, vbTrue, vbFalse, vbTrue) & "|" _
          & Format(m_d_KeReal, haApp.FmtDbl2) & ";"
    .AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
    
    
    
    
        
    .TextAlign = taLeftTop
    .FontItalic = False
    .FontSize = 12
    .FontName = "Arial"
    .FontBold = True
    .IndentLeft = 300
    .SpaceBefore = 200
    .SpaceAfter = 150
    .Paragraph = CStr(lNumber) & ".2." & " Воздействие на модель"
    .SpaceBefore = 0
    .SpaceAfter = 0
    .IndentLeft = 0
                    
    .FontItalic = True
    .FontSize = 10
    .FontBold = True
        
    Fmt = "^+1000|^+2100|^+2100|^+900|^+900|^+2000|^+2000;"
          
    s = "Фактор|Оценка до|Оценка после|Изменение|Переполнение|Доверие до|Доверие после;"
    
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
                    
    Fmt = ">+1000|<+2100|<+2100|<+900|^+900|<+2000|<+2000;"
    .FontItalic = False
    .TablePenTB = 1
    .FontBold = False
    .LineSpacing = 100
    .SpaceBefore = 50
    .SpaceAfter = 50
        
    If IsValidFac(m_arr_Fac) Then
      For k = LBound(m_arr_Fac) To UBound(m_arr_Fac)
        With m_arr_Fac(k)
          
          sBody = .sNameShort & "|" & .sVal1 & "|" _
            & .sVal2 & "|" & TSign(.iDelta) & CStr(Abs(.iDelta)) & "|" _
            & IIf(.iOver = 0, "-", CStr(.iOver)) & "|" & .sTrust1 & "|" & .sTrust2 & ";"
                    
          If k = UBound(m_arr_Fac) Then vp.TablePenTB = 20
          vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          
        End With
      Next k
    End If
          
   .SpaceBefore = 0
   .SpaceAfter = 0
   .Paragraph = " "
    
  End With
End Sub

Public Property Get IRepItem_Title() As String
  IRepItem_Title = "Оценка эффективности набора мер"
End Property

Public Sub InitReport(ByVal mgn As MGertNet, ByVal ha As HazardApp, ByVal mgnBase As MGertNet, ByVal collTmp As collSF, ByVal cSummProfit As Currency, ByVal dSummQ As Double, ByVal sComplName As String, ByVal dPNew As Double, ByVal dPOld As Double)

  m_c_AveDamage = mgn.AverageDamage
  m_s_ComplName = sComplName
  
  m_d_New = dPNew
  m_d_Old = dPOld

  If collTmp.Count > 0 Then
    ReDim m_arr(collTmp.Count - 1) As TSPData
    Dim sf As SafetyPrecaution, l As Long
    l = 0
    For Each sf In collTmp
      With m_arr(l)
        .cCost = sf.Cost
        .cProfit = sf.Profit
        .dKe = sf.Ke
        .dQ = sf.DeltaQ
        .sName = sf.Name
      End With
      l = l + 1
    Next sf
  End If
  
  With collTmp
    m_c_CostSumm = .Cost
    m_c_ProfitSumm = cSummProfit
    m_d_QSumm = dSummQ
    
    m_c_ProfitReal = .Profit
    m_d_QReal = .DeltaQ
    m_d_KeReal = .Ke
  End With
  
  Dim f As Factor, ibs As IBSTRKey
  l = 0
  For Each f In mgn.Factors
    With f
      Set ibs = f
      If .Value <> mgnBase.Factors(ibs.Get()).Value Then
        l = l + 1
      End If
    End With
  Next f
  
  If l > 0 Then
    ReDim m_arr_Fac(l - 1) As TFacData
    
    l = 0
    For Each f In mgn.Factors
      With f
        Set ibs = f
        Dim f0 As Factor
        Set f0 = mgnBase.Factors(ibs.Get())
        If .Value <> f0.Value Then
          With m_arr_Fac(l)
            .iOver = f.Overwrap
            .iDelta = f.Value - f0.Value
            .sName = f.Name
            .sNameShort = ibs.Get()
            
            mgn.GetFactorPresent f, .sVal2, .sTrust2
            mgnBase.GetFactorPresent f0, .sVal1, .sTrust1
          End With
        
          l = l + 1
          
        End If
      End With
    Next f
    
  End If
  
End Sub


