VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRepRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IRepItem


Private m_s_FileName As String
Private m_l_N As Long, m_l_K As Long
Private m_mt_Alho As GERTNETLib.ModelType
Private m_dt_Time As Date
Private m_dt_TimeOld As Date

Private m_RunMode As GERTNETLib.ModelType

Private Type TStateDta
  sName As String
  dMx As Double
  dDx As Double
  lMin As Long
  lMax As Long
  
  dIdxMin As Double
  dIdxMax As Double
  dIdxMx As Double
  dIdxDx As Double
End Type

Private m_arr() As TStateDta
Private m_arrOld() As TStateDta

Private Type TFactorsData
  sShortName As String
  sName As String
  sVal As String
  sValOld As String
  dVal As Double
  dValOld As Double
  iDelta As Integer
  trLvl As GERTNETLib.TrustLevel
  trLvlOld As GERTNETLib.TrustLevel
  sLvl As String
  sLvlOld As String
End Type

Private m_arrFacCh() As TFactorsData

Private Function AlhoDescr(ByVal dscr As GERTNETLib.ModelType) As String
  Select Case dscr
    Case MT_Analytical
      AlhoDescr = "Аналитический"
    Case MT_Imitate
      AlhoDescr = "Имитационный"
    Case MT_ImitateIndistinct
      AlhoDescr = "Имитационный устойчивый"
    Case MT_AnalyticalIndistinct
      AlhoDescr = "Аналитический устойчивый"
    Case MT_Analytical2
      AlhoDescr = "Аналитический устойчивый приближённый"
  End Select
End Function

Private Function IsValidTFD(vArr() As TFactorsData) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidTFD = True
  Else
    IsValidTFD = False
  End If
End Function

Private Function IsValidTSD(vArr() As TStateDta) As Boolean
  On Error Resume Next
  Dim l As Long
  l = UBound(vArr)
  If l <> -1 And Err.Number = 0 Then
    IsValidTSD = True
  Else
    IsValidTSD = False
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
    .TextAlign = taLeftTop
    .Paragraph = "Алгоритм: '" & AlhoDescr(m_mt_Alho) & "'"
    If m_mt_Alho <> MT_Analytical2 And m_mt_Alho <> MT_Analytical And m_mt_Alho <> MT_AnalyticalIndistinct Then _
      .Paragraph = "Параметры прогона: N = " & CStr(m_l_N) & ", K = " & CStr(m_l_K)
    
    .Paragraph = "Время счёта: " & Format(m_dt_Time, "Long Time")
    If IsValidTSD(m_arrOld) Then _
      .Paragraph = "Разность времён счёта: " & Format(m_dt_Time - m_dt_TimeOld, "Long Time")
            
    .FontItalic = False
    .FontSize = 12
    .FontName = "Arial"
    .FontBold = True
    .IndentLeft = 300
    .SpaceBefore = 200
    .SpaceAfter = 150
    .Paragraph = CStr(lNumber) & ".1." & " Результаты моделирования"
    .SpaceBefore = 0
    .SpaceAfter = 0
    .IndentLeft = 0
                    
    .FontItalic = True
    .FontSize = 10
    .FontBold = True
    
    Dim Fmt As String, s As String
    Fmt = "^+1600|^+2000|^+2000|^+1100|^+1100|^+1100|^+2000;"
          
    
    s = "Состояние|Частота|Среднее кв. отклонение|Min|Max|Матем. ожидание|Ср.кв.отклонение;"
    
     
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
    vp.AddTable "^+5600|^+5300;", "Состояние|Индекс опасности;", "", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), False
    vp.AddTable Fmt, s, "", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), False
    .SpaceBefore = 0
    .SpaceAfter = 0
            
    Fmt = ">+1600|<+2000|<+2000|>+1100|>+1100|>+1100|<+2000;"
    .FontItalic = False
    .TablePenTB = 1
    .FontBold = False
    .LineSpacing = 100
    .SpaceBefore = 50
    .SpaceAfter = 50
    
    Dim k As Long
    For k = LBound(m_arr) To UBound(m_arr)
      Dim sBody As String
      With m_arr(k)
        
        If m_RunMode = MT_Imitate Or m_RunMode = MT_ImitateIndistinct Then
          sBody = .sName & "|" & Format(.dMx, haApp.FmtDbl) & "|" _
            & Format(.dDx, haApp.FmtDbl) & "|" & Format(.dIdxMin, haApp.FmtDbl2) & "|" _
            & Format(.dIdxMax, haApp.FmtDbl2) & "|" & Format(.dIdxMx, haApp.FmtDbl2) & "|" & Format(.dIdxDx, haApp.FmtDbl2) & ";"
        Else
          sBody = .sName & "|" & Format(.dMx, haApp.FmtDbl) & "|" _
            & "-" & "|" & Format(.dIdxMin, haApp.FmtDbl2) & "|" _
            & Format(.dIdxMax, haApp.FmtDbl2) & "|" & Format(.dIdxMx, haApp.FmtDbl2) & "|" & Format(.dIdxDx, haApp.FmtDbl2) & ";"
        End If
        
        If k = UBound(m_arr) Then vp.TablePenTB = 20
        vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
        
      End With
    Next k
    
    Dim iCntPara As Integer
    iCntPara = 2
            
    If IsValidTFD(m_arrFacCh) Then
      .SpaceBefore = 0
      .SpaceAfter = 0
      .Paragraph = ""
      
      .FontItalic = False
      .FontSize = 12
      .FontName = "Arial"
      .FontBold = True
      .IndentLeft = 300
      .SpaceBefore = 200
      .SpaceAfter = 150
      .TextAlign = taLeftTop
      .Paragraph = CStr(lNumber) & "." & iCntPara & "." & " Модифицированные факторы"
      iCntPara = iCntPara + 1
      .SpaceBefore = 0
      .SpaceAfter = 0
      .IndentLeft = 0
      
      Fmt = "^+900|^+2100|^+2100|^+900|^+1500|^+1500;"
          
      s = "Фактор|Оценка до|Оценка после|Изменение|Уверенность до|Уверенность после;"
       
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
              
      Fmt = ">+900|<+2100|<+2100|>+900|>+1500|>+1500;"
      .FontItalic = False
      .TablePenTB = 1
      .FontBold = False
      .LineSpacing = 100
      .SpaceBefore = 50
      .SpaceAfter = 50
      
      For k = LBound(m_arrFacCh) To UBound(m_arrFacCh)
        With m_arrFacCh(k)
          
          sBody = .sShortName & "|" & .sValOld & "|" _
            & .sVal & "|" & CStr(.iDelta) & "|" _
            & .sLvlOld & "|" & .sLvl & ";"
            
          If k = UBound(m_arrFacCh) Then vp.TablePenTB = 20
          vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          
        End With
      Next k
    End If
    
    If IsValidTSD(m_arrOld) Then
      .SpaceBefore = 0
      .SpaceAfter = 0
      .Paragraph = ""
      
      .FontItalic = False
      .FontSize = 12
      .FontName = "Arial"
      .FontBold = True
      .IndentLeft = 300
      .SpaceBefore = 200
      .SpaceAfter = 150
      .TextAlign = taLeftTop
      .Paragraph = CStr(lNumber) & "." & iCntPara & "." & " Изменение выходных параметров модели"
      iCntPara = iCntPara + 1
      .SpaceBefore = 0
      .SpaceAfter = 0
      .IndentLeft = 0
      
      Fmt = "^+1600|^+2100|^+2100|^+1100|^+1100|^+1100;"
          
      s = "Состояние|Частота|Среднее кв. отклонение|Min(индекса опасности)|Max(индекса опасности)|Средний индекс опасности;"
       
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
              
      Fmt = ">+1600|<+2100|<+2100|>+1100|>+1100|>+1100;"
      .FontItalic = False
      .TablePenTB = 1
      .FontBold = False
      .LineSpacing = 100
      .SpaceBefore = 50
      .SpaceAfter = 50
      
      For k = LBound(m_arr) To UBound(m_arr)
        With m_arr(k)
          
          sBody = .sName & "|" & Format(.dMx - m_arrOld(k).dMx, haApp.FmtDbl) & "|" _
          & Format(.dDx - m_arrOld(k).dDx, haApp.FmtDbl) & "|" & Format(.dIdxMin - m_arrOld(k).dIdxMin, haApp.FmtDbl2) & "|" _
          & Format(.dIdxMax - m_arrOld(k).dIdxMax, haApp.FmtDbl2) & "|" & Format(.dIdxMx - m_arrOld(k).dIdxMx, haApp.FmtDbl2) & ";"
          
          If k = UBound(m_arr) Then vp.TablePenTB = 20
          vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          
        End With
      Next k
    End If
            
    .SpaceBefore = 0
    .SpaceAfter = 0
    .Paragraph = ""
  End With
End Sub

Public Property Get IRepItem_Title() As String
  IRepItem_Title = "Прогон модели"
End Property

Public Sub InitReport(ByVal mgn As MGertNet, ByVal ha As HazardApp, ByVal mgnOld As MGertNet)
  m_s_FileName = ha.NameDoc
  m_RunMode = mgn.RunMode
  With mgn
    m_l_N = .N
    m_l_K = .k
    m_mt_Alho = .RunMode
    m_dt_Time = .TimeWork
    If Not mgnOld Is Nothing Then _
      m_dt_TimeOld = mgnOld.TimeWork
  End With
  
  'ReDim m_arrFacCh( ) As TFactorsData
  
  LoadFrom mgn, m_arr
  If Not mgnOld Is Nothing Then
    LoadFrom mgnOld, m_arrOld
        
    Dim cFac As Factor, cFacOld As Factor, ibsKey As IBSTRKey
    Dim lCnt As Long
    lCnt = 0
    For Each cFac In mgn.Factors
      Set ibsKey = cFac
      Set cFacOld = mgnOld.Factors(ibsKey.Get())
      If cFac.Value <> cFacOld.Value Or cFac.TrustLvl <> cFacOld.TrustLvl Then _
        lCnt = lCnt + 1
    Next cFac
    
    If lCnt > 0 Then _
      ReDim m_arrFacCh(lCnt - 1) As TFactorsData
    
    lCnt = 0
    For Each cFac In mgn.Factors
      Set ibsKey = cFac
      Set cFacOld = mgnOld.Factors(ibsKey.Get())
      If cFac.Value <> cFacOld.Value Or cFac.TrustLvl <> cFacOld.TrustLvl Then
        With m_arrFacCh(lCnt)
          mgn.GetFactorPresent cFac, .sVal, .sLvl, .dVal
          mgnOld.GetFactorPresent cFacOld, .sValOld, .sLvlOld, .dValOld
          .iDelta = cFac.Value - cFacOld.Value
          .sName = cFac.Name
          .sShortName = ibsKey.Get()
          .trLvl = cFac.TrustLvl
          .trLvlOld = cFacOld.TrustLvl
        End With
        lCnt = lCnt + 1
      End If
    Next cFac
    
  Else
    'ReDim m_arrOld(-1) As TStateDta
  End If
  
End Sub

Private Sub LoadFrom(ByVal mgn As MGertNet, ByRef arr() As TStateDta)
  Dim l As Long
  Dim sarr()
  sarr = Array("Гомеостазис", "Нарушение ==", "Адаптация D", _
    "Оп. ситуация", "Адаптация F", "Кр. ситуация", _
    "Адаптация H", "Происшествие")
          
  ReDim arr(7) As TStateDta
              
  For l = 66 To 73
    With arr(l - 66)
      mgn.GetInfSampleK l, .lMin, .lMax, .dMx, .dDx
      .sName = sarr(l - 66)
      
      'If Not (mgn.RunMode = MT_Analytical Or mgn.RunMode = MT_AnalyticalIndistinct) Then
      mgn.GetInfSampleKIdx l, .dIdxMin, .dIdxMax, .dIdxMx, .dIdxDx
      
      'If mgn.RunMode = MT_Analytical Or mgn.RunMode = MT_AnalyticalIndistinct Then
        '.dDx = .dIdxDx
    End With
  Next l
End Sub
