VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRepModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IRepItem

Private m_s_FileName As String
Private m_l_1 As Long
Private m_l_2 As Long
Private m_l_3 As Long
Private m_l_4 As Long

Private m_d_1 As Double
Private m_d_2 As Double
Private m_d_3 As Double
Private m_d_4 As Double

Private Const m_fmtI = "#0.0###;;0"

Private Type FData
  sShortName As String
  sName As String
  dVal As Double
  sVal As String
  tlLevel As GERTNETLib.TrustLevel
  sLevel As String
  arrIdx() As Integer
End Type

Private fdarr() As FData

Public Sub IRepItem_GenerateReport(ByVal vp0 As Object, ByVal lNumber As Long)
  Dim vp As VSPrinter
  Set vp = vp0
  
  With vp
    .LineSpacing = 100
        
    .FontSize = 10
    .FontName = "Courier New"
    .Font.Charset = 204
    .FontBold = False
    .FontItalic = False
    .TextAlign = taLeftTop
    .Paragraph = "Модель: '" & m_s_FileName & "'"
    
    .FontSize = 12
    .FontName = "Arial"
    .FontBold = True
    .IndentLeft = 300
    .SpaceBefore = 200
    .SpaceAfter = 150
    .Paragraph = CStr(lNumber) & ".1." & " Факторы опасности"
    .SpaceBefore = 0
    .SpaceAfter = 0
    .IndentLeft = 0
                
    '.FontName = "Arial"
    .FontItalic = True
    .FontSize = 10
    .FontBold = True
    
    Dim Fmt As String, s As String
    Fmt = "^+1500|^+2000|^+900|^+2000|^+1200;"                   '^Center > Right
          
    s = "Фактор опасности|Лингвистическая оценка|Значение оценки|Уверенность|Индексы;"
     
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
        
    Fmt = "+1500|>+2000|>+900|>+2000|>+1200;"
    .FontItalic = False
    
    Dim k As Long
    Dim arrLit: arrLit = Array("H", "M", "C", "T")
    Dim arrLit2: arrLit2 = Array("Человек;", "Машина;", "Среда;", "Технология;")
    For k = LBound(arrLit) To UBound(arrLit)
      .FontBold = True
      .FontItalic = True
      .LineSpacing = 150
      .TablePenTB = 20
      .SpaceBefore = 60
      vp.AddTable "^+7600;", arrLit2(k), "", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), False
      .SpaceBefore = 0
      .FontBold = False
      .FontItalic = False
      .LineSpacing = 100
      .TablePenTB = 1
      Dim l As Long
      For l = LBound(fdarr) To UBound(fdarr)
        Dim sBody As String
        With fdarr(l)
          If Left(.sShortName, 1) = arrLit(k) Then
            Dim iIdxC As Integer, sIdx As String
            sIdx = ""
            For iIdxC = 0 To UBound(.arrIdx)
              sIdx = sIdx & CStr(.arrIdx(iIdxC))
              If iIdxC < UBound(.arrIdx) Then _
                sIdx = sIdx & "/"
            Next iIdxC
            sBody = .sShortName & "|" & .sVal & "|" & CStr(.dVal) & "|" & .sLevel & "|" & sIdx & ";"
            If k = UBound(arrLit) And l = UBound(fdarr) Then vp.TablePenTB = 20
            vp.AddTable Fmt, "", sBody, RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
          End If
        End With
      Next l
    Next k
    
    .SpaceBefore = 0
    .SpaceAfter = 0
    .Paragraph = ""
    '.PreviewPage = 1
          
    .SpaceAfter = 150
    .FontSize = 12
    .FontName = "Arial"
    .FontBold = True
    .FontItalic = False
    .IndentLeft = 300
    .TextAlign = taLeftTop
    '.TextAlign = taCenterMiddle
    .Paragraph = CStr(lNumber) & ".2." & " Краны"
    .IndentLeft = 0
    .SpaceBefore = 0
    .SpaceAfter = 0
    
    .FontItalic = True
    .FontSize = 10
    .FontBold = True
        
    Fmt = "^+900|^+900|^+900|^+900|^+900|^+900|^+900|^+900;"
          
    s = "I|II|III|IV|Iуст.|IIуст.|IIIуст.|IVуст.;"
     
    .PenWidth = 1
    .TextAlign = taCenterTop
  
    .FontSize = 10
    .TableBorder = 7
    .TablePen = 1
    .TablePenTB = 20
    .TablePenLR = 20
        
    .LineSpacing = 100
    .SpaceBefore = 0
    
    vp.AddTable Fmt, s, "", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), False
    
    .FontBold = False
    .FontItalic = False
    .LineSpacing = 100
    Fmt = "^+900|^+900|^+900|^+900|^+900|^+900|^+900|^+900;"
    vp.AddTable Fmt, "", CStr(m_l_1) & "|" & CStr(m_l_2) & "|" & CStr(m_l_3) & "|" & CStr(m_l_4) & _
      "|" & Format(m_d_1, m_fmtI) & "|" & Format(m_d_2, m_fmtI) & "|" & Format(m_d_3, m_fmtI) & "|" & Format(m_d_4, m_fmtI) & ";", RGB(&HFF, &HFF, &HFF), RGB(&HFF, &HFF, &HFF), True
    
    .Paragraph = ""
    .SpaceBefore = 0
    .SpaceAfter = 0
  End With
End Sub

Public Property Get IRepItem_Title() As String
  IRepItem_Title = "Модель"
End Property


Public Sub InitReport(ByVal mgn As MGertNet, ByVal ha As HazardApp)
  m_s_FileName = ha.NameDoc
  ReDim fdarr(mgn.Factors.Count - 1) As FData
  Dim f As Factor, l As Long
  
  With mgn
    m_l_1 = .VParam(1)
    m_l_2 = .VParam(2)
    m_l_3 = .VParam(3)
    m_l_4 = .VParam(4)
    
    m_d_1 = .VParamIndistinct(1)
    m_d_2 = .VParamIndistinct(2)
    m_d_3 = .VParamIndistinct(3)
    m_d_4 = .VParamIndistinct(4)
  End With
  
  l = 0
  For Each f In mgn.Factors
    With fdarr(l)
      .sName = f.Name
      .tlLevel = f.TrustLvl
      ReDim .arrIdx(0 To f.NIdx - 1) As Integer
      
      Dim i As Integer
      For i = 0 To UBound(.arrIdx)
        .arrIdx(i) = f.Idx(i)
      Next i
      
      Dim ibs As IBSTRKey
      Set ibs = f
      
      .sShortName = ibs.Get
            
      mgn.GetFactorPresent f, .sVal, .sLevel, .dVal
      l = l + 1
    End With
  Next f
End Sub
