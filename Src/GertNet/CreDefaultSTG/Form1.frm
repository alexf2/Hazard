VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "graph32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCalc 
      Caption         =   "Calc"
      Height          =   465
      Left            =   5235
      TabIndex        =   6
      Top             =   300
      Width           =   1140
   End
   Begin VB.TextBox txtDir 
      Height          =   360
      Left            =   3300
      TabIndex        =   5
      Top             =   3870
      Width           =   3075
   End
   Begin VB.CommandButton cmExtractAll 
      Caption         =   "Extract To:"
      Height          =   465
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   1140
   End
   Begin VB.CommandButton cmRefresh 
      Caption         =   "Refresh"
      Height          =   465
      Left            =   5250
      TabIndex        =   3
      Top             =   1575
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   5205
      TabIndex        =   2
      Text            =   "C01"
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   465
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2265
      Width           =   1140
   End
   Begin GraphLib.Graph Graph1 
      Height          =   3525
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   4950
      _Version        =   65536
      _ExtentX        =   8731
      _ExtentY        =   6218
      _StockProps     =   96
      BorderStyle     =   1
      AutoInc         =   0
      DrawMode        =   3
      GraphStyle      =   5
      GraphType       =   6
      GridStyle       =   3
      RandomData      =   0
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   1
      GraphData[]     =   5
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sStgName As String = "g:\default_stg.atg"
Private magG As MGertNet

Private Sub Extract(ByVal sFactName As String, ByVal bToFile As Boolean)
  Dim mgDbg As IMGertNet_Debug
  Set mgDbg = magG
  Dim arrdPts() As Double
  arrdPts = mgDbg.GetPtsForFactor(sFactName)
  Dim l As Long, lSz As Long, lPts As Long
  lSz = UBound(arrdPts) - LBound(arrdPts) + 1
  lPts = lSz / 2
  
  With Graph1
    .DrawMode = GraphLib.DrawModeConstants.gphNoAction
    
    .AutoInc = AutoIncConstants.gphOff
    .RandomData = RandomDataConstants.gphOff
    .DataReset = gphGraphData
    .NumSets = 1
    .ThisSet = 1
    .NumPoints = lPts
    '.TickEvery = 2
    
    '.Background = gphBlue
    
    .ThisSet = 1
    
    l = 0
    .ThisPoint = 1
    Dim lPtCnt As Long
    lPtCnt = 1
    Do While l < lSz
          
      .ThisPoint = lPtCnt
      lPtCnt = lPtCnt + 1
      .XPosData = arrdPts(l)
      l = l + 1
      .GraphData = arrdPts(l)
      l = l + 1
        
    Loop
            
    .PrintStyle = PrintStyleConstants.gphColor
    .GridStyle = GridStyleConstants.gphBoth
    
    Dim sDescr As String, sTrust As String, dValue As Double
    Dim fFact As Factor
    Set fFact = magG.Factors(sFactName)
    magG.GetFactorPresent fFact, sDescr, sTrust, dValue
    
    If magG.RunMode <> MT_ImitateIndistinct Then sTrust = "*"
    
    Dim ss As String
    .BottomTitle = sFactName & "(" & CStr(lPts) & " pts) <" & sTrust & "> = " & FormatNumber(dValue) & "(" & sDescr & ")"
    .GraphTitle = fFact.Name
    '.Labels = LabelsConstants.gphOn
    
    .FontUse = gphOtherTitles
    .FontStyle = GraphLib.FontStyleConstants.gphItalic
    .DrawMode = GraphLib.DrawModeConstants.gphBlit
            
    If bToFile Then
      .ImageFile = txtDir + sFactName + ".bmp"
      .DrawMode = GraphLib.DrawModeConstants.gphWrite
    End If
    
  End With
End Sub

Private Sub cmCalc_Click()
  Dim mgDbg As IMGertNet_Debug
  Set mgDbg = magG
  
  With Graph1
    .DrawMode = GraphLib.DrawModeConstants.gphNoAction
    .AutoInc = AutoIncConstants.gphOff
    .RandomData = RandomDataConstants.gphOff
    .DataReset = gphGraphData
    .NumSets = 1
    .ThisSet = 1
    .NumPoints = CLng(1 / 0.01)
        
    Dim dbl As Double, i As Integer
    i = 1
    For dbl = 0 To 1 Step 0.01
      .ThisPoint = i
      i = i + 1
      .XPosData = dbl
      .GraphData = mgDbg.TestFunc(Text1.Text, dbl)
    Next dbl
    .DrawMode = GraphLib.DrawModeConstants.gphBlit
  End With
End Sub

Private Sub cmExtractAll_Click()
  'On Error GoTo ERR_H
  
  Dim arrFacs As Variant
  arrFacs = Array( _
    "H01", "H02", "H03", "H04", "H05", "H06", "H07", "H08", "H09", "H12", "H13", "H14", _
    "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", _
    "T01", "T02", "T03", "T04", "T05", "T06", _
    "C01", "C02", "C03", "C04")
  
  Dim l As Long
  For l = LBound(arrFacs) To UBound(arrFacs) Step 1
    Extract arrFacs(l), True
  Next l
    
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub cmRefresh_Click()
  On Error GoTo ERR_H
  
  Extract Text1.Text, False
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Private Sub Form_Load()

  On Error GoTo ERRH

  Dim mag As New MGertNet
  Dim iStg As IPersistStorage
  Set iStg = mag
  iStg.InitNew Nothing

  mag.InTest1 = 59
  mag.InTest2 = 359
  mag.InTest3 = 2140
  mag.InTest4 = 3270

  AddLingvs mag.Enumerators
  AddFactors mag.Factors

  Dim stg As IStorage
  Set stg = CreateCF(sStgName, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)

  iStg.Save stg, 0
  
  Dim tCollSFT As New CollSF
  Set iStg = tCollSFT
  iStg.InitNew Nothing
  CreCollSF tCollSFT
  
  Dim stg2 As IStorage
  Set stg2 = CreateStorageInt(stg, "CollSF_Main", _
    STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    
      
  iStg.Save stg2, 0
  
  Set stg = Nothing
  Set iStg = Nothing
  Set mag = Nothing

  'MkTest
  
  Exit Sub
ERRH:
  MsgBox Err.Description, vbOKOnly Or vbExclamation, "Error"
  
End Sub

Private Sub AddFactors(ByVal cll As CollFac)
  Dim f As New Factor
  f.Name = "Пригодность по физиологическим показателям"
  f.Value = 5
  f.IDEnum = 1
  cll.Add f, "H01"
  
  Set f = New Factor
  f.Name = "Технологическая дисциплинированность"
  f.Value = 5
  f.IDEnum = 1
  cll.Add f, "H02"
  
  Set f = New Factor
  f.Name = "Качество приёма и декодирования информации"
  f.Value = 7
  f.IDEnum = 2
  cll.Add f, "H03"
  
  Set f = New Factor
  f.Name = "Навык выполнения работы"
  f.Value = 6
  f.IDEnum = 3
  cll.Add f, "H04"
  
  Set f = New Factor
  f.Name = "Качество мотивационной установки"
  f.Value = 4
  f.IDEnum = 2
  cll.Add f, "H05"
  
  Set f = New Factor
  f.Name = "Знание технологии работ"
  f.Value = 7
  f.IDEnum = 2
  cll.Add f, "H06"
  
  Set f = New Factor
  f.Name = "Знание физической сущности процессов в системе"
  f.Value = 5
  f.IDEnum = 2
  cll.Add f, "H07"
  
  Set f = New Factor
  f.Name = "Способность правильно оценивать информацию"
  f.Value = 6
  f.IDEnum = 1
  cll.Add f, "H08"
  
  Set f = New Factor
  f.Name = "Качество принятия решения (оперативность мышл.)"
  f.Value = 6
  f.IDEnum = 2
  cll.Add f, "H09"
  
  Set f = New Factor
  f.Name = "Самообладание в экстремальных ситуациях"
  f.Value = 2
  f.IDEnum = 2
  cll.Add f, "H12"
  
  Set f = New Factor
  f.Name = "Обученность действиям в сложных условиях обстановки"
  f.Value = 4
  f.IDEnum = 1
  cll.Add f, "H13"
  
  Set f = New Factor
  f.Name = "Точность корректирующих действий"
  f.Value = 7
  f.IDEnum = 1
  cll.Add f, "H14"
  
  '-------------------------------
  
  Set f = New Factor
  f.Name = "Качество конструкции рабочего места оператора"
  f.Value = 4
  f.IDEnum = 2
  cll.Add f, "M01"
  
  Set f = New Factor
  f.Name = "Степень учета особенностей работоспособности человека"
  f.Value = 4
  f.IDEnum = 1
  cll.Add f, "M02"
  
  Set f = New Factor
  f.Name = "Оснащенность источниками опасных и вредных факторов"
  f.Value = 6
  f.IDEnum = 5
  cll.Add f, "M03"
  
  Set f = New Factor
  f.Name = "Безотказность прочих элементов"
  f.Value = 7
  f.IDEnum = 1
  cll.Add f, "M04"
  
  Set f = New Factor
  f.Name = "Безотказность других ответственных элементов"
  f.Value = 8
  f.IDEnum = 1
  cll.Add f, "M05"
  
  Set f = New Factor
  f.Name = "Длительность действия опасных и вредных факторов"
  f.Value = 3
  f.IDEnum = 1
  cll.Add f, "M06"
  
  Set f = New Factor
  f.Name = "Уровень потенциала опасных и вредных факторов"
  f.Value = 4
  f.IDEnum = 5
  cll.Add f, "M07"
  
  Set f = New Factor
  f.Name = "Безотказность приборов и устройств безопасности"
  f.Value = 9
  f.IDEnum = 1
  cll.Add f, "M08"
  
  '-------------------------------
  Set f = New Factor
  f.Name = "Удобство подготовки и выполнения работ"
  f.Value = 4
  f.IDEnum = 2
  cll.Add f, "T01"
  
  Set f = New Factor
  f.Name = "Удобство технического обслуживания и ремонта"
  f.Value = 2
  f.IDEnum = 2
  cll.Add f, "T02"
  
  Set f = New Factor
  f.Name = "Сложность алгоритмов оператора"
  f.Value = 6
  f.IDEnum = 5
  cll.Add f, "T03"
  
  Set f = New Factor
  f.Name = "Возможность появления человека в опасной зоне"
  f.Value = 4
  f.IDEnum = 5
  cll.Add f, "T04"
  
  Set f = New Factor
  f.Name = "Возм-ть появления др. незащищенных эл-тов в опасной зоне"
  f.Value = 4
  f.IDEnum = 5
  cll.Add f, "T05"
  
  Set f = New Factor
  f.Name = "Надежность технологич. средств обеспечения безопасности"
  f.Value = 8
  f.IDEnum = 5
  cll.Add f, "T06"
  
  '-------------------------------
  Set f = New Factor
  f.Name = "Комфортность по физ.-хим. параметрам воздушной среды"
  f.Value = 4
  f.IDEnum = 1
  cll.Add f, "C01"
  
  Set f = New Factor
  f.Name = "Качество информационной модели о состоянии среды"
  f.Value = 6
  f.IDEnum = 2
  cll.Add f, "C02"
  
  Set f = New Factor
  f.Name = "Возможность внешних опасных воздействий"
  f.Value = 1
  f.IDEnum = 5
  cll.Add f, "C03"
  
  Set f = New Factor
  f.Name = "Возможность внешних неблагоприятных воздействий"
  f.Value = 2
  f.IDEnum = 5
  cll.Add f, "C04"
  
End Sub

Private Sub AddLingvs(ByVal cll As CollLingvo)
  Dim le As New LingvoEnum 'ая = 1
  le.Mark(0) = "Очень, очень низкая"
  le.Mark(1) = "Очень низкая"
  le.Mark(2) = "Низкая"
  le.Mark(3) = "Ниже среднего"
  le.Mark(4) = "Средняя"
  le.Mark(5) = "Выше среднего"
  le.Mark(6) = "Хорошая"
  le.Mark(7) = "Очень хорошая"
  le.Mark(8) = "Высокая"
  le.Mark(9) = "Очень высокая"
  le.Mark(10) = "Очень, очень высокая"
  
  cll.Add le
  
  Set le = New LingvoEnum 'ое = 2
  le.Mark(0) = "Очень, очень низкое"
  le.Mark(1) = "Очень низкое"
  le.Mark(2) = "Низкое"
  le.Mark(3) = "Ниже среднего"
  le.Mark(4) = "Среднее"
  le.Mark(5) = "Выше среднего"
  le.Mark(6) = "Хорошее"
  le.Mark(7) = "Очень хорошее"
  le.Mark(8) = "Высокое"
  le.Mark(9) = "Очень высокое"
  le.Mark(10) = "Очень, очень высокое"
  
  cll.Add le
  
  Set le = New LingvoEnum 'ие = 3
  le.Mark(0) = "Очень, очень низкие"
  le.Mark(1) = "Очень низкие"
  le.Mark(2) = "Низкие"
  le.Mark(3) = "Ниже среднего"
  le.Mark(4) = "Средние"
  le.Mark(5) = "Выше среднего"
  le.Mark(6) = "Хорошие"
  le.Mark(7) = "Очень хорошие"
  le.Mark(8) = "Высокоие"
  le.Mark(9) = "Очень высокие"
  le.Mark(10) = "Очень, очень высокие"
  
  cll.Add le
  
  Set le = New LingvoEnum 'ий = 4
  le.Mark(0) = "Очень, очень низкий"
  le.Mark(1) = "Очень низкий"
  le.Mark(2) = "Низкий"
  le.Mark(3) = "Ниже среднего"
  le.Mark(4) = "Средний"
  le.Mark(5) = "Выше среднего"
  le.Mark(6) = "Хороший"
  le.Mark(7) = "Очень хороший"
  le.Mark(8) = "Высокоий"
  le.Mark(9) = "Очень высокий"
  le.Mark(10) = "Очень, очень высокий"
  
  cll.Add le
  
  Set le = New LingvoEnum 'оснащённость = 5
  le.Mark(0) = "Очень, очень низкая"
  le.Mark(1) = "Очень низкая"
  le.Mark(2) = "Низкая"
  le.Mark(3) = "Ниже среднего"
  le.Mark(4) = "Средняя"
  le.Mark(5) = "Выше среднего"
  le.Mark(6) = "Большая"
  le.Mark(7) = "Очень большая"
  le.Mark(8) = "Высокая"
  le.Mark(9) = "Очень высокая"
  le.Mark(10) = "Очень, очень высокая"
  
  cll.Add le
    
End Sub

Private Sub MkTest()
  Set magG = New MGertNet
  Dim iStg As IPersistStorage
  Set iStg = magG
  
  Dim stg As IStorage
  Set stg = OpenCF(sStgName, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  
  iStg.Load stg
  
  'magG.RunMode = MT_Imitate
  magG.RunMode = MT_ImitateIndistinct
  magG.Run 5, 10, False, -1, True
    
End Sub

Private Sub CreCollSF(ByVal coll As CollSF)
  Dim sf As New SafetyPrecaution
  Dim fc As FChange
  With sf
    .Cost = 100
    .Name = "Мера 100"
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "H12"
    .FChanges.Add fc
    
    Set fc = New FChange
    fc.Delta = 1: fc.NameFactor = "M03"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 500
    .Name = "Мера 500"
    Set fc = New FChange
    fc.Delta = 3: fc.NameFactor = "H14"
    .FChanges.Add fc
    
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "C01"
    .FChanges.Add fc
    
    Set fc = New FChange
    fc.Delta = 4: fc.NameFactor = "T02"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 90
    .Name = "Мера 90"
    Set fc = New FChange
    fc.Delta = 1: fc.NameFactor = "M04"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 80
    .Name = "Мера 80"
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "H02"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 197
    .Name = "Мера 197"
    Set fc = New FChange
    fc.Delta = 3: fc.NameFactor = "H08"
    .FChanges.Add fc
    
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "M07"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
End Sub

