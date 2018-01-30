Attribute VB_Name = "InitDef"
Option Explicit

Public Sub InitDefaultMG(ByVal mag As MGertNet)
  With mag
    .VParam(1) = 56
    .VParam(2) = 350
    .VParam(3) = 2112
    .VParam(4) = 3240
    
    .VParamIndistinct(1) = 50.7338436903943
    .VParamIndistinct(2) = 253.95384929772
    .VParamIndistinct(3) = 1318.28002747634
    .VParamIndistinct(4) = 2165.0814569974

    .VProbability(1) = 2.0756699141479E-03
    .VProbability(2) = 2.290085356652E-04
    .VProbability(3) = 2.400032467948E-05
    .VProbability(4) = 3.0634111572526E-06
    
    .k = 200: .N = 40000: .RndBase = 1
  End With

  AddLingvs mag.Enumerators
  AddFactors mag.Factors
End Sub

Public Sub InitDefaultSF(ByVal coll As CollectionX)
  Dim tCollSFT As New collSF, iStg As IPersistStorage
  Set iStg = tCollSFT
  iStg.InitNew Nothing
  CreCollSF tCollSFT
  
  coll.Add tCollSFT, "����������� �������� ���"
End Sub

Private Sub CreCollSF(ByVal coll As collSF)
  Dim sf As New SafetyPrecaution
  Dim fc As FChange
  With sf
    .Cost = 100
    .Name = "���� 100"
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
    .Name = "���� 500"
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
    .Name = "���� 90"
    Set fc = New FChange
    fc.Delta = 1: fc.NameFactor = "M04"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 80
    .Name = "���� 80"
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "H02"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
  Set sf = New SafetyPrecaution
  With sf
    .Cost = 197
    .Name = "���� 197"
    Set fc = New FChange
    fc.Delta = 3: fc.NameFactor = "H08"
    .FChanges.Add fc
    
    Set fc = New FChange
    fc.Delta = 2: fc.NameFactor = "M07"
    .FChanges.Add fc
    
    coll.Add sf
  End With
  
End Sub


Private Sub AddFactors(ByVal cll As collFac)
  Dim f As New Factor
  f.Name = "����������� �� ��������������� �����������"
  f.Value = 5
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H01"
  
  Set f = New Factor
  f.Name = "��������������� ��������������������"
  f.Value = 5
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 2
  cll.Add f, "H02"
  
  Set f = New Factor
  f.Name = "�������� ����� � ������������� ����������"
  f.Value = 7
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1: f.Idx(2) = 2: f.Idx(3) = 3
  cll.Add f, "H03"
  
  Set f = New Factor
  f.Name = "����� ���������� ������"
  f.Value = 6
  f.IDEnum = 3
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H04"
  
  Set f = New Factor
  f.Name = "�������� ������������� ���������"
  f.Value = 4
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H05"
  
  Set f = New Factor
  f.Name = "������ ���������� �����"
  f.Value = 7
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H06"
  
  Set f = New Factor
  f.Name = "������ ���������� �������� ��������� � �������"
  f.Value = 5
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H07"
  
  Set f = New Factor
  f.Name = "����������� ��������� ��������� ����������"
  f.Value = 6
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H08"
  
  Set f = New Factor
  f.Name = "�������� �������� ������� (������������� ��������)"
  f.Value = 6
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1: f.Idx(2) = 2
  cll.Add f, "H09"
  
  Set f = New Factor
  f.Name = "������������� � ������������� ���������"
  f.Value = 2
  f.IDEnum = 2
  f.Idx(0) = 1: f.Idx(1) = 2
  cll.Add f, "H12"
  
  Set f = New Factor
  f.Name = "����������� ��������� � ������� �������� ����������"
  f.Value = 4
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "H13"
  
  Set f = New Factor
  f.Name = "�������� �������������� ��������"
  f.Value = 7
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1: f.Idx(2) = 2
  cll.Add f, "H14"
  
  '-------------------------------
  
  Set f = New Factor
  f.Name = "�������� ����������� �������� ����� ���������"
  f.Value = 4
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 2
  cll.Add f, "M01"
  
  Set f = New Factor
  f.Name = "������� ����� ������������ ����������������� ��������"
  f.Value = 4
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "M02"
  
  Set f = New Factor
  f.Name = "������������ ����������� ������� � ������� ��������"
  f.Value = 6
  f.IDEnum = 5
  f.Idx(0) = 2: f.Idx(1) = 0
  cll.Add f, "M03"
  
  Set f = New Factor
  f.Name = "������������� ������ ���������"
  f.Value = 7
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "M04"
  
  Set f = New Factor
  f.Name = "������������� ������ ������������� ���������"
  f.Value = 8
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 2
  cll.Add f, "M05"
  
  Set f = New Factor
  f.Name = "������������ �������� ������� � ������� ��������"
  f.Value = 3
  f.IDEnum = 1
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "M06"
  
  Set f = New Factor
  f.Name = "������� ���������� ������� � ������� ��������"
  f.Value = 4
  f.IDEnum = 5
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "M07"
  
  Set f = New Factor
  f.Name = "������������� �������� � ��������� ������������"
  f.Value = 9
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "M08"
  
  '-------------------------------
  Set f = New Factor
  f.Name = "�������� ���������� � ���������� �����"
  f.Value = 4
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "T01"
  
  Set f = New Factor
  f.Name = "�������� ������������ ������������ � �������"
  f.Value = 2
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 2
  cll.Add f, "T02"
  
  Set f = New Factor
  f.Name = "��������� ���������� ���������"
  f.Value = 6
  f.IDEnum = 5
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "T03"
  
  Set f = New Factor
  f.Name = "����������� ��������� �������� � ������� ����"
  f.Value = 4
  f.IDEnum = 5
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "T04"
  
  Set f = New Factor
  f.Name = "����������� ��������� ������ ������������ ��������� � ������� ����"
  f.Value = 4
  f.IDEnum = 5
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "T05"
  
  Set f = New Factor
  f.Name = "���������� ��������������� ������� ����������� ������������"
  f.Value = 8
  f.IDEnum = 5
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "T06"
  
  '-------------------------------
  Set f = New Factor
  f.Name = "������������ �� ������-���������� ���������� ��������� �����"
  f.Value = 4
  f.IDEnum = 1
  f.Idx(0) = 0: f.Idx(1) = 1
  cll.Add f, "C01"
  
  Set f = New Factor
  f.Name = "�������� �������������� ������ ��������� �����"
  f.Value = 6
  f.IDEnum = 2
  f.Idx(0) = 0: f.Idx(1) = 2
  cll.Add f, "C02"
  
  Set f = New Factor
  f.Name = "����������� ������� ������� �����������"
  f.Value = 1
  f.IDEnum = 5
  f.Idx(0) = 2: f.Idx(1) = 0
  cll.Add f, "C03"
  
  Set f = New Factor
  f.Name = "����������� ������� ��������������� �����������"
  f.Value = 2
  f.IDEnum = 5
  f.Idx(0) = 1: f.Idx(1) = 0
  cll.Add f, "C04"
  
End Sub

Private Sub AddLingvs(ByVal cll As CollLingvo)
  Dim le As New LingvoEnum '�� = 1
  le.Mark(0) = "�����, ����� ������"
  le.Mark(1) = "����� ������"
  le.Mark(2) = "������"
  le.Mark(3) = "���� ��������"
  le.Mark(4) = "�������"
  le.Mark(5) = "���� ��������"
  le.Mark(6) = "�������"
  le.Mark(7) = "����� �������"
  le.Mark(8) = "�������"
  le.Mark(9) = "����� �������"
  le.Mark(10) = "�����, ����� �������"
  
  cll.Add le
  
  Set le = New LingvoEnum '�� = 2
  le.Mark(0) = "�����, ����� ������"
  le.Mark(1) = "����� ������"
  le.Mark(2) = "������"
  le.Mark(3) = "���� ��������"
  le.Mark(4) = "�������"
  le.Mark(5) = "���� ��������"
  le.Mark(6) = "�������"
  le.Mark(7) = "����� �������"
  le.Mark(8) = "�������"
  le.Mark(9) = "����� �������"
  le.Mark(10) = "�����, ����� �������"
  
  cll.Add le
  
  Set le = New LingvoEnum '�� = 3
  le.Mark(0) = "�����, ����� ������"
  le.Mark(1) = "����� ������"
  le.Mark(2) = "������"
  le.Mark(3) = "���� ��������"
  le.Mark(4) = "�������"
  le.Mark(5) = "���� ��������"
  le.Mark(6) = "�������"
  le.Mark(7) = "����� �������"
  le.Mark(8) = "��������"
  le.Mark(9) = "����� �������"
  le.Mark(10) = "�����, ����� �������"
  
  cll.Add le
  
  Set le = New LingvoEnum '�� = 4
  le.Mark(0) = "�����, ����� ������"
  le.Mark(1) = "����� ������"
  le.Mark(2) = "������"
  le.Mark(3) = "���� ��������"
  le.Mark(4) = "�������"
  le.Mark(5) = "���� ��������"
  le.Mark(6) = "�������"
  le.Mark(7) = "����� �������"
  le.Mark(8) = "��������"
  le.Mark(9) = "����� �������"
  le.Mark(10) = "�����, ����� �������"
  
  cll.Add le
  
  Set le = New LingvoEnum '������������ = 5
  le.Mark(0) = "�����, ����� ������"
  le.Mark(1) = "����� ������"
  le.Mark(2) = "������"
  le.Mark(3) = "���� ��������"
  le.Mark(4) = "�������"
  le.Mark(5) = "���� ��������"
  le.Mark(6) = "�������"
  le.Mark(7) = "����� �������"
  le.Mark(8) = "�������"
  le.Mark(9) = "����� �������"
  le.Mark(10) = "�����, ����� �������"
  
  cll.Add le
    
End Sub

