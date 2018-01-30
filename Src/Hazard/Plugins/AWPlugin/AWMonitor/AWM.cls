VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GERTNETLib.IFactorAssign

Private m_lMRes As Long
Public PathToGostsStg As String
Private m_bIsMainStgGOSTS As Boolean

Friend Property Get IsMainStgGOSTS() As Boolean
  IsMainStgGOSTS = m_bIsMainStgGOSTS
End Property

Private Sub Class_Initialize()
  PathToGostsStg = GetDefaultGosts
  m_bIsMainStgGOSTS = True
End Sub

Private Function GetDefaultGosts() As String
  Dim sDir As String
  CutPath App.Path, sDir
  GetDefaultGosts = sDir & MainGOSTStgName
End Function

Public Property Get IFactorAssign_ModalResult() As Long
  IFactorAssign_ModalResult = m_lMRes
End Property


Public Sub IFactorAssign_About(ByVal frm As Object, ByVal lHwnd As Long)
  frmAbout.Show vbModal, frm
End Sub

Public Sub IFactorAssign_AssignFactor(ByVal frm As Object, ByVal lHwnd As Long, ByVal mg As MGertNet, ByVal f As Factor, ByVal sCfgName As String, ByVal sPath As String)

  m_lMRes = 0
  On Error GoTo ERR_H
  
  Dim bAss As Boolean: bAss = False
  
  If Not FetchStgGostsName(sPath) Then
    'm_lMRes = 0
    Exit Sub
  End If
  
  With frmAnswer
    If .FirstStart = 0 Then CenterForm frmAnswer, frm
    .AssData mg, f, sPath, PathToGostsStg, sCfgName
    bAss = True
    .Show vbModal, frm
    m_lMRes = IIf(.ModalResult, 1, 0)
    DoEvents
    
    If .ModalResult Then
      Dim le As LingvoEnum
      Set le = mg.GetEnumForFactor(f)
            
      f.Value = le.ValueIdx(le.RoundS(.ValueOfFactor / 10!))
    End If
      
    .UnassData
  End With
  
  Exit Sub
  
ERR_H:
  m_lMRes = 0
  If bAss Then frmAnswer.UnassData
  'MsgBox Err.Description, vbOKOnly, "TT"
  Err.Raise Err.Number
End Sub

Public Property Get IFactorAssign_Description() As String
  IFactorAssign_Description = "Универсальный модуль оценки, использующий метод средневзвешенного"
End Property


Public Sub IFactorAssign_Configure(ByVal frm As Object, ByVal lHwnd As Long, ByVal sConfigName As String, ByVal sFileName As String)
  'MsgBox "Конфигурация '" & sConfigName & "'. Файл: " & sFileName, vbOKOnly, "GasPipeLine"
  Dim bFlAss As Boolean: bFlAss = False
  
  If Not FetchStgGostsName(sFileName) Then
    m_lMRes = 0
    Exit Sub
  End If
  
  On Error Resume Next
  FileLen GetDefaultGosts
  If Err.Number = 53 Then
    Dim iStg As IStorage
    On Error GoTo ERR_H
    Set iStg = CreateCF(GetDefaultGosts, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    Set iStg = Nothing
  Else
    On Error GoTo ERR_H
    FileLen GetDefaultGosts
  End If
  
  On Error GoTo ERR_H
    With frmFactors
      If .FirstStart = 0 Then CenterForm frmFactors, frm
      .Caption = sConfigName
      .AssData sFileName, PathToGostsStg, IsMainStgGOSTS, GetDefaultGosts
      bFlAss = True
      .Show vbModal, frm
      DoEvents
      .UnassData
    End With
  Exit Sub
ERR_H:
  If bFlAss Then frmFactors.UnassData
  Err.Raise Err.Number
  
'  On Error GoTo ERR_H
'
'  With frmGostEdit
'    .KeyTopic = 0
'    .KeyGOST = 1
'    .PathToGostsStg = sFileName
'    .AssData
'    .Show vbModal, frm
'    DoEvents
'    .UnassData
'  End With
'
'  Exit Sub
'ERR_H:
'  frmGostEdit.UnassData
'  Err.Raise Err.Number
End Sub

Friend Function FetchStgGostsName(ByVal sFileName As String) As Boolean

  FetchStgGostsName = True
  
  Dim iStg As IStorage
  Dim sTmp As String
  Set iStg = OpenCF(sFileName, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  On Error Resume Next
  Dim iStrm As IStream
  Set iStrm = OpenStreamInt(iStg, StreamStgGostsName, STGM_DIRECT1 Or STGM_READ1 Or STGM_SHARE_EXCLUSIVE1)
  If Err.Number = 0 Then
    On Error GoTo 0
    sTmp = GetString(iStrm)
    On Error Resume Next
    FileLen sTmp
    If Err.Number = 0 Then
      PathToGostsStg = sTmp
      m_bIsMainStgGOSTS = False
    Else
      Dim vbRes As VbMsgBoxResult
      vbRes = MsgBox("Файл хранилища нормативов '" & sTmp & "' не найден. Используется стандартный: '" & PathToGostsStg & "'" & vbCrLf & "Продолжить ?", vbExclamation Or vbYesNo, "Ошибка")
      If vbRes = vbNo Then FetchStgGostsName = False
    End If
  End If
  Err.Clear
End Function

