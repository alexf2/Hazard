VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPluginEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ProgId As String
Public Description As String
Public Dirty As AdvDirty
Public Configs As CollectionX
Public Mounted As Boolean
Public OldProgId As Variant

Public Property Get Dirty2() As Boolean
  Dim cpce As CPlgCfgEntry
  For Each cpce In Configs
    If cpce.Dirty <> AD_None Then
      Dirty2 = True
      Exit Property
    End If
  Next
  Dirty2 = False
End Property

Public Sub Modify()
  Select Case Dirty
    Case Ad_Updated
    Case AD_None
      Dirty = Ad_Updated
    Case Ad_Deleted
    Case Ad_New
  End Select
End Sub

Public Function CheckMounted() As Long
  Dim pce As CPlgCfgEntry
  CheckMounted = 0
  
  For Each pce In Configs
    With pce
      On Error Resume Next
      Err.Clear
      FileLen .FileName
      If Err.Number = 0 Then
        .Mounted = True
      Else
        .Mounted = False
        CheckMounted = CheckMounted + 1
      End If
      On Error GoTo 0
    End With
  Next pce
End Function

Private Sub Class_Initialize()
  Dirty = AD_None
  Mounted = False
  Set Configs = New CollectionX
  Configs.DefaultItem = Nothing
  Configs.AutoDefaultGetKey = True
  OldProgId = Null
End Sub

Public Sub Clear()
  If Configs.Count > 0 Then _
    Configs.Remove Start:=1, Count:=Configs.Count
End Sub

Public Function AddConfig(ByVal sName As String, ByVal sFileName As String) As CPlgCfgEntry
  Dim pce As New CPlgCfgEntry, pceTmp As CPlgCfgEntry
  With pce
    .Name = sName: .FileName = sFileName: .Dirty = Ad_New
  End With
  
  With Configs
    Set pceTmp = .Item(sName)
    If Not pceTmp Is .DefaultItem Then
      If pceTmp.Dirty <> Ad_Deleted Then _
        Err.Raise vbObjectError + 3, "CPluginEntry.AddConfig", "Нельзя добавить: конфигурация '" & sName & "' уже имеется"
      .Remove sName
    End If
    
    .Add pce, sName
  End With
  
  Set AddConfig = pce
  Set pce = Nothing
End Function

Public Sub LoadFromRegistry(ByVal hKey As Long)
  Dim lNSubKeys As Long, lMaxlen As Long, lRes As Long
  
  Clear
  
  lRes = RegQueryInfoKey(hKey, vbNullString, 0&, 0&, VarPtr(lNSubKeys), VarPtr(lMaxlen), 0&, 0&, 0&, 0&, 0&, 0&)
  If lRes = ERROR_SUCCESS Then
    Dim sTmp As String, l As Long, lLenIn As Long
        
    For l = lNSubKeys - 1 To 0 Step -1
      sTmp = String$(lMaxlen + 1, vbNullChar)
      lLenIn = lMaxlen + 1
      lRes = RegEnumKeyEx(hKey, l, sTmp, VarPtr(lLenIn), 0&, vbNullString, 0&, 0&)
      If lRes <> ERROR_SUCCESS Then
        Err.Raise vbObjectError + 1, "CPluginEntry.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
      Else
        If lLenIn > 0 Then
          If Chr$(0) = Mid$(sTmp, lLenIn, 1) Then lLenIn = lLenIn - 1
        End If
        sTmp = Left$(sTmp, lLenIn)
        
        Dim hKey2 As Long
        lRes = RegOpenKeyEx(hKey, sTmp, 0, KEY_ALL_ACCESS, hKey2)
        If lRes <> ERROR_SUCCESS Then
          Err.Raise vbObjectError + 2, "CPluginEntry.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
        Else
          Dim sTmp2 As String, lPath As Long: lPath = MAX_PATH * 2
          sTmp2 = String$(MAX_PATH * 2, vbNullChar)
          lRes = RegQueryValueEx(hKey2, "", 0&, 0&, sTmp2, lPath)
          
          If lPath > 0 Then
            If Chr$(0) = Mid$(sTmp2, lPath, 1) Then lPath = lPath - 1
          End If
          sTmp2 = Left$(sTmp2, lPath)
          
          RegCloseKey hKey2
          If lRes <> ERROR_SUCCESS Then _
            Err.Raise vbObjectError + 2, "CPluginEntry.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
            
          Dim pce As CPlgCfgEntry
          Set pce = New CPlgCfgEntry
          With pce
            .Name = sTmp
            .OldName = sTmp
            .FileName = sTmp2
            .Dirty = AD_None
          End With
          Configs.Add pce, sTmp
          Set pce = Nothing
        End If
      End If
    Next l
  Else
    Err.Raise vbObjectError + 2, "CPluginEntry.LoadFromRegistry", "Ошибка чтения из реестра: '" & Win32ErrInfo(lRes) & "'"
  End If
End Sub

Public Sub ResolveChanges(ByVal hKey As Long)
  Dim ple As CPlgCfgEntry, lRes As Long, hKey2 As Long
  Dim collRemov As New CollectionX
  
  On Error GoTo ERR_H
  
  Dim cllTmp As CollectionX
  Set cllTmp = GetResolveReady(Configs)
  
  For Each ple In cllTmp
    With ple
      If .Dirty <> AD_None Then
              
        Select Case .Dirty
                  
          Case Ad_Deleted
            If Not IsNull(.OldName) Then _
              If ExitsKey(hKey, .OldName) Then MkRemoveKey hKey, .OldName
              
            collRemov.Add .Name
            
          Case Ad_New, Ad_Updated
            If .Dirty = Ad_New Then _
              If ExitsKey(hKey, .Name) Then MkRemoveKey hKey, .Name
              
            lRes = RegCreateKeyEx(hKey, .Name, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey2, 0&)
            If lRes <> ERROR_SUCCESS Then _
              Err.Raise vbObjectError + 2, "CPluginEntry.ResolveChanges", "Ошибка создания ключа '" & .Name & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
              
            lRes = RegSetValueEx2(hKey2, "", 0&, REG_SZ, .FileName, Len(.FileName))
            RegCloseKey hKey2
            If lRes <> ERROR_SUCCESS Then _
              Err.Raise vbObjectError + 2, "CPluginEntry.ResolveChanges", "Ошибка записи ключа '" & .Name & "' в реестре: '" & Win32ErrInfo(lRes) & "'"
              
            .OldName = .Name
              
        End Select
        .Dirty = AD_None
      End If
    End With
  Next ple
  
  RemoveCollItems Configs, collRemov
  
  Exit Sub
ERR_H:
  RemoveCollItems Configs, collRemov
  Err.Raise Err.Number
End Sub
