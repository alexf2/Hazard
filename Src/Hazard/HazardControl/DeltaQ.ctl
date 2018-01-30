VERSION 5.00
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A26A2C69-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTProgss.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DeltaQ 
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   HasDC           =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3810
   Begin VB.CommandButton cmdVisible 
      Caption         =   "П"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3285
      TabIndex        =   4
      ToolTipText     =   "Показать Hazard"
      Top             =   3600
      Width           =   360
   End
   Begin GreenTreeProgress.GTProgress gtp 
      Height          =   300
      Left            =   165
      ToolTipText     =   "Просчёт модели"
      Top             =   3915
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      PropA           =   1
      Object.Align           =   0
      Appearance      =   1
      BackColor       =   -2147483633
      BorderStyle     =   1
      Caption         =   "0%"
      Enabled         =   -1  'True
      FloodColor      =   -2147483635
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HelpContextID   =   0
      Max             =   100
      Min             =   0
      OLEDropMode     =   0
      Orientation     =   0
      Style           =   1
      Object.WhatsThisHelpID =   0
   End
   Begin MSComctlLib.ListView lstFac 
      Height          =   2895
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3285
      TabIndex        =   2
      Top             =   3210
      Width           =   360
   End
   Begin VB.CommandButton btnCalc 
      Caption         =   "Вычислить"
      Height          =   375
      Left            =   150
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   4290
      Width           =   3495
   End
   Begin VB.TextBox txtPath 
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   3195
      Width           =   3105
   End
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   75
      Top             =   45
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
      WindowText1     =   "DeltaQ"
      WindowKey1      =   "DeltaQ"
   End
   Begin MSComDlg.CommonDialog cdPath 
      Left            =   45
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "hzd"
      Filter          =   "Файлы hazard'а (*.hzd)|*.hzd|Все файлы (*.*)|*.*"
   End
   Begin GreenTreeProgress.GTProgress gt0 
      Height          =   300
      Left            =   165
      ToolTipText     =   "Модификация переменных"
      Top             =   3615
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   529
      PropA           =   1
      Object.Align           =   0
      Appearance      =   1
      BackColor       =   -2147483633
      BorderStyle     =   1
      Caption         =   "0%"
      Enabled         =   -1  'True
      FloodColor      =   -2147483635
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HelpContextID   =   0
      Max             =   100
      Min             =   0
      OLEDropMode     =   0
      Orientation     =   0
      Style           =   1
      Object.WhatsThisHelpID =   0
   End
End
Attribute VB_Name = "DeltaQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Default Property Values:
Const m_def_Path As String = ""
'Const m_def_IsActive = False
'Property Variables:
Dim m_Path As String
'Dim m_IsActive As Boolean

Private m_haApp As HazardApp
Dim m_arrNames() As String

Private m_lCntVal As Long
Private m_coll_Result As Collection
Private m_coll_ToCalc As Collection
Private m_s_ResultVars As String
Public Property Get FactorDelts() As Collection
  Set FactorDelts = m_coll_Result
End Property
Public Property Get ResultVars() As String
  ResultVars = m_s_ResultVars
End Property

Public Function GetToCalc() As Collection
  Set GetToCalc = m_coll_ToCalc
End Function




Public Property Get SelectedFNs() As Collection
  Dim vvt As ListItem
  Set SelectedFNs = New Collection
  For Each vvt In lstFac.ListItems
    If vvt.Checked Then SelectedFNs.Add m_arrNames(vvt.Tag - 1)
  Next vvt
End Property

Public Property Let SelectedFNs(ByVal cll As Collection)
  Dim vvt As ListItem
  For Each vvt In lstFac.ListItems
    vvt.Checked = False
  Next vvt
  
  Dim sTmp As Variant
  For Each sTmp In cll
    Dim lIdx As Long
    If VarType(sTmp) <> vbString Then Err.Raise vbObjectError + 2, "Property Let SelectedFNs", "В коллекции имён факторов опасности допустимы только строки"
    lIdx = NameToIdx(sTmp)
    If lIdx = -1 Then Err.Raise vbObjectError + 1, "Property Let SelectedFNs", "Ошибочный фактор опасности: " & sTmp
    lstFac.ListItems(lIdx + 1).Checked = True
  Next sTmp
End Property

Private Function NameToIdx(ByVal sName As String) As Long
  Dim l As Long
  For l = LBound(m_arrNames) To UBound(m_arrNames)
    If StrComp(m_arrNames(l), sName, vbTextCompare) = 0 Then
      NameToIdx = l
      Exit Function
    End If
  Next l
  NameToIdx = -1
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
  BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  UserControl.BackStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  UserControl.Refresh
End Sub

Private Function NeedOpen() As Boolean
  NeedOpen = True
  If Not m_haApp.HaveDoc Then Exit Function
  If StrComp(m_haApp.NameDoc, Path, vbTextCompare) <> 0 Then Exit Function
  NeedOpen = False
End Function

Private Sub IncrementVars(ByVal bFirstInc As Boolean)
  Dim objGN As Object
  Set objGN = m_haApp.GertNetMainDsp
  If bFirstInc Then m_lCntVal = 0 Else m_lCntVal = m_lCntVal + 1
  Dim vvt
  For Each vvt In m_coll_ToCalc
    objGN.Factors(vvt).Value = m_lCntVal
  Next vvt
End Sub

Private Sub btnCalc_Click()
  On Error GoTo ERR_H
  If Not m_haApp Is Nothing Then
    If m_haApp.IsCalcAny Then
      m_haApp.CancelCalc
      EndCalc
      Exit Sub
    End If
  Else
    Set m_haApp = CreateObject("Hazard.HazardApp")
    cmdVisible.Enabled = True
  End If
  
  Path = txtPath.Text
  
  Set m_coll_ToCalc = Me.SelectedFNs
  Set m_coll_Result = New Collection
  
  Dim vvt: m_s_ResultVars = ""
  For Each vvt In m_coll_ToCalc
    If m_s_ResultVars <> "" Then m_s_ResultVars = m_s_ResultVars & "; "
    m_s_ResultVars = m_s_ResultVars & vvt
  Next vvt
  
  If NeedOpen Then
    m_haApp.OpenDoc Path, False
    m_haApp.GertNetMainDsp.Snapshot True
  End If
  
  If m_coll_ToCalc.Count < 1 Then Exit Sub
  
  Dim objGN As Object
  Set objGN = m_haApp.GertNetMainDsp
  With objGN
    .Revert
    IncrementVars True
    .NotifyWnd = hwnd
    m_haApp.StartCalc .K, .N, .RndBase, TP_BELOW_NORMAL, Nothing
  End With
  
  cmdPath.Enabled = False
  gtp.Value = 0: gtp.Caption = "0%"
  gt0.Value = 0: gt0.Caption = "0%"
  btnCalc.Caption = "Прервать"
  'UserControl.ForeColor = &HFF&
  gtp.ForeColor = &HFF&
  btnCalc.FontBold = True
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
End Sub

Private Sub EndCalc()
  cmdPath.Enabled = True
  btnCalc.Caption = "Вычислить"
  'UserControl.ForeColor = &H0&
  gtp.ForeColor = &HFFFFFF
  btnCalc.FontBold = False
  m_lCntVal = 0
End Sub

Private Sub cmdPath_Click()
  With cdPath
    .CancelError = True
    .DialogTitle = "Открыть файл модели Hazard"
    .Flags = FileOpenConstants.cdlOFNExplorer Or FileOpenConstants.cdlOFNHideReadOnly Or FileOpenConstants.cdlOFNLongNames Or FileOpenConstants.cdlOFNFileMustExist
    If .InitDir = "" Then _
      .InitDir = App.Path
                
    .FileName = ""
                
    On Error GoTo ERR_H0
    .ShowOpen
    DoEvents
    
    On Error GoTo 0
                
    txtPath.Text = .FileName
    Path = .FileName
    .InitDir = .FileName
    Exit Sub
ERR_H0:
  End With
End Sub

Private Sub cmdVisible_Click()
  If m_haApp Is Nothing Then Exit Sub
  If m_haApp.Visible Then
    m_haApp.Visible = False
    cmdVisible.Caption = "П"
    cmdVisible.ToolTipText = "Показать Hazard"
  Else
    m_haApp.Visible = True
    cmdVisible.Caption = "С"
    cmdVisible.ToolTipText = "Скрыть Hazard"
  End If
End Sub

Private Sub lstFac_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  PropertyChanged "SelectedFactors"
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'  Select Case PropertyName
'    Case "Font"
'      Set UserControl.Font = Ambient.Font
'    Case "ForeColor"
'      UserControl.ForeColor = Ambient.ForeColor
'  End Select
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
  UserControl.Appearance() = New_Appearance
  PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
  hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Initialize()
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnEndOfWork", Value:=NM_OnEndOfWork, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnStepOfWork", Value:=NM_OnStepOfWork, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnCancel", Value:=NM_OnCancel, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnEndOfWork2", Value:=NM_OnEndOfWork2, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnStepOfWork2", Value:=NM_OnStepOfWork2, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnEndOfWork3", Value:=NM_OnEndOfWork3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnStepOfWork3", Value:=NM_OnStepOfWork3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnCancel3", Value:=NM_OnCancel3, ProcessDefaultProc:=False
  GTMsgHook1.Windows.Item("DeltaQ").Messages.Add _
    Key:="NM_OnErrorCalc", Value:=NM_OnErrorCalc, ProcessDefaultProc:=False

  GTMsgHook1.Subclass = True


  Dim arrNames As Variant
  arrNames = Array("H01", "H02", "H03", "H04", "H05", "H06", "H07", "H08", "H09", "H12", "H13", "H14", _
   "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", _
   "T01", "T02", "T03", "T04", "T05", "T06", _
   "C01", "C02", "C03", "C04")
   
  Dim l As Long, liTmp As ListItem
   
  ReDim m_arrNames(UBound(arrNames)) As String
  For l = LBound(arrNames) To UBound(arrNames)
    m_arrNames(l) = arrNames(l)
  Next l
  
 On Error GoTo ERR_H
  For l = LBound(m_arrNames) To UBound(m_arrNames)
    Set liTmp = lstFac.ListItems.Add(, , m_arrNames(l))
    liTmp.Tag = liTmp.Index
  Next l
  Exit Sub
ERR_H:
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'  Set UserControl.Font = Ambient.Font
'  UserControl.BackColor = Ambient.BackColor
'  UserControl.ForeColor = Ambient.ForeColor
'  m_IsActive = m_def_IsActive
  m_Path = m_def_Path
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  Set UserControl.Font = PropBag.ReadProperty("Font")
  UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
'  m_IsActive = PropBag.ReadProperty("IsActive", m_def_IsActive)
  
  m_Path = PropBag.ReadProperty("Path", m_def_Path)
  txtPath.Text = m_Path
  
  Dim sTmp As String, l As Long, cllTmp As New Collection
  
  sTmp = PropBag.ReadProperty("SelectedFactors")
  For l = Len(sTmp) To 1 Step -1
    If Mid$(sTmp, l, 1) = "1" Then _
      cllTmp.Add m_arrNames(l - 1)
  Next l
  SelectedFNs = cllTmp
End Sub

Private Sub UserControl_Terminate()
  GTMsgHook1.Subclass = False
  
  On Error GoTo ERR_H
  If Not m_haApp Is Nothing Then
    If m_haApp.IsCalcAny Then m_haApp.CancelCalc
    m_haApp.Quit
    Set m_haApp = Nothing
  End If
  
  Exit Sub
  
ERR_H:
  Set m_haApp = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("Font", UserControl.Font)
  Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
'  Call PropBag.WriteProperty("IsActive", m_IsActive, m_def_IsActive)
  Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
  
  
  Dim vvt As ListItem, sTmp As String
  For Each vvt In lstFac.ListItems
    sTmp = sTmp & IIf(vvt.Checked, "1", "0")
  Next vvt
  
  PropBag.WriteProperty "SelectedFactors", sTmp
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,""
Public Property Get Path() As String
  Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
  m_Path = New_Path
  PropertyChanged "Path"
  txtPath.Text = m_Path
End Property


Private Sub GTMsgHook1_Message(ByVal hwnd As Long, ByVal MsgID As Long, ByVal wParam As Long, ByVal lParam As Long)
  On Error GoTo ERR_H
  Select Case MsgID
    Case NM_OnEndOfWork
      
      Dim gn As MGertNet
      Set gn = m_haApp.GN_Run
      
      Dim l1 As Long, l2 As Long, m As Double, d As Double
      gn.GetInfSampleK 73, l1, l2, m, d
      Dim mp As New MyPoint
      mp.x = m_lCntVal: mp.y = m
      m_coll_Result.Add mp
      
      gt0.Value = m_lCntVal * 10
      gt0.Caption = CStr(gt0.Value) & "%"
      If m_lCntVal >= 10 Then
       EndCalc
       'm_haApp.GertNetMainDsp.Revert
       m_haApp.CloseCalc
      Else
        Dim objGN As Object
        Set objGN = m_haApp.GertNetMainDsp
        objGN.Revert
        IncrementVars False
        m_haApp.StartCalc 200, 500000, 1, TP_BELOW_NORMAL, Nothing
        gtp.Value = 0: gtp.Caption = "0%"
      End If
    Case NM_OnStepOfWork
      gtp.Value = wParam
      gtp.Caption = CStr(wParam) & "%"
    Case NM_OnCancel
      EndCalc
      m_haApp.CloseCalc
    Case NM_OnEndOfWork2
    Case NM_OnStepOfWork2
    Case NM_OnEndOfWork3
    Case NM_OnStepOfWork3
    Case NM_OnCancel3
    Case NM_OnErrorCalc
      EndCalc
      m_haApp.CloseCalc
      MsgBox m_haApp.GN_Run.LastCalcError, vbCritical Or vbOKOnly, "Ошибка вычисления"
  End Select
  Exit Sub
ERR_H:
  EndCalc
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
End Sub

