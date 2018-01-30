VERSION 5.00
Object = "{E07BAAA3-2DEE-11D2-B2E4-006008BF0C53}#2.0#0"; "OtxMenu.dll"
Object = "{7FCAEF84-D390-11D0-8849-006097BFD99B}#2.0#0"; "formx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Hazard monitor"
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6165
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "frmMain"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   150
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "hzd"
      Filter          =   "Файлы hazard'а (*.hzd)|*.hzd|Все файлы (*.*)|*.*"
   End
   Begin FormXLib.OtxCommandManager OtxCommandManager1 
      Left            =   630
      Top             =   1035
      _cx             =   953
      _cy             =   953
      DefaultTextAlignment=   3
      DisplayMode     =   3
      LargeIcons      =   0   'False
      CoolLook        =   -1  'True
      ToolTips        =   -1  'True
      HighWaterMark   =   62160
      SmallImageWidth =   16
      SmallImageHeight=   16
      LargeImageWidth =   32
      LargeImageHeight=   32
      Count           =   0
      CommandGroupCount=   4
      BeginProperty Group1 {44F69A45-D928-11D1-9CD7-00C04F91E286} 
         Name            =   "File"
         NumCommands     =   8
         BeginProperty Command1 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "New"
            Picture         =   "MainForm.frx":0442
            TooltipText     =   "Создать новую модель"
            Caption         =   "&Новая"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command2 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Open"
            Picture         =   "MainForm.frx":0554
            TooltipText     =   "Загрузить модель из файла"
            Caption         =   "&Загрузить"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command3 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Save"
            Picture         =   "MainForm.frx":0666
            TooltipText     =   "Сохранить изменённую модель"
            Caption         =   "&Сохранить"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command4 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "SaveAs"
            Picture         =   "MainForm.frx":0778
            TooltipText     =   "С&охранить в новом файле"
            Caption         =   "С&охранить..."
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command5 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Snapshot"
            TooltipText     =   ""
            Caption         =   "Снимок параметров модели"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command6 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Revert"
            TooltipText     =   ""
            Caption         =   "Восстановить модель из снимка"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command7 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Exit"
            TooltipText     =   ""
            Caption         =   "Выход"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command8 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Module"
            Picture         =   "MainForm.frx":088A
            TooltipText     =   ""
            Caption         =   "Модуль"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
      EndProperty
      BeginProperty Group2 {44F69A45-D928-11D1-9CD7-00C04F91E286} 
         Name            =   "Item"
         NumCommands     =   2
         BeginProperty Command1 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Model"
            Picture         =   "MainForm.frx":09F4
            TooltipText     =   "Монитор модели"
            Caption         =   "&Модель"
            Description     =   ""
            Style           =   2
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command2 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "SF"
            Picture         =   "MainForm.frx":0B06
            TooltipText     =   "Монитор комплексов мер безопасности"
            Caption         =   "М&еры"
            Description     =   ""
            Style           =   2
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
      EndProperty
      BeginProperty Group3 {44F69A45-D928-11D1-9CD7-00C04F91E286} 
         Name            =   "View"
         NumCommands     =   3
         BeginProperty Command1 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "ToolBar"
            TooltipText     =   "Показать/скрыть панель инструментов"
            Caption         =   ""
            Description     =   ""
            Style           =   2
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command2 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "Reports"
            TooltipText     =   "Показать/скрыть отчёты"
            Caption         =   ""
            Description     =   ""
            Style           =   2
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command3 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "GraphBar"
            TooltipText     =   "Показать/скрыть графики"
            Caption         =   ""
            Description     =   ""
            Style           =   2
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
      EndProperty
      BeginProperty Group4 {44F69A45-D928-11D1-9CD7-00C04F91E286} 
         Name            =   "Reports"
         NumCommands     =   3
         BeginProperty Command1 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "ModelRep"
            Picture         =   "MainForm.frx":0C18
            TooltipText     =   ""
            Caption         =   "ModelRep"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command2 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "SFRep"
            Picture         =   "MainForm.frx":0D2A
            TooltipText     =   ""
            Caption         =   "SFRep"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
         BeginProperty Command3 {644D4630-DE08-11D1-9CDB-00C04F91E286} 
            Key             =   "RepOpt"
            Picture         =   "MainForm.frx":0E3C
            TooltipText     =   ""
            Caption         =   "Настройка"
            Description     =   ""
            Style           =   0
            Enabled         =   -1  'True
            Checked         =   0
            MaskColor       =   -2
         EndProperty
      EndProperty
   End
   Begin FormXLib.OtxToolbar tbMain 
      Left            =   360
      Top             =   165
      _Version        =   131072
      _ExtentX        =   8943
      _ExtentY        =   952
      _StockProps     =   0
      Caption         =   "Файл"
      CommandList     =   "key=New!dm=0,key=Open!dm=0,key=Save!dm=0,key=SaveAs!dm=0,-1,key=Model!dm=0,key=SF!dm=0"
   End
   Begin FormXLib.MDIFormX frmxMain 
      Left            =   5280
      Top             =   3300
      _Version        =   131072
      _ExtentX        =   953
      _ExtentY        =   953
      _StockProps     =   0
   End
   Begin OTXMENULibCtl.OtxMenu OtxMenu1 
      Left            =   2805
      Top             =   1950
      _cx             =   953
      _cy             =   953
      BeginProperty MenuMap {56B749E7-62F5-11D2-A32C-A253D8000000} 
         Count           =   2
         BeginProperty Item1 {56B749E5-62F5-11D2-A32C-A253D8000000} 
            MenuKey         =   "RepModel"
            ToolBarKey      =   "ModelRep"
            MaskColor       =   -2
            DynamicRule     =   0   'False
         EndProperty
         BeginProperty Item2 {56B749E5-62F5-11D2-A32C-A253D8000000} 
            MenuKey         =   "RepSP"
            ToolBarKey      =   "SFRep"
            MaskColor       =   -2
            DynamicRule     =   0   'False
         EndProperty
      EndProperty
      CommandManagerUISync=   -1  'True
      CommandManagerMenuHandler=   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Файл"
      NegotiatePosition=   1  'Left
      Begin VB.Menu New 
         Caption         =   "&Новый"
      End
      Begin VB.Menu Open 
         Caption         =   "&Загрузить"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Save 
         Caption         =   "&Сохранить"
         Shortcut        =   {F2}
      End
      Begin VB.Menu SaveAs 
         Caption         =   "С&охранить в новом файле"
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu Snapshot 
         Caption         =   "'Снимок' параметров модели"
         Shortcut        =   {F11}
      End
      Begin VB.Menu Revert 
         Caption         =   "Восстановить параметры из 'снимка'"
         Shortcut        =   {F12}
      End
      Begin VB.Menu s11 
         Caption         =   "-"
      End
      Begin VB.Menu Module 
         Caption         =   "Модуль вычисления оценок ФО"
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Вы&ход"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Вид"
      NegotiatePosition=   1  'Left
      Begin VB.Menu ToolBar 
         Caption         =   "&Панель инструментов"
      End
      Begin VB.Menu Reports 
         Caption         =   "&Отчёты"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Отчёты"
      Begin VB.Menu RepModel 
         Caption         =   "Модель"
      End
      Begin VB.Menu RepSP 
         Caption         =   "Текущий комплекс мер"
      End
      Begin VB.Menu s22 
         Caption         =   "-"
      End
      Begin VB.Menu RepOpt 
         Caption         =   "Настройка"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Справка"
      NegotiatePosition=   3  'Right
      Begin VB.Menu About 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    ' N.B.  SystemsParametersInfo() is used for screen dimensions instead
    '       instead of Screen.Width and Screen.Height to account for systray.


Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Setff Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private m_b_ArgsChecked As Boolean

Private Sub About_Click()
  On Error GoTo ERR_H
  Load frmAbout
  frmAbout.Show vbModal, frmMain
  Unload frmAbout
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  On Error Resume Next
  Unload frmAbout
End Sub

Private Sub Exit_Click()
'  Dim c As Integer
'  c = 0
'  MDIForm_QueryUnload c, vbFormControlMenu
'  If c = 0 Then Unload Me
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("Exit").Execute
  
'    Dim cmdGroup As OtxCommandGroup
'For Each cmdGroup In OtxCommandManager1.CommandGroups
'   Dim cmd As OtxStandardCommand
'   For Each cmd In cmdGroup
'      MsgBox "Command=" & cmd.key & ", Group=" & cmdGroup.Name
'   Next
'Next

End Sub

'Private Sub GraphBar_Click()
'  OtxCommandManager1.CommandGroups("View")("GraphBar").Execute
'End Sub

Private Sub MDIForm_Activate()
  'frxSPMonitor.SetFocus
  'frmSP.SetFocus
'SetActiveWindow frmSP.hwnd
  'Setff frmSP.hwnd
  'SetForegroundWindow frmSP.hwnd
  'frmSP.ZOrder 0
'Setff frmSP.hwnd
 
  'Set frmSP.ActiveControl = frmSP.pvnCost
  'SendMessage frxSPMonitor.vsTabModel.hwnd, WM_ACTIVATE, WA_CLICKACTIVE, 0
  'Setff frxSPMonitor.vsTabModel.hwnd
  'SendMessage frmSP.hwnd, WM_ACTIVATE, WA_CLICKACTIVE, 0
  'frmSP.pvnCost.SetFocus
  'Setff frmSP.hwnd
End Sub

Private Sub MDIForm_Deactivate()
'Dim ii
'ii = 1
End Sub

Private Sub MDIForm_Initialize()
  m_b_ArgsChecked = False
End Sub

Private Function ScreenRegKey() As String
  ScreenRegKey = S_RegAppKey & "MainWindow\" & CStr(GetSystemMetrics(SM_CXSCREEN)) & "x" & CStr(GetSystemMetrics(SM_CYSCREEN))
End Function

Private Sub CentrMDI()
  Dim workspace(1 To 4) As Long ' (1)=left (2)=top (3)=right (4)=bottom
  SystemParametersInfo &H30, 0, workspace(1), 0 ' &h30 = get workspace size
  Dim x As Long, y As Long, w As Long, h As Long
  x = workspace(1) * Screen.TwipsPerPixelX
  y = workspace(2) * Screen.TwipsPerPixelY
  w = workspace(3) * Screen.TwipsPerPixelX - x
  h = workspace(4) * Screen.TwipsPerPixelY - y
  Move x, y, w, h   ' 10% margins
End Sub

Private Sub MDIForm_Load()

    Dim lhKey As Long, sKey  As String, rc As Long
    sKey = ScreenRegKey()
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, lhKey)
    If rc = ERROR_SUCCESS Then
      Dim lLeft As Long, lTop As Long, lHeight As Long, lWidth As Long, lState As Long
      Dim lValueLen As Long
      lValueLen = 4
      rc = RegQueryValueEx2(lhKey, "State", 0&, 0&, VarPtr(lState), VarPtr(lValueLen))
      If rc = ERROR_SUCCESS Then
        On Error Resume Next
        If lState = vbNormal Then
          lValueLen = 4
          rc = RegQueryValueEx2(lhKey, "Left", 0&, 0&, VarPtr(lLeft), VarPtr(lValueLen))
          If rc = ERROR_SUCCESS Then
            lValueLen = 4
            rc = RegQueryValueEx2(lhKey, "Top", 0&, 0&, VarPtr(lTop), VarPtr(lValueLen))
            If rc = ERROR_SUCCESS Then
              lValueLen = 4
              rc = RegQueryValueEx2(lhKey, "Width", 0&, 0&, VarPtr(lWidth), VarPtr(lValueLen))
              If rc = ERROR_SUCCESS Then
                lValueLen = 4
                rc = RegQueryValueEx2(lhKey, "Height", 0&, 0&, VarPtr(lHeight), VarPtr(lValueLen))
                If rc = ERROR_SUCCESS Then
                  Me.WindowState = vbNormal
                  Move lLeft, lTop, lWidth, lHeight
                End If
              End If
            End If
          End If
        Else
          CentrMDI
          Me.WindowState = lState
        End If
      End If
      RegCloseKey lhKey
      On Error GoTo 0
    Else
      CentrMDI
      ' make main form large:  entire workspace less 10% margins on each side
    End If
    
    With OtxCommandManager1.ToolBars("Файл").CommandUIs("Model").Command
      .Style = otxCheck
      .Checked = otxUnpressed
    End With
    With OtxCommandManager1.ToolBars("Файл").CommandUIs("SF").Command
      .Style = otxCheck
      .Checked = otxUnpressed
    End With
            
    
    ' show the child forms
    '      The names of these forms are hard-wired based on the project template "OTX Docking Forms.vbp".
    '      Simply delete the lines that do not apply if you diverge from that template.
'    frmDocked.Show
'    frmDockedOnly.Show
'    frmMDIChild.Show
'    frmStdMDIChild.Show
'    frmFloating.Show
'    frmFloatingOnly.Show
     Load frmCancel
     Dim lX As Long, lY As Long
     frmCancel.GetSz lX, lY
     frmCancel.Dock otxTopEdge, 0, 1, 25, otxPosEndRel Or otxPanelEndRel Or otxPanelSpread, lY
     'frmCancel.Dock otxTopEdge, 0, 1, , otxPosEndRel Or otxPanelEndRel
     frmCancel.Show
          
     
     Load frxRep
     frxRep.Show
     
     CommonDialog1.InitDir = App.Path
     
     Load frmGraphBar
     frmGraphBar.Dock otxBottomEdge, 0, 1, 50, otxPosEndRel Or otxPanelEndRel
                 
     DoEvents
                          
End Sub

Public Property Get MeX() As MDIFormX
    ' Property MeX is the MDIFormX equivalent of the keyword Me
    Set MeX = frmxMain
    ' N.B. A more general definition for MeX that does not hardwire the name of the embedded MDIFormX control
    '      is as follows.  The more general definition is less efficient because it calls a helper function.
    '   Public Property Get MeX() as MDIFormX
    '       Set MeX = otxMDIFormX()
    '   End Property
    '
    '   The general definition is convenient if you paste the definition into MDIForm modules
    '   that use different names for the embedded MDIFormX control.
End Property

Public Property Get DefAutoFrame() As CollectionX
    Set DefAutoFrame = MeX.DefAutoFrame
End Property

Public Property Set DefAutoFrame(ByVal NewValue As CollectionX)
    Set MeX.DefAutoFrame = NewValue
End Property

Public Property Get DefCanDragDocked() As Boolean
    DefCanDragDocked = MeX.DefCanDragDocked
End Property

Public Property Let DefCanDragDocked(ByVal NewValue As Boolean)
    MeX.DefCanDragDocked = NewValue
End Property

Public Property Get DefClientEdge() As Boolean
    DefClientEdge = MeX.DefClientEdge
End Property

Public Property Let DefClientEdge(ByVal NewValue As Boolean)
    MeX.DefClientEdge = NewValue
End Property

Public Property Get DefUseCaptionButton() As Boolean
    DefUseCaptionButton = MeX.DefUseCaptionButton
End Property

Public Property Let DefUseCaptionButton(ByVal NewValue As Boolean)
    MeX.DefUseCaptionButton = NewValue
End Property

Public Property Get DefUseCaptionDblClick() As Boolean
    DefUseCaptionDblClick = MeX.DefUseCaptionDblClick
End Property

Public Property Let DefUseCaptionDblClick(ByVal NewValue As Boolean)
    MeX.DefUseCaptionDblClick = NewValue
End Property

Public Property Get DefUseCaptionRtDblClick() As Boolean
    DefUseCaptionRtDblClick = MeX.DefUseCaptionRtDblClick
End Property

Public Property Let DefUseCaptionRtDblClick(ByVal NewValue As Boolean)
    MeX.DefUseCaptionRtDblClick = NewValue
End Property

Public Property Get FrameType() As otxMDIFormFrameType
    FrameType = MeX.FrameType
End Property

Public Property Let FrameType(ByVal NewValue As otxMDIFormFrameType)
    MeX.FrameType = NewValue
End Property

Public Function Forms(Optional ByVal Types, Optional ByVal Options, Optional ByVal Attribs) As CollectionX
    Set Forms = MeX.Forms(Types, Options, Attribs)
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ERR_H
  
  If haApp.IsCalcAny Or Not frmTM.CanClose Then
    MsgBox "Перед завершением приложения необходимо завершить все вычислительные процессы", vbCritical Or vbOKOnly, "Ошибка"
    Cancel = 1
    Exit Sub
  End If
  
  'If UnloadMode <> vbFormCode Then
    If Not haApp.GertNetMain Is Nothing Then
      Dim lRes As Long
      lRes = haApp.CloseDoc(True)
      If lRes = vbCancel Then _
        Cancel = 1: Exit Sub
    End If
'  Else
'    haApp.CloseDoc False
'  End If
  
  Exit Sub
  
ERR_H:
  Cancel = 1
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  Unload frmPlugin

  Dim ff As Form
  For Each ff In Me.Forms
    Unload ff
  Next ff
    
  
'  Unload frmCancel
  Unload frxRep
'  Unload frmGraphBar
  Unload frmSave
  Unload frmIdx
  Unload frmAlhoOpt
  Unload frmApply
'  Unload frxModelMon
'  Unload frxSPMonitor
'  Set frxModelMon = Nothing
'  Set frxSPMonitor = Nothing
'  Set frmCancel = Nothing
'  Set frxRep = Nothing
'  Set frmGraphBar = Nothing

  
  Unload frmAbout
  
  'Set frmProcs = Nothing
  Set frmCalibrate = Nothing
  Set frmMEditor = Nothing
  
  Set frmComplSP = Nothing
  Set frmOptimiz = Nothing
  Set frmSP = Nothing
  
  Set frmMon = Nothing
  
  
  Set haApp = Nothing
  
  Dim lhKey As Long, sKey  As String, rc As Long, lDispos As Long
  sKey = ScreenRegKey()
  'rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, lhKey)
  'If rc <> ERROR_SUCCESS Then
  rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lhKey, lDispos)
  
  If rc = ERROR_SUCCESS Then
    On Error Resume Next
    Dim lValueType As Long, lValueLen As Long, lv As Long
    lValueLen = 4
    lv = Me.WindowState
    rc = RegSetValueEx(lhKey, "State", 0&, REG_DWORD, VarPtr(lv), lValueLen)
    If Me.WindowState = vbNormal Then
      lv = Me.Left
      rc = RegSetValueEx(lhKey, "Left", 0&, REG_DWORD, VarPtr(lv), lValueLen)
      lv = Me.Top
      rc = RegSetValueEx(lhKey, "Top", 0&, REG_DWORD, VarPtr(lv), lValueLen)
      lv = Me.Width
      rc = RegSetValueEx(lhKey, "Width", 0&, REG_DWORD, VarPtr(lv), lValueLen)
      lv = Me.Height
      rc = RegSetValueEx(lhKey, "Height", 0&, REG_DWORD, VarPtr(lv), lValueLen)
    End If
    RegCloseKey lhKey
    On Error GoTo 0
  End If
  
End Sub

Private Sub New_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("New").Execute
End Sub

Private Sub Open_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("Open").Execute
End Sub

Private Sub OtxCommandManager1_Handle(ByVal Command As FormXLib.IOtxCommand, context As Variant, stopRouting As Boolean)
  On Error GoTo ERR_H
  
  Dim bm As CBeam
  
  Select Case Command.Group.Name
    Case "File"
      Select Case Command.Key
        Case "New"
          stopRouting = True
                    
          Set bm = New CBeam
          bm.Beam Me
          DoEvents
          
          haApp.NewDoc True
          UpdateCaption
          
        Case "Open"
          stopRouting = True
          
          With frmMain.CommonDialog1
            .CancelError = True
            .DialogTitle = "Открыть файл модели"
            .Flags = FileOpenConstants.cdlOFNExplorer Or FileOpenConstants.cdlOFNHideReadOnly Or FileOpenConstants.cdlOFNLongNames Or FileOpenConstants.cdlOFNFileMustExist
            If Trim(.InitDir) = "" Then _
              .InitDir = App.Path
                                                              
            .FileName = ""
                        
            On Error GoTo ERR_H0
            .ShowOpen
            DoEvents
            
            Dim sDir As String
            CutPath .FileName, sDir
            .InitDir = sDir
            
            On Error GoTo ERR_H
                        
            Set bm = New CBeam
            bm.Beam Me
            
            haApp.OpenDoc .FileName, True
            UpdateCaption
            Exit Sub
ERR_H0:
          End With
            
        Case "SaveAs"
          stopRouting = True
          If Not haApp.HaveDoc Then
            MsgBox "Нет документа", vbOKOnly Or vbExclamation, "Предупреждение"
          Else
                        
            haApp.SetNewDoc
            haApp.SaveDoc True, True, True
            UpdateCaption
            
'            If haApp.AskFileName() Then
'              DoEvents
'
'              bm.Beam Me
'
'              UpdateCaption
'              haApp.SaveDoc True, True, True
'            End If
          End If
          
        Case "Save"
          stopRouting = True
                                        
                    
          If Not haApp.IsNewDoc Then
            frmSave.ssNo.Enabled = False
            haApp.SelectSave True
            If frmSave.ModalResult = vbOK Then
              Set bm = New CBeam
              bm.Beam Me
              DoEvents
              haApp.SaveDoc True, frmSave.SaveModel, frmSave.SaveSP
              UpdateCaption
            End If
          Else
            Set bm = New CBeam
            bm.Beam Me
            DoEvents
            haApp.SaveDoc True
            UpdateCaption
          End If
          
        Case "Snapshot", "Revert"
          stopRouting = True
          
          DoEvents
          
          If Command.Key = "Snapshot" Then
            Set bm = New CBeam
            bm.Beam Me
            haApp.GertNetMain.Snapshot True
          Else
          
            If MsgBox("Восстановить параметры модели ?", vbYesNo Or vbQuestion, "Подтвердите операцию") <> vbYes Then Exit Sub
          
            Set bm = New CBeam
            bm.Beam Me
            haApp.GertNetMain.Revert
            
            frmMEditor.UpdateValuesAll
            
            If frmMEditor.PsscheckPreview.Value = True Then
              haApp.GertNetMain.Run 5, 10, False, -1, True
              frmMEditor.ShowSPFunc frmMEditor.CurrFac
            End If
          End If
          
        Case "Module"
          stopRouting = True
          MakeAssModule
          
        Case "Exit"
          stopRouting = True
          Unload Me
          
      End Select
      
    Case "Item"
      Select Case Command.Key
        Case "Model"
          stopRouting = True
          'If OtxCommandManager1.ToolBars("Файл").CommandUIs("Model").Command.Checked Then
          If Command.Checked = otxPressed Then
            If MMNothing() Then
              Load frxModelMon
              frxModelMon.Show
            Else
              If frxSPMonitor.WindowState = vbMaximized Then frxSPMonitor.Hide
              DoEvents
              frxModelMon.ZOrder
              'If Not SPNothing() Then frxSPMonitor.Visible = False
            End If
            If Not frxModelMon.Visible Then frxModelMon.Visible = True
            If haApp.HaveDoc And Not frxModelMon.IsBoundModel Then _
              frxModelMon.BindModel haApp.GertNetMain
          Else
            Command.Checked = otxPressed
          End If
          DoEvents
          UpdateWindow frxModelMon.hWnd
        Case "SF"
          stopRouting = True
          'If OtxCommandManager1.ToolBars("Файл").CommandUIs("SF").Command.Checked Then
           If Command.Checked = otxPressed Then
            If SPNothing() Then
              Load frxSPMonitor
              frxSPMonitor.Show
            Else
              If frxModelMon.WindowState = vbMaximized Then frxModelMon.Hide
              DoEvents
              frxSPMonitor.ZOrder
              
              'If Not MMNothing() Then frxModelMon.Visible = False
            End If
            If Not frxSPMonitor.Visible Then frxSPMonitor.Visible = True
            
            If Not frxSPMonitor.IsBoundModel Then
              If haApp.HaveDoc And (haApp.XCollection Is Nothing) Then
                haApp.OpenSF True
              ElseIf Not haApp.XCollection Is Nothing Then
                If haApp.CurrSFColl <> -1 Then _
                  frxSPMonitor.BindModel haApp.XCollection
              End If
            End If
          Else
            Command.Checked = otxPressed
          End If
          DoEvents
          UpdateWindow frxSPMonitor.hWnd
      End Select
      
    Case "View"
      Select Case Command.Key
        Case "ToolBar"
          stopRouting = True
          OtxCommandManager1.ToolBars("Файл").Visible = Not OtxCommandManager1.ToolBars("Файл").Visible
          
        Case "Reports"
          stopRouting = True
          frxRep.Visible = Not frxRep.Visible
          
'        Case "GraphBar"
'          stopRouting = True
'          frmGraphBar.Visible = Not frmGraphBar.Visible
      End Select
      
    Case "Reports"
      Select Case Command.Key
        Case "ModelRep"
          stopRouting = True
          MakeModelReport
        Case "SFRep"
          stopRouting = True
          MakeSFReport
        Case "RepOpt"
          stopRouting = True
          MakeRepOpt
      End Select
  End Select
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
  Err.Clear
End Sub

Private Sub OtxCommandManager1_InitCommandUI(ByVal CommandUI As FormXLib.IOtxCommandUI)
'Dim ii
'ii = 1
End Sub

Private Sub OtxCommandManager1_UpdateUI(ByVal Command As FormXLib.IOtxCommand, stopRouting As Boolean)
  On Error GoTo ERR_H
    
  
  If Not m_b_ArgsChecked And Not haApp Is Nothing Then
    m_b_ArgsChecked = True
    
    'haApp.OpenDoc "g:\t1\t1.2hzd", True
    
    'MsgBox Interaction.Command$(), vbOKOnly, "hh"
    Dim bm As CBeam
    Dim sCmd As String, sPref2 As String, sPref3 As String
    sCmd = Trim$(Interaction.Command$())
    sPref2 = Left$(sCmd, 2)
    sPref3 = Left$(sCmd, 3)
    
    If sPref2 = "/n" Then
      Set bm = New CBeam
      bm.Beam Me
      haApp.NewDoc True
    ElseIf sPref2 = "/o" Then
      Set bm = New CBeam
      bm.Beam Me
'MsgBox Right$(sCmd, Len(sCmd) - 3)
      haApp.OpenDoc Right$(sCmd, Len(sCmd) - 3), True
      UpdateCaption
    ElseIf sPref3 = "/r2" Then
      RegisterShellext False
    ElseIf sPref3 = "/u2" Then
      UnregisterShellext
    ElseIf sPref2 = "/?" Or sPref2 = "/h" Then
      MsgBox "/n - новая модель" & vbCrLf & "/o - открыть модель" _
         & vbCrLf & "/r2 - зарегистрировать shell extensions" & vbCrLf & "/u2 - отменить регистрацию", vbInformation Or vbOKOnly, "Ключи"
    End If
  End If
    
  'If Not OtxCommandManager1.ToolBars("Файл").Visible Then OtxCommandManager1.ToolBars("Файл").Visible = True
    
  Dim bFl As Boolean
  bFl = True
  Select Case Command.Group.Name
    Case "File"
      Select Case Command.Key
        Case "New", "Open"
          Command.Enabled = Not haApp.IsCalcAny
          stopRouting = True
            
        Case "SaveAs", "Save", "Snapshot", "Revert", "Module"
          stopRouting = True
          If Not haApp.HaveDoc Then bFl = False
          Command.Enabled = bFl
          
      End Select
      
    Case "Item"
      Select Case Command.Key
        Case "Model", "SF"
          stopRouting = True
          
      End Select
      
    Case "View"
      Select Case Command.Key
        Case "ToolBar"
          stopRouting = True
          Command.Checked = IIf(OtxCommandManager1.ToolBars("Файл").Visible, otxPressed, otxUnpressed)
          
        Case "Reports"
          stopRouting = True
          Command.Checked = IIf(frxRep.Visible = True, otxPressed, otxUnpressed)
          
'        Case "GraphBar"
'          stopRouting = True
'          Command.Checked = IIf(frmGraphBar.Visible = True, otxPressed, otxUnpressed)
          
      End Select
    Case "Reports"
      Select Case Command.Key
        Case "ModelRep", "RepOpt"
          stopRouting = True
          Command.Enabled = haApp.HaveDoc
        Case "SFRep"
          stopRouting = True
          Command.Enabled = haApp.CurrSFColl <> -1
          'Command.Enabled = False
      End Select
  End Select
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
  Err.Clear
End Sub

Friend Sub UpdateCaption()
  If Not haApp.HaveDoc Then
    Caption = "Hazard"
  Else
    Caption = "Hazard - [" & haApp.NameDoc & "]"
  End If
End Sub

Friend Sub MakeModelReport()
  If haApp.Rep1 Is Nothing Or haApp.GertNetMain Is Nothing Then
    MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
    Exit Sub
  End If
  
  Dim rep As New CRepModel
  rep.InitReport haApp.GertNetMain, haApp
  Dim ir As IRepItem
  Set ir = rep
  haApp.Rep1.Add ir
  
  UpdateGUI_Rep1 haApp
End Sub

Friend Sub MakeSFReport()
  If haApp.Rep2 Is Nothing Or haApp.GertNetMain Is Nothing Then
    MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
    Exit Sub
  End If
  
  Dim rep As New CRepSF
  frmSP.LoadCurrSP False
  rep.InitReport haApp.XCollection(haApp.CurrSFColl), haApp.XCollection.Key(haApp.CurrSFColl), frmSP.P0, haApp
  Dim ir As IRepItem
  Set ir = rep
  haApp.Rep2.Add ir
  
  UpdateGUI_Rep2 haApp
  
End Sub
        
Friend Sub MakeRepOpt()
  
  On Error GoTo ERR_H
  With frmRepOpt
    .AssData haApp
    .Show vbModal, frmMain
    DoEvents
    
    If .ModalResult Then
      haApp.CreRepRun = .ssRun.Value
      haApp.CreRepOpt = .ssOpt.Value
      haApp.CreRepVApply = .ssVApply.Value
      If haApp.HaveDoc Then
        With haApp.GertNetMain
          .NFormatPr = IIf(frmRepOpt.ssNormalPr.Value, NF_Normal, NF_Scientific)
          .NFormatOther = IIf(frmRepOpt.ssNormalOther.Value, NF_Normal, NF_Scientific)
          .NDigitsPr = frmRepOpt.gtsPr.Value
          .NDigitsOther = frmRepOpt.gtsOther.Value
        End With
      End If
      
      CallSetupN frmCalibrate
      CallSetupN frmMEditor
      CallSetupN frmMon
      CallSetupN frmComplSP
      CallSetupN frmOptimiz
      CallSetupN frmSP

    End If
  End With
  
  Exit Sub
ERR_H:
  
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CallSetupN(ByVal obj As Object)
  If Not obj Is Nothing Then obj.SetupNumbers
End Sub

Private Sub RepModel_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("Reports")("RepModel").Execute
End Sub

Private Sub RepOpt_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("Reports")("RepOpt").Execute
End Sub

Private Sub Reports_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("View")("Reports").Execute
End Sub

Private Sub RepSP_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("Reports")("RepSP").Execute
End Sub

Private Sub Revert_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("Revert").Execute
End Sub

Private Sub Save_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("Save").Execute
End Sub

Private Sub SaveAs_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("SaveAs").Execute
End Sub

Private Sub Snapshot_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("File")("Snapshot").Execute
End Sub

Private Sub ToolBar_Click()
  If Me.Enabled Then _
    OtxCommandManager1.CommandGroups("View")("ToolBar").Execute
End Sub

Private Sub MakeAssModule()
  On Error GoTo ERR_H
    
  frmAssMgr.AssData haApp.GertNetMain
  frmAssMgr.Show vbModal, frmMain
  DoEvents
  frmAssMgr.UnassData
  Unload frmAssMgr
  
  Exit Sub
ERR_H:
  frmAssMgr.UnassData
  Unload frmAssMgr
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
