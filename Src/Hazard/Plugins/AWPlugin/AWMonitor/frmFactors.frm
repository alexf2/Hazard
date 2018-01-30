VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{A26A2CE8-6B79-11D1-BF3C-000000000000}#1.1#0"; "GTMsghk.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmFactors 
   Caption         =   "Факторы опасности"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmFactors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInPlace 
      BackColor       =   &H00FFFBF0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1845
      TabIndex        =   11
      Text            =   "Text1"
      ToolTipText     =   "Вызов редактора - F2"
      Top             =   3255
      Visible         =   0   'False
      Width           =   2250
   End
   Begin GreenTreeMsgHook.GTMsgHook GTMsgHook1 
      Left            =   4785
      Top             =   3435
      _ExtentX        =   635
      _ExtentY        =   635
      PropA           =   1
      NumWindows      =   1
      WindowText1     =   "frmFactors"
   End
   Begin Threed.SSFrame ssfGOST 
      Height          =   2220
      Left            =   105
      TabIndex        =   1
      Top             =   3555
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3916
      _Version        =   131074
      Caption         =   "Хранилище ГОСТов"
      Begin VB.TextBox txtCurrStg 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1740
         Width           =   3825
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H00FFFBF0&
         Height          =   345
         Left            =   135
         TabIndex        =   4
         Top             =   1080
         Width           =   3105
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
         Left            =   3270
         TabIndex        =   5
         Top             =   1095
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Текущее хранилище"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1515
         Width           =   2220
      End
      Begin Threed.SSCommand ssEdit 
         Height          =   375
         Left            =   2490
         TabIndex        =   7
         ToolTipText     =   "Редактировать ГОСТы"
         Top             =   240
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   131074
         PictureFrames   =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmFactors.frx":0442
         Caption         =   "Редактировать"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSOption ssMainStg 
         Height          =   300
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   131074
         Caption         =   "Основное"
         Value           =   -1
      End
      Begin Threed.SSOption ssExtraStg 
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   750
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   131074
         Caption         =   "Дополнительное"
      End
      Begin Threed.SSCommand ssApplyPath 
         Height          =   360
         Left            =   3675
         TabIndex        =   6
         ToolTipText     =   "Применить"
         Top             =   1065
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmFactors.frx":0554
         AutoSize        =   2
         ButtonStyle     =   3
      End
   End
   Begin VB.ListBox lstFactors 
      BackColor       =   &H00FFFBF0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3030
      IntegralHeight  =   0   'False
      ItemData        =   "frmFactors.frx":06BE
      Left            =   135
      List            =   "frmFactors.frx":06C5
      TabIndex        =   0
      Top             =   165
      Width           =   5550
   End
   Begin MSComDlg.CommonDialog cdPath 
      Left            =   0
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "htg"
      Filter          =   "Файлы хранилищ ГОСТов (*.htg)|*.htg|Все файлы (*.*)|*.*"
   End
   Begin Threed.SSCommand ssClose 
      Cancel          =   -1  'True
      Height          =   675
      Left            =   4410
      TabIndex        =   8
      Top             =   4185
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1191
      _Version        =   131074
      PictureFrames   =   1
      BackStyle       =   1
      PictureUseMask  =   -1  'True
      Picture         =   "frmFactors.frx":06D5
      Caption         =   "Закрыть"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PostCmds
  cmdCreFac
  cmdEditFac
  cmdRemoveFac
  cmdRenameFac
End Enum

Public FirstStart As Integer

Private m_b_Res As Boolean
Private m_sFileName As String
Private m_b_LockResize As Boolean
Private m_sPathToGostsStg As String
Private m_sPathToMainGostsStg As String
Private m_collFactors As CollFactors
Private m_iLockedM As Integer

Private m_vEditIdx As Variant

Private Const T_PAD = 100
Private Const L_PAD = 100
Private Const T_PAD2 = 100
Private Const L_PAD2 = 100

Private m_iStg As IStorage

Private Sub cmdPath_Click()
  With cdPath
    .CancelError = True
    .DialogTitle = "Выбрать хранилище Нормативов"
    .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
    If .InitDir = "" Then _
      .InitDir = App.Path
                
    .FileName = ""
                
    On Error GoTo ERR_H0
    .ShowOpen
    DoEvents
    
    On Error GoTo 0
                
    txtPath.Text = .FileName
    
    Dim sDir As String
    CutPath .FileName, sDir
    .InitDir = sDir
    Exit Sub
ERR_H0:
  End With
End Sub

Private Sub Form_Initialize()
  m_b_LockResize = False
  
  With GTMsgHook1
    .Windows.Item(1).Messages.Add _
      Key:="WM_PENDING_OP", Value:=WM_PENDING_OP, ProcessDefaultProc:=False
  
    .Subclass = True
  End With
  
End Sub

Private Sub Form_Resize()
  If m_b_LockResize Then Exit Sub
  
  On Error Resume Next
  m_b_LockResize = True
  
  If Me.WindowState <> vbMinimized Then
    If Width < 5000 Then Width = 5000
    If Height < 3000 Then Height = 3000
  Else
    m_b_LockResize = False
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  
  lstFactors.Move L_PAD, T_PAD, ScaleWidth - 2 * L_PAD, ScaleHeight - 3 * T_PAD - ssfGOST.Height
  ssfGOST.Move lstFactors.Left, lstFactors.Top + lstFactors.Height + T_PAD, lstFactors.Width * (3 / 4)
  ssClose.Move ssfGOST.Left + ssfGOST.Width + L_PAD, ssfGOST.Top + (ssfGOST.Height - ssClose.Height) / 2, lstFactors.Width - ssfGOST.Width - L_PAD
  ssEdit.Move ssfGOST.Width - L_PAD2 - ssEdit.Width
  ssApplyPath.Move ssfGOST.Width - L_PAD2 - ssApplyPath.Width
  cmdPath.Move ssApplyPath.Left - cmdPath.Width - ScaleX(1, vbPixels, vbTwips)
  txtPath.Width = ssfGOST.Width - (ssfGOST.Width - cmdPath.Left) - txtPath.Left - ScaleX(1, vbPixels, vbTwips)
  txtCurrStg.Width = ssfGOST.Width - txtCurrStg.Left - (ssfGOST.Width - ssApplyPath.Left - ssApplyPath.Width)
  
  m_b_LockResize = False
  Exit Sub
ERR_H:
  m_b_LockResize = False
  Err.Raise Err.Number
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    ssClose_Click
  ElseIf KeyAscii = vbKeyReturn Then
    'ssOk_Click
  End If
End Sub

Private Sub Form_Load()
  m_b_Res = False
  m_iLockedM = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    txtInPlace_LostFocus
    ssClose_Click
  End If
End Sub


Private Sub Form_Terminate()
  GTMsgHook1.Subclass = False
End Sub

Private Sub GTMsgHook1_Message(ByVal hWnd_ As Long, ByVal MsgID As Long, ByVal wParam As Long, ByVal lParam As Long)
  If hWnd_ <> Me.hWnd Then Exit Sub
  If MsgID <> WM_PENDING_OP Then Exit Sub
    
  On Error GoTo ERR_H
  
  Select Case lParam
    Case cmdCreFac
      CmCreFac
    Case cmdEditFac
      CmEditFac
    Case cmdRemoveFac
      CmRemoveFac
    Case cmdRenameFac
      CmRenameFac
  End Select
  
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub lstFactors_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyReturn Then
        
  ElseIf KeyCode = vbKeyDelete Then
    KeyCode = 0
    CmRemoveFac
    
  ElseIf KeyCode = vbKeyF2 Then
    KeyCode = 0
    CmRenameFac
    
  ElseIf KeyCode = vbKeyF3 Then
    KeyCode = 0
    CmEditFac
    
  End If
End Sub

Private Sub lstFactors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
  If m_iLockedM = 1 Or Button <> 2 Then Exit Sub
  
  With frmShadow
          
    Dim lX As Long, lY As Long, lSnd As Long
    lX = ScaleX(x, vbTwips, vbPixels)
    lY = ScaleY(y, vbTwips, vbPixels)
    
          
'    On Error Resume Next
'    m_iLockedM = 1
'    SendMessage lstFactors.hWnd, WM_LBUTTONDOWN, 2, lSnd
'    SendMessage lstFactors.hWnd, WM_LBUTTONUP, 2, lSnd
'    m_iLockedM = 0
'    On Error GoTo 0
    Dim pt As POINTAPI: pt.x = lX: pt.y = lY
    With lstFactors
      'ScreenToClient .hWnd, pt

      lSnd = MakeLong(pt.x, pt.y)
      lSnd = LoLong(SendMessageLong(.hWnd, LB_ITEMFROMPOINT, 0, lSnd))

      If lSnd >= 0 Then _
        .Selected(lSnd) = True
                
    End With
    
    If lstFactors.ListCount < 1 Or lSnd < 0 Then
      .mnuCreFac.Enabled = True
      .mnuRenameFac.Enabled = False
      .mnuEditFac.Enabled = False
      .mnuRemoveFac.Enabled = False
    Else
      .mnuCreFac.Enabled = True
      .mnuRenameFac.Enabled = True
      .mnuEditFac.Enabled = True
      .mnuRemoveFac.Enabled = True
    End If
    
    PopupMenu .mnuFacEdit, vbPopupMenuLeftAlign, x, y
    
  End With
End Sub

Private Sub ssClose_Click()
  m_b_Res = False
  
  Dim ips As IPersistStorage
  Dim ic As IStCollItem
  Set ic = m_collFactors
  If ic.Status <> OS_None Then
    Set ips = m_collFactors
    ips.Save Nothing, 1
  End If
  Me.Hide
  
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssClose_LostFocus()
  HighLight2 ssClose, False, Me.hWnd
End Sub

Private Sub ssClose_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, True, Me.hWnd
End Sub

Private Sub ssClose_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssClose, False, Me.hWnd
End Sub


Private Sub ssEdit_Click()
  On Error GoTo ERR_H
  
  Dim vt As Variant
  With frmGostEdit
    .KeyTopic = vt
    .KeyGOST = vt
    .PathToGostsStg = m_sPathToGostsStg
    If .FirstStart = 0 Then CenterForm frmGostEdit, Me
    .AssData
    .Show vbModal, Me
    DoEvents
    .UnassData
  End With

  Exit Sub
ERR_H:
  frmGostEdit.UnassData
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub ssEdit_LostFocus()
  HighLight2 ssEdit, False, Me.hWnd
End Sub

Private Sub ssEdit_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssEdit, True, Me.hWnd
End Sub

Private Sub ssEdit_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssEdit, False, Me.hWnd
End Sub

Public Property Get ModalResult() As Boolean
  ModalResult = m_b_Res
End Property


Friend Sub AssData(ByVal sFileName As String, ByVal sPathToGostsStg As String, ByVal bFlMain As Boolean, ByVal sPathToMainGostsStg As String)
  FirstStart = 1
  
  m_b_Res = False
  If bFlMain Then
    ssMainStg.Value = True
    txtPath.Text = ""
    txtPath.Enabled = False
    cmdPath.Enabled = False
    ssApplyPath.Enabled = False
  Else
    ssExtraStg.Value = True
    txtPath.Text = sPathToGostsStg
    txtPath.Enabled = True
    cmdPath.Enabled = True
    ssApplyPath.Enabled = True
  End If
  
  m_sFileName = sFileName
  m_sPathToGostsStg = sPathToGostsStg
  m_sPathToMainGostsStg = sPathToMainGostsStg
  txtCurrStg.Text = sPathToGostsStg
  
  Dim bm As CBeam: Set bm = New CBeam
  bm.Beam Me
  
  Set m_iStg = OpenCF(sFileName, STGM_DIRECT1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
    
    
  Set m_collFactors = New CollFactors
  Dim ips As IPersistStorage
  Dim ci As IStCollItem
  Set ci = m_collFactors: Set ips = m_collFactors
  ci.DefaultCU = False
  m_collFactors.SetUpdateMode Nothing, False
  
  If IsEmptyStg(m_iStg) Then
    ips.InitNew m_iStg
    ips.Save Nothing, 1
  Else
    ips.Load m_iStg
  End If
  
  LoadToList
  
End Sub

Private Sub LoadToList()
  With lstFactors
    .Clear
    Dim lCnt As Long: lCnt = m_collFactors.Count
    Dim l As Long
    For l = 1 To lCnt
      .AddItem m_collFactors.NameNth(l)
      .ItemData(.NewIndex) = m_collFactors.KeyNth(l)
    Next l
    
    If .ListCount > 0 Then .Selected(0) = True
  End With
End Sub

Friend Sub UnassData()
  Set m_iStg = Nothing
  Set m_collFactors = Nothing
End Sub


Private Sub ssExtraStg_Click(Value As Integer)
  txtPath.Enabled = True
  cmdPath.Enabled = True
  ssApplyPath.Enabled = True
End Sub

Private Sub StoreStgPath()
  Dim istm As IStream
  Set istm = CreateStreamInt(m_iStg, StreamStgGostsName, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  PutString istm, txtPath.Text
End Sub

Private Sub ssApplyPath_Click()

  txtPath.Text = Trim(txtPath.Text)
  
  If StrComp(m_sPathToMainGostsStg, txtPath.Text, vbTextCompare) = 0 Then
    MsgBox "Выбранный файл принадлежит основному хранилищу", vbExclamation Or vbOKOnly, "Нельзя назначить"
    Exit Sub
  End If
  
  On Error Resume Next
  FileLen txtPath.Text
  If Err.Number <> 0 Then
    Dim vbRes As VbMsgBoxResult
    vbRes = MsgBox("Файл '" & txtPath.Text & "' не найден. Создать новое хранилище ?", vbExclamation Or vbYesNo, "Ошибка")
    DoEvents
    If vbRes = vbYes Then
      Dim iStg As IStorage
      On Error GoTo ERR_H
      Set iStg = CreateCF(txtPath.Text, STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
      Set iStg = Nothing
    Else
      txtPath.SetFocus
    End If
    Exit Sub
  End If
  
  On Error GoTo ERR_H
  StoreStgPath
  m_sPathToGostsStg = txtPath.Text
  txtCurrStg.Text = txtPath.Text
  ssApplyPath.Enabled = False
  Exit Sub
  
ERR_H:
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub


Private Sub ssMainStg_Click(Value As Integer)
  txtPath.Enabled = False
  cmdPath.Enabled = False
  ssApplyPath.Enabled = False
  
  m_sPathToGostsStg = m_sPathToMainGostsStg
  txtCurrStg.Text = m_sPathToGostsStg
       
  On Error GoTo ERR_H
  DestroyCF m_iStg, StreamStgGostsName
  Exit Sub
  
ERR_H:
  If Err.Number <> &H80030002 Then
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
    Err.Clear
  End If
End Sub

Private Sub txtPath_Change()
  ssApplyPath.Enabled = True
End Sub


Friend Sub PostCreFac()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdCreFac
End Sub
Friend Sub PostEditFac()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdEditFac
End Sub
Friend Sub PostRemoveFac()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdRemoveFac
End Sub
Friend Sub PostRenameFac()
  PostMessage hWnd, WM_PENDING_OP, 0, cmdRenameFac
End Sub

Private Sub CmCreFac()
  Dim isci As IStCollItem
  Set isci = m_collFactors.Add("<Новый>")
  With lstFactors
    .AddItem "<Новый>"
    .ItemData(.NewIndex) = isci.Key
    .Selected(.NewIndex) = True
    CmRenameFac
  End With
End Sub
    
Private Sub CmEditFac()
  On Error GoTo ERR_H
  Dim fAss1 As Boolean, fAss2 As Boolean, vt As Variant
  fAss1 = False: fAss2 = False
    
  With frmEditFT
    If .FirstStart = 0 Then CenterForm frmEditFT, Me
    .Caption = lstFactors.List(lstFactors.ListIndex)
    
    With frmGostEdit
      .PathToGostsStg = m_sPathToGostsStg
      .KeyTopic = vt
      .KeyGOST = vt
      If .FirstStart = 0 Then CenterForm frmGostEdit, Me
      .AssData
    End With
    
    fAss1 = True
    .AssData m_collFactors(lstFactors.ItemData(lstFactors.ListIndex))
    fAss2 = True
    .Show vbModal, Me
    DoEvents
    .UnassData
    frmGostEdit.UnassData
  End With

  Exit Sub
ERR_H:
  If fAss2 Then frmEditFT.UnassData
  If fAss1 Then frmGostEdit.UnassData
  MsgBox Err.Description, vbCritical Or vbOKOnly, "Ошибка"
  Err.Clear
End Sub

Private Sub CmRemoveFac()
  With lstFactors
    If .ListIndex = -1 Then Exit Sub
    m_collFactors.Remove .ItemData(.ListIndex)
    Dim lK As Long: lK = .ListIndex
    .RemoveItem .ListIndex
    
    On Error Resume Next
    If lK < .ListCount Then
      .Selected(lK) = True
    Else
      .Selected(lK - 1) = True
    End If
  End With
End Sub

Private Sub CmRenameFac()
  With lstFactors
    If .ListIndex = -1 Then Exit Sub
    
    m_vEditIdx = .ListIndex
    Dim r As RECT, pt1 As POINTAPI, pt2 As POINTAPI
    SendMessage .hWnd, LB_GETITEMRECT, .ListIndex, r
    
    pt1.x = r.Left:  pt1.y = r.Top
    MapWindowPoints .hWnd, hWnd, pt1, 1
    pt2.x = r.Right:  pt2.y = r.Bottom
    MapWindowPoints .hWnd, hWnd, pt2, 1
    
    Dim x As Single, y As Single, w As Single, h As Single
    
        
    x = ScaleX(pt1.x, vbPixels, vbTwips)
    w = ScaleX(pt2.x - pt1.x, vbPixels, vbTwips)
    y = ScaleY(pt1.y, vbPixels, vbTwips)
    h = ScaleY(pt2.y - pt1.y, vbPixels, vbTwips)
              
    KeyPreview = False
    ssClose.Cancel = False
      
    With txtInPlace
      .Move x, y, w, h
            
      .Text = lstFactors.List(lstFactors.ListIndex)
      .Visible = True
      .ZOrder 0
      .SetFocus
    End With
  End With
  
  Exit Sub
ERR_H:
  m_vEditIdx = Null
  txtInPlace.Visible = False
  KeyPreview = True
  ssClose.Cancel = True
  
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear

End Sub

Private Sub txtInPlace_KeyPress(KeyAscii As Integer)
  If IsNull(m_vEditIdx) Then Exit Sub
  
  On Error GoTo ERR_H
  
  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    GoTo L_UNLOCK
    
  ElseIf KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Dim szText As String: szText = Trim(txtInPlace.Text)
    If Len(szText) < 1 Then
      MsgBox "Нельзя использовать пустое имя", vbExclamation Or vbOKOnly, "Ошибка"
      GoTo L_UNLOCK
    End If
    
    Dim ic As IStCollItem
    
    m_collFactors.NameKey(lstFactors.ItemData(m_vEditIdx)) = szText
    lstFactors.List(m_vEditIdx) = szText
    
    GoTo L_UNLOCK
            
  End If
  
  Exit Sub
  
L_UNLOCK:
  On Error GoTo 0
  Me.KeyPreview = True
  ssClose.Cancel = True
  m_vEditIdx = Null
  txtInPlace.Visible = False
  lstFactors.SetFocus
  Exit Sub
ERR_H:
  MsgBox Err.Description, vbExclamation Or vbOKOnly, "Ошибка"
  Err.Clear
  GoTo L_UNLOCK
End Sub

Private Sub txtInPlace_LostFocus()

  If Not IsNull(m_vEditIdx) Then
    Me.KeyPreview = True
    ssClose.Cancel = True
    m_vEditIdx = Null
    txtInPlace.Visible = False
    lstFactors.SetFocus
  End If
End Sub

