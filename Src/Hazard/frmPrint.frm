VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#6.0#0"; "pvnum.ocx"
Begin VB.Form frmPrint 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Печать"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   ClipControls    =   0   'False
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Height          =   2250
      Left            =   -45
      TabIndex        =   0
      Top             =   -180
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   3969
      _Version        =   131074
      BackColor       =   12632256
      BackStyle       =   1
      PictureBackgroundStyle=   2
      BorderWidth     =   0
      BevelOuter      =   0
      Begin Threed.SSFrame SSFrame1 
         Height          =   1110
         Left            =   195
         TabIndex        =   1
         Top             =   255
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1958
         _Version        =   131074
         BackStyle       =   1
         Caption         =   "Страницы"
         ClipControls    =   0   'False
         Begin PVNumericLib.PVNumeric edEnd 
            Height          =   300
            Left            =   2580
            TabIndex        =   3
            Top             =   645
            Width           =   810
            _Version        =   393216
            _ExtentX        =   1429
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "1"
            ForeColor       =   -2147483640
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            ValueMin        =   1
            ValueReal       =   1
         End
         Begin PVNumericLib.PVNumeric edSta 
            Height          =   300
            Left            =   1395
            TabIndex        =   2
            Top             =   645
            Width           =   810
            _Version        =   393216
            _ExtentX        =   1429
            _ExtentY        =   529
            _StockProps     =   253
            Text            =   "1"
            ForeColor       =   -2147483640
            BackColor       =   16776176
            Appearance      =   1
            BackColor       =   16776176
            DecimalSeparator=   ","
            LimitValue      =   -1  'True
            ValueMin        =   1
            ValueReal       =   1
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   240
            Left            =   2310
            TabIndex        =   4
            Top             =   675
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   423
            _Version        =   131074
            BackStyle       =   1
            Caption         =   "до:"
            BevelOuter      =   0
            AutoSize        =   1
         End
         Begin Threed.SSOption optInterval 
            Height          =   255
            Left            =   195
            TabIndex        =   6
            Top             =   675
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   131074
            BackStyle       =   1
            Caption         =   "Интервал"
         End
         Begin Threed.SSOption optAll 
            Height          =   255
            Left            =   195
            TabIndex        =   5
            Top             =   300
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   450
            _Version        =   131074
            BackStyle       =   1
            Caption         =   "Все"
            Value           =   -1
         End
      End
      Begin Threed.SSCommand ssPageSetup 
         Height          =   690
         Left            =   3990
         TabIndex        =   9
         Top             =   1410
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1217
         _Version        =   131074
         PictureMaskColor=   65280
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPrint.frx":000C
         Caption         =   "Страница"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssPrint 
         Height          =   615
         Left            =   180
         TabIndex        =   8
         Top             =   1485
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         _Version        =   131074
         PictureMaskColor=   65280
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPrint.frx":0182
         Caption         =   "Начать"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand ssPrintSetup 
         Height          =   690
         Left            =   3990
         TabIndex        =   7
         Top             =   375
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1217
         _Version        =   131074
         PictureMaskColor=   65280
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "frmPrint.frx":02F8
         Caption         =   "Принтер"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_b_Act As Boolean



Friend Sub InitData()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter
  
  edSta.ValueMax = prt.PageCount
  edEnd.ValueMax = prt.PageCount
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then _
     PostMessage hWnd, WM_CLOSE, 0, 0
End Sub

Private Sub Form_Load()
  Set SSPanel1.PictureBackground = LoadResPicture(101, vbResBitmap)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then _
    Cancel = 1: Me.Hide
End Sub

Private Sub Form_Resize()
  'SSPanel1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub ssPageSetup_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  prt.PrintDialog pdPageSetup
  'prt.PageHeight
  frxRep.UpdateControlBar
End Sub

Private Sub ssPageSetup_LostFocus()
  HighLight2 ssPageSetup, False, Me.hWnd
End Sub

Private Sub ssPageSetup_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPageSetup, True, Me.hWnd
End Sub

Private Sub ssPageSetup_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPageSetup, False, Me.hWnd
End Sub

Private Sub ssPrint_Click()
  If edSta.ValueInteger > edEnd.ValueInteger Then
    MsgBox "Начальная страница должна быть <= конечной", vbExclamation Or vbOKOnly, "Предупреждение"
    Exit Sub
  End If
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter
  With prt
    .AbortCaption = prt.ToolTipText
    .AbortTextButton = "Отменить печать"
    .AbortTextDevice = "%s на %s"
    .AbortTextPage = "Сейчас печатается страница %d из"
    .AbortWindow = True
    .AbortWindowPos = awAppWindow
    
    If optAll.Value = True Then
      .PrintDoc False, 1, .PageCount
    Else
      .PrintDoc False, edSta.ValueInteger, edEnd.ValueInteger
    End If
  End With
End Sub

Private Sub ssPrint_LostFocus()
  HighLight2 ssPrint, False, Me.hWnd
End Sub

Private Sub ssPrint_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPrint, True, Me.hWnd
End Sub

Private Sub ssPrint_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPrint, False, Me.hWnd
End Sub

Private Sub ssPrintSetup_Click()
  Dim prt As VSPrinter, ii1 As Integer, ii2 As Integer
  Set prt = frxRep.CurrPrinter
  
  ii1 = prt.PrintQuality
  prt.Action = paChoosePrinter
  
  'prt.PrintDialog pdPrinterSetup
  ii2 = prt.PrintQuality
  frxRep.UpdateControlBar
End Sub

Private Sub ssPrintSetup_LostFocus()
  HighLight2 ssPrintSetup, False, Me.hWnd
End Sub

Private Sub ssPrintSetup_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPrintSetup, True, Me.hWnd
End Sub

Private Sub ssPrintSetup_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  HighLight2 ssPrintSetup, False, Me.hWnd
End Sub

