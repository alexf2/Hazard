VERSION 5.00
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Begin VB.Form SWTstForm 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   390
      Left            =   4095
      TabIndex        =   4
      Top             =   4815
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   1530
      TabIndex        =   3
      Top             =   4905
      Width           =   1335
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4395
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   7752
      _Version        =   131074
      PaneTree        =   "SWTstForm.frx":0000
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   30
         TabIndex        =   2
         Top             =   1530
         Width           =   5850
      End
      Begin VB.TextBox Text1 
         Height          =   1410
         Left            =   30
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   5850
      End
   End
End
Attribute VB_Name = "SWTstForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  SSSplitter1.Panes(1).Control = Nothing
  List1.Visible = False
  SSSplitter1.Panes.Remove "Pane B"
  
End Sub

Private Sub Command2_Click()
  SSSplitter1.Visible = False
  'List1.Visible = False
  'Text1.Visible = False
  SSSplitter1.Panes(0).Control = Nothing
  SSSplitter1.Panes.Add "Pane A", ssBottomOfSplit, "Pane B"
  SSSplitter1.Panes(0).Control = Text1
  SSSplitter1.Panes(1).Control = List1
  'List1.Visible = True
  'Text1.Visible = True
  SSSplitter1.Visible = True
  List1.Visible = True
End Sub

Private Sub Form_Load()
Exit Sub
  CreTstFac
  Exit Sub

  Dim ct As New AWPLUGINLib.CollTopics
  Dim cIt As AWPLUGINLib.IStCollItem
  Set cIt = ct
  cIt.DefaultCU = True
  Dim ips As IPersistStorage
  Set ips = ct
  ips.InitNew Nothing
  
  Dim cllGosts As AWPLUGINLib.ICollGosts
  Set cllGosts = ct.Add("ГОСТЫ 1")
  
  Dim gt As IGostTable
  Set gt = cllGosts.Add("ГОСТ 1.1")
  With gt
    .ClmName = "Параметр ГОСТа 1.1"
    .InsertSlots 5
    
    .NDescr(0) = "Кол-во вещества от 1 до 1"
    .NLingvoVal(0) = "Лингво 1"
    .NComment(0) = "Коментарий 1"
    
    .NDescr(1) = "Кол-во вещества от 2 до 2"
    .NLingvoVal(1) = "Лингво 2"
    .NComment(1) = "Коментарий 2"
    
    .NDescr(2) = "Кол-во вещества от 3 до 3"
    .NLingvoVal(2) = "Лингво 3"
    .NComment(2) = "Коментарий 3"
    
    .NDescr(3) = "Кол-во вещества от 4 до 4"
    .NLingvoVal(3) = "Лингво 4"
    .NComment(3) = "Коментарий 4"
    
    .NDescr(4) = "Кол-во вещества от 5 до 5"
    .NLingvoVal(4) = "Лингво 5"
    .NComment(4) = "Коментарий 5"
    
    .SetUniformGrad
  End With
  
  Set gt = cllGosts.Add("ГОСТ 1.2")
  With gt
    .ClmName = "Параметр ГОСТа 1.2"
    .InsertSlots 2
    
    .NDescr(0) = "Кол-во вещества от 1 до 1"
    .NLingvoVal(0) = "Лингво 1"
    .NComment(0) = "Коментарий 1"
    
    .NDescr(1) = "Кол-во вещества от 2 до 2"
    .NLingvoVal(1) = "Лингво 2"
    .NComment(1) = "Коментарий 2"
    
    .SetUniformGrad
  End With
  
  Dim st As IStorage
  Set st = CreateCF("g:\cg.stg", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ips.Save st, 0
  
End Sub
  
Private Sub CreTstFac()
  Dim f As New CollFactors
  Dim ips As IPersistStorage
  Dim st As IStCollItem
  Set st = f
  Set ips = f
  st.DefaultCU = True
  ips.InitNew Nothing
  
  Dim cft As CollFTables
  Set cft = f.Add("M01")
  
  Dim ift As IFactorTable
  Set ift = cft.Add("Factor 1")
    
  With ift
    .InsertSlots 7
    .NComponentName(0) = "Slot 1"
    .NAssSlotType 0, FT_None
    
    .NComponentName(1) = "Slot 2"
    .NAssSlotType 1, FT_Gost, 0, 2
    
    .NComponentName(2) = "Slot 3"
    .NAssSlotType 2, FT_None
    
    .NComponentName(3) = "Slot 4"
    .NAssSlotType 3, FT_Gost, 1, 1
    
    .NComponentName(4) = "Slot 5"
    .NAssSlotType 4, FT_Self, 1
    
    .NComponentName(5) = "Slot 6"
    .NAssSlotType 5, FT_Self, 2
    
    .NComponentName(6) = "Slot 7"
    .NAssSlotType 6, FT_Gost, 1, 0
  End With
  
  Set ift = cft.Add("Factor 2")
  With ift
    .InsertSlots 2
    .NComponentName(0) = "Slot 1"
    .NAssSlotType 0, FT_None
    
    .NComponentName(1) = "Slot 2"
    .NAssSlotType 1, FT_Gost, 0, 0
  End With
  
  Set ift = cft.Add("Factor 3")
  With ift
    .InsertSlots 3
    .NComponentName(0) = "Slot 1"
    .NAssSlotType 0, FT_None
    
    .NComponentName(1) = "Slot 2"
    .NAssSlotType 1, FT_Gost, 0, 2
    
    .NComponentName(2) = "Slot 3"
    .NAssSlotType 2, FT_Self, 3
  End With
  
  Set ift = cft.Add("Factor 4")
  With ift
    .InsertSlots 4
    .NComponentName(0) = "Slot 1"
    .NAssSlotType 0, FT_None
    
    .NComponentName(1) = "Slot 2"
    .NAssSlotType 1, FT_Gost, 0, 2
    
    .NComponentName(2) = "Slot 3"
    .NAssSlotType 2, FT_Gost, 0, 2
    
    .NComponentName(3) = "Slot 4"
    .NAssSlotType 3, FT_Gost, 0, 2
    
  End With
  
  Dim stg As IStorage
  Set stg = CreateCF("g:\cg.stg", STGM_DIRECT1 Or STGM_CREATE1 Or STGM_READWRITE1 Or STGM_SHARE_EXCLUSIVE1)
  ips.Save stg, 0
End Sub
