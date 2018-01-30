VERSION 5.00
Begin VB.Form frmPrtShadow 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmPrtShadow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuRpt 
      Caption         =   "Отчёт"
      Begin VB.Menu mnuPrint 
         Caption         =   "Печать"
         Begin VB.Menu PrtPrint 
            Caption         =   "Печатать"
         End
         Begin VB.Menu s1 
            Caption         =   "-"
         End
         Begin VB.Menu PrtRemovePages 
            Caption         =   "Удалить все страницы"
         End
         Begin VB.Menu PrtRemReports 
            Caption         =   "Удалить отчёты..."
         End
         Begin VB.Menu s3 
            Caption         =   "-"
         End
         Begin VB.Menu PrtCopy 
            Caption         =   "Копировать страницу в буфер обмена"
         End
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Масштаб"
         Begin VB.Menu Zoom100 
            Caption         =   "100%"
         End
         Begin VB.Menu ZoomEntirePage 
            Caption         =   "Вся страница"
         End
         Begin VB.Menu ZoomWidth 
            Caption         =   "Ширина страницы"
         End
         Begin VB.Menu s2 
            Caption         =   "-"
         End
         Begin VB.Menu ZoomZoom 
            Caption         =   "Распахнуть окно отчётов"
         End
         Begin VB.Menu s5 
            Caption         =   "-"
         End
         Begin VB.Menu ZoomPage 
            Caption         =   "Страница"
         End
         Begin VB.Menu ZoomScale 
            Caption         =   "Масштаб"
         End
      End
   End
End
Attribute VB_Name = "frmPrtShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
  Unload frmPrint
  Unload frmScalePage
End Sub

Private Sub PrtCopy_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  prt.Action = paCopyPage
End Sub


Private Sub PrtPrint_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  'Load frmPrint
  frmPrint.Caption = "Печать (" & prt.ToolTipText & ")"
  frmPrint.InitData
  frmPrint.Show vbModal, frxRep
  
  DoEvents
  
  If prt Is frxRep.vsModel Then
    UpdateGUI_Rep1 haApp
  Else
    UpdateGUI_Rep2 haApp
  End If
End Sub

Private Sub PrtRemovePages_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  prt.StartDoc
  prt.EndDoc
  frxRep.UpdateControlBar
  If frxRep.tabPrint.CurrTab = 0 Then
    If haApp.Rep1 Is Nothing Then
      MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
      Exit Sub
    End If
    haApp.Rep1.Remove Start:=1, Count:=haApp.Rep1.Count
  Else
    If haApp.Rep2 Is Nothing Then
      MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
      Exit Sub
    End If
    haApp.Rep2.Remove Start:=1, Count:=haApp.Rep2.Count
  End If
End Sub

Private Sub PrtRemReports_Click()
  
  If frxRep.tabPrint.CurrTab = 0 Then
    If haApp.Rep1 Is Nothing Then
      MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
      Exit Sub
    End If
    frmRepEditor.Assign frxRep.vsModel, haApp.Rep1
    frmRepEditor.Caption = "Удаление отчётов по модели"
  Else
    If haApp.Rep2 Is Nothing Then
      MsgBox "Нет документа", vbInformation Or vbOKOnly, "Предупреждение"
      Exit Sub
    End If
    frmRepEditor.Assign frxRep.vpSP, haApp.Rep2
    frmRepEditor.Caption = "Удаление отчётов по комплексу мер"
  End If
  
  On Error GoTo ERR_H
  frmRepEditor.Show vbModal, frmMain
  frmRepEditor.Unassign
  
  If frmRepEditor.ModalResult Then
    If frxRep.tabPrint.CurrTab = 0 Then
      UpdateGUI_Rep1 haApp
    Else
      UpdateGUI_Rep2 haApp
    End If
  End If
  
  frxRep.UpdateControlBar
  
  Exit Sub
ERR_H:
   frmRepEditor.Unassign
   Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Zoom100_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  prt.Zoom = 100
  frxRep.UpdateControlBar
End Sub

Private Sub ZoomEntirePage_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  Dim dW As Double, dH As Double, dK As Double
  dW = 100# / frxRep.tabPrint.ClientWidth * prt.PageWidth
  dH = 100# / frxRep.tabPrint.ClientHeight * prt.PageHeight
  dK = IIf(dW > dH, dW, dH)
  prt.Zoom = 1# / dK * 10000# * 0.95
  
  frxRep.UpdateControlBar
End Sub

Private Sub ZoomPage_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  
  With frmScalePage
    .PVNumeric1.ValueMin = IIf(prt.PageCount = 0, 0, 1)
    .PVNumeric1.ValueMax = prt.PageCount
    .PVNumeric1.ValueInteger = prt.PreviewPage
    .Caption = "Страница " & CStr(.PVNumeric1.ValueMin) & " - " & .PVNumeric1.ValueMax
    .ModalResult = False
    .Show vbModal, frmMain
    If .ModalResult Then
      prt.PreviewPage = .PVNumeric1.ValueInteger
      frxRep.UpdateControlBar
    End If
  End With
End Sub

Private Sub ZoomScale_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  
  With frmScalePage
    .PVNumeric1.ValueMin = prt.ZoomMin
    .PVNumeric1.ValueMax = prt.ZoomMax
    .PVNumeric1.ValueReal = prt.Zoom
    .Caption = "Масштаб " & CStr(prt.ZoomMin) & "% - " & prt.ZoomMax & "%"
    .ModalResult = False
    .Show vbModal, frmMain
    If .ModalResult Then
      prt.Zoom = .PVNumeric1.ValueReal
      frxRep.UpdateControlBar
    End If
  End With
End Sub

Private Sub ZoomWidth_Click()
  Dim prt As VSPrinter
  Set prt = frxRep.CurrPrinter()
  Dim dW As Double
  dW = 100# / frxRep.tabPrint.ClientWidth * prt.PageWidth
  prt.Zoom = 1# / dW * 10000# * 0.95
  
  frxRep.UpdateControlBar
End Sub

Private Sub ZoomZoom_Click()
  Dim ll As Long
  ll = GetSystemMetrics(SM_CYFULLSCREEN)
  If frxRep.MeX.FrameType = otxFloating Then
    frxRep.Move 0, 0, ScaleX(GetSystemMetrics(SM_CXFULLSCREEN), vbPixels, vbTwips), ScaleY(GetSystemMetrics(SM_CYFULLSCREEN), vbPixels, vbTwips)
  Else
    frxRep.MeX.Float 0, 0, ScaleX(GetSystemMetrics(SM_CXFULLSCREEN), vbPixels, vbTwips), ScaleY(GetSystemMetrics(SM_CYFULLSCREEN), vbPixels, vbTwips)
  End If
End Sub
