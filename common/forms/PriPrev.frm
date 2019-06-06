VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Preview"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   Icon            =   "PriPrev.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   12330
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   540
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   1000
      Left            =   1800
      SmallChange     =   100
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      LargeChange     =   1000
      Left            =   6600
      SmallChange     =   100
      TabIndex        =   9
      Top             =   2400
      Width           =   255
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2040
      ScaleHeight     =   2745
      ScaleWidth      =   3825
      TabIndex        =   8
      Top             =   1800
      Width           =   3855
      Begin VB.PictureBox picCanvas 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   3375
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2505
         ScaleWidth      =   3585
         TabIndex        =   17
         Top             =   120
         Width           =   3615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "P&revious"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "&Goto"
         Height          =   375
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPrinterSetup 
         Caption         =   "&Setup"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   6960
         TabIndex        =   14
         Top             =   0
         Width           =   1575
         Begin VB.TextBox txtPageNum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "1"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblPages 
            BackStyle       =   0  'Transparent
            Caption         =   "/   200"
            Height          =   255
            Left            =   840
            TabIndex        =   16
            Top             =   160
            Width           =   615
         End
      End
      Begin VB.ComboBox cboZoom 
         Height          =   315
         ItemData        =   "PriPrev.frx":000C
         Left            =   3360
         List            =   "PriPrev.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   5775
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16563
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "04/08/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:58"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const llngSpacer = 500
Dim llngToolbarHeight As Long
Dim llngStatusHeight As Long
Dim llngPageNumber As Long
Sub SetPageControls()

    lblPages = "/   " & gstrReport.intPagesInReport

    If llngPageNumber >= gstrReport.intPagesInReport Then
        cmdNext.Enabled = False
        cmdPrevious.Enabled = True
    Else
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
    End If
    
    If llngPageNumber = 1 Then
        cmdPrevious.Enabled = False
    End If
    
End Sub
Sub RefreshDisplay(pobjObject As Object, Optional pintScale As Variant)
Dim lintKeptScale As Variant

    If IsMissing(pintScale) Then
        pintScale = gintScaleFactor
    End If
    
    lintKeptScale = gintScaleFactor
    
    picCanvas.Visible = False
    picPaper.Visible = False
    DoEvents
    
    gintScaleFactor = pintScale
    
    PrintNPreview llngPageNumber, pobjObject
    
    gintScaleFactor = lintKeptScale
    
    Call Form_Resize

    picCanvas.Visible = True
    picPaper.Visible = True
    
    SetPageControls
    
    ' Position the horizontal scroll bar.
    HScroll1.Top = picBackground.Height + llngToolbarHeight
    HScroll1.Left = 0
    HScroll1.Width = picBackground.Width
    
    ' Position the vertical scroll bar.
    VScroll1.Top = 0
    VScroll1.Left = picBackground.Width
    VScroll1.Height = picBackground.Height + llngToolbarHeight
    
    ' Set the Max value for the scroll bars.
    HScroll1.Max = (picCanvas.Width - picBackground.Width) + llngSpacer
    VScroll1.Max = (picCanvas.Height - picBackground.Height) + (llngToolbarHeight * 2)
    HScroll1.Min = -llngSpacer
    VScroll1.Min = -llngSpacer
    
    HScroll1.Value = HScroll1.Min
    VScroll1.Value = VScroll1.Min
    
    DoEvents
    
End Sub
Private Sub cboZoom_Click()
Dim lsngScale As Single
    picCanvas.Cls

    Select Case cboZoom
    Case "300%"
        gintScaleFactor = rpScale300
    Case "200%"
        gintScaleFactor = rpScale200
    Case "150%"
        gintScaleFactor = rpScale150
    Case "100%"
        gintScaleFactor = rpScale100
    Case "87.5%"
        gintScaleFactor = rpScale87_5
    Case "75%"
        gintScaleFactor = rpScale75
    Case "60%"
        gintScaleFactor = rpScale60
    Case "50%"
        gintScaleFactor = rpScale50
    End Select
   
    RefreshDisplay picCanvas
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdNext_Click()

    txtPageNum = txtPageNum + 1
    llngPageNumber = txtPageNum
    RefreshDisplay picCanvas
    
    SetPageControls

    DoEvents
        
End Sub

Private Sub cmdOptions_Click()

    frmReportOptions.Show vbModal
    RefreshDisplay picCanvas
    
End Sub

Private Sub cmdPage_Click()

    If Val(txtPageNum) <= gstrReport.intPagesInReport Then
        llngPageNumber = txtPageNum
    Else
        txtPageNum = gstrReport.intPagesInReport
        llngPageNumber = txtPageNum
    End If
    
    SetPageControls
    RefreshDisplay picCanvas
    
End Sub

Private Sub cmdPrevious_Click()

    If Val(txtPageNum) > 1 Then
        txtPageNum = txtPageNum - 1
        llngPageNumber = txtPageNum
        RefreshDisplay picCanvas
    End If
    
    SetPageControls
    
End Sub

Private Sub cmdPrint_Click()
Dim llngNumOfCopies As Long
Dim lintArrInc As Long
Dim lintArrInc2 As Long

    On Error Resume Next
    Printer.TrackDefault = False

'CommonDialog1.ShowPrinter

    With CommonDialog1
        .DialogTitle = "Print"
        .CancelError = True
        .Min = 1
        .Max = gstrReport.intPagesInReport
        .Flags = cdlPDPageNums + cdlPDHidePrintToFile + cdlPDNoSelection + cdlPDUseDevModeCopies
        
        .FromPage = 1
        .ToPage = gstrReport.intPagesInReport

        .ShowPrinter
        
        If Err.Number = 32755 Then
            Exit Sub
        End If
        On Error GoTo 0
                
        For lintArrInc2 = 1 To .Copies
            For lintArrInc = .FromPage To .ToPage
                If lintArrInc > .FromPage Then
                    Printer.NewPage
                End If
                PrintNPreview lintArrInc, Printer
            Next lintArrInc
        Next lintArrInc2
        
    End With
    
    Printer.EndDoc
    
End Sub

Private Sub cmdPrinterSetup_Click()

 Printer.TrackDefault = False

'CommonDialog1.ShowPrinter
    On Error GoTo Err_Handler
    With CommonDialog1
        .DialogTitle = "Print Setup"
        .CancelError = True
        .Flags = cdlPDPrintSetup ' Or cdlPDNoWarning Or cdlPDReturnIC Or cdlPDReturnDC
        .ShowPrinter
    End With
    Printer.NewPage
    Printer.KillDoc
    RefreshDisplay picCanvas
    Exit Sub
Err_Handler:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    Case Else
        Resume Next
    End Select
    
End Sub

Sub Form_Load()
Dim lstrCurrentMode As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
                       
    lstrCurrentMode = gstrUserMode
    gstrUserMode = ""
    NameForm Me
    gstrUserMode = lstrCurrentMode

    Me.Caption = gconstrTitlPrefix & "Print Preview"
    
    lblPages = "/   " & gstrReport.intPagesInReport

    If gstrReport.booShowPageSetup = False Then
        cmdPrinterSetup.Visible = False
    End If
    If gstrReport.booShowOptions = False Then
        cmdOptions.Visible = False
    End If
    
    If gstrReport.booHideZoom = True Then
        cboZoom.Visible = False
    End If
    
    llngPageNumber = 1
    
    cboZoom.Text = "100%"
    llngToolbarHeight = Toolbar1.Height
    llngStatusHeight = sbStatusBar.Height
    picBackground.Move 0, llngToolbarHeight, ScaleWidth - VScroll1.Width, ((ScaleHeight - HScroll1.Height) - llngStatusHeight) - llngToolbarHeight
    picCanvas.Move llngSpacer, llngSpacer
    CalcPrintableArea
    picPaper.Move picCanvas.Left - gstrPMarg.lngNonPrintableLeftMargin, picCanvas.Top - gstrPMarg.lngNonPrintableTopMargin
    
    ' Position the horizontal scroll bar.
    HScroll1.Top = picBackground.Height + llngToolbarHeight
    HScroll1.Left = 0
    HScroll1.Width = picBackground.Width
    
    ' Position the vertical scroll bar.
    VScroll1.Top = 0
    VScroll1.Left = picBackground.Width
    VScroll1.Height = picBackground.Height + llngToolbarHeight
    
    ' Set the Max value for the scroll bars.
    HScroll1.Max = (picCanvas.Width - picBackground.Width) + llngSpacer
    VScroll1.Max = (picCanvas.Height - picBackground.Height) + (llngToolbarHeight * 2)
    HScroll1.Min = -llngSpacer
    VScroll1.Min = -llngSpacer
    
    SetPageControls
    
End Sub

Sub HScroll1_Change()

    picCanvas.Left = -HScroll1.Value
    picPaper.Left = picCanvas.Left - (gstrPMarg.lngNonPrintableLeftMargin) / gintScaleFactor

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Debug.Print "X=" & X & " Y=" & Y
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sbStatusBar.Panels(1).Text = "X=" & X & " Y=" & Y

End Sub

Private Sub picCanvas_Resize()

    picPaper.Height = (Printer.Height + glngPageAdjustHeight) / gintScaleFactor
    picPaper.Width = (Printer.Width + glngPageAdjustWidth) / gintScaleFactor

End Sub

Private Sub txtPageNum_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtPageNum_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        llngPageNumber = txtPageNum
        RefreshDisplay picCanvas
        
        SetPageControls
    End If
    
End Sub

Private Sub txtPageNum_KeyPress(KeyAscii As Integer)

    If InStr("0123456789", Chr(KeyAscii)) <> 0 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub txtPageNum_LostFocus()

    txtPageNum = Val(txtPageNum)
    
End Sub

Sub VScroll1_Change()
  
    picCanvas.Top = -VScroll1.Value
    picPaper.Top = picCanvas.Top - (gstrPMarg.lngNonPrintableTopMargin) / gintScaleFactor
    
End Sub

Sub Form_Resize()
    ' When the form size is changed, the picBackground dimensions are changed
    ' to match.
    picBackground.Height = Me.Height
    
    picBackground.Width = Me.Width
    ' Re-Initializes picture postitions & scroll bars.
    
    picBackground.Move 0, llngToolbarHeight, ScaleWidth - VScroll1.Width, ((ScaleHeight - HScroll1.Height) - llngStatusHeight) - llngToolbarHeight
    
    picCanvas.Move llngSpacer, llngSpacer
    picPaper.Move picCanvas.Left - (gstrPMarg.lngNonPrintableLeftMargin) / gintScaleFactor, _
        picCanvas.Top - (gstrPMarg.lngNonPrintableTopMargin) / gintScaleFactor
    
    HScroll1.Top = picBackground.Height + llngToolbarHeight
    HScroll1.Left = 0
    HScroll1.Width = picBackground.Width
    VScroll1.Top = 0
    VScroll1.Left = picBackground.Width
    VScroll1.Height = picBackground.Height + llngToolbarHeight
    HScroll1.Max = (picCanvas.Width - picBackground.Width) + llngSpacer
    VScroll1.Max = (picCanvas.Height - picBackground.Height) + (llngToolbarHeight * 2)
    HScroll1.Min = -llngSpacer
    VScroll1.Min = -llngSpacer
    
End Sub
