VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Menu"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMaxMin 
      Caption         =   "&Maximize"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraSummary 
      Caption         =   "Summary Information"
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame fraKey 
         Caption         =   "Key"
         Height          =   615
         Left            =   3720
         TabIndex        =   13
         Top             =   1920
         Width           =   2655
         Begin VB.CommandButton cmdInfo 
            Caption         =   "&Info"
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bad"
            Height          =   255
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Good"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Shape Shape5 
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   735
         End
         Begin VB.Shape Shape4 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   960
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label lblSummaryItem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Summary Item"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   270
         Width           =   3855
      End
      Begin VB.Shape shpSummaryItem 
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdQA 
      Caption         =   "O&rder Maintenence"
      Enabled         =   0   'False
      Height          =   492
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2172
   End
   Begin VB.CommandButton cmdConsignments 
      Caption         =   "&Parcelforce Consignments"
      Height          =   492
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2172
   End
   Begin VB.CommandButton cmdFinance 
      Caption         =   "&Cash Book"
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   2172
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton cmdAccountMaint 
      Caption         =   "&Account Maintenence"
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   2172
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdOrderEnq 
      Caption         =   "Order &Enquiry"
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2172
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "&Log Out && Exit"
      Height          =   492
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   2172
   End
   Begin VB.CommandButton cmdMinder 
      Caption         =   "&Minder"
      Height          =   492
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   2172
   End
   Begin VB.CommandButton cmdPackers 
      Caption         =   "&Packing"
      Height          =   492
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2172
   End
   Begin VB.CommandButton cmdOrderEntry 
      Caption         =   "&Order Entry"
      Height          =   492
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2172
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   6720
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "26/09/01"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:44"
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   2895
      Left            =   2520
      Top             =   960
      Width           =   3855
   End
   Begin VB.Image Image3 
      Height          =   2880
      Left            =   2520
      Picture         =   "Main.frx":014A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   120
      Top             =   840
      Width           =   8655
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Company Mail Order System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmMainX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbooIfSummaryInfoBuilt As Boolean
Sub ShowSummaryInfo(pbooWithBuild As Boolean, Optional pstrLastUpdated As String)
Dim lintObjects As Integer
Dim lintArrInc As Integer
Dim llngLeftPos As Long
Const lconlngWidth = 4270

    fraSummary.Visible = True
    If Me.Height < 7365 And Me.WindowState <> vbMaximized Then
        Me.Height = 7365
    End If
    llngLeftPos = 120
    
    If pbooWithBuild = True Then
        lintObjects = BuildSummaryItems(lblSummaryItem, shpSummaryItem, "CLIENT", 3800, Me)
    Else
        lintObjects = lblSummaryItem.Count
    End If
    
    For lintArrInc = 1 To lintObjects - 1
        shpSummaryItem(lintArrInc).Top = lblSummaryItem(lintArrInc - 1).Top + lblSummaryItem(lintArrInc - 1).Height + 40
        lblSummaryItem(lintArrInc).Top = shpSummaryItem(lintArrInc).Top + 25
                
        If shpSummaryItem(lintArrInc).Top + shpSummaryItem(lintArrInc).Height > (fraSummary.Height - fraKey.Height) Then
            lblSummaryItem(lintArrInc).Top = shpSummaryItem(0).Top + 30 '+ 25
            shpSummaryItem(lintArrInc).Top = shpSummaryItem(0).Top + 5
            llngLeftPos = llngLeftPos + lconlngWidth
        End If
        
        lblSummaryItem(lintArrInc).Left = llngLeftPos + 120
        shpSummaryItem(lintArrInc).Left = llngLeftPos
        
    Next lintArrInc

End Sub

Private Sub cmdAccountMaint_Click()

    Unload Me
    frmCustAcctSel.Route = gconstrAccount
    frmCustAcctSel.Show
        
End Sub

Private Sub cmdConsignments_Click()

    Unload Me
    frmPForce.Route = gconstrConsignmentNorm
    frmPForce.CallingForm = frmAbout
    frmPForce.Show
    
End Sub

Private Sub cmdFinance_Click()

    Unload Me
    frmCheque.Route = gconstrFinance
    frmCheque.Show
    
End Sub

Private Sub cmdHelp_Click()

    RunNWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1001.htm"
    
End Sub

Private Sub cmdInfo_Click()

    MsgBox "When items appear in Red (Bad) or Green (Good) this relates to " & vbCrLf & _
           "comparitive calculations performed on other known statistics." & vbCrLf & vbCrLf & _
           "The status is assigned when in 10% of the comparitive statistic." & vbCrLf & vbCrLf & _
           "Please move your mouse over a statistic for more details." & vbCrLf & vbCrLf & _
           "The current statistics were last updated X", _
            vbInformation, gconstrTitlPrefix & "Summary Information"
End Sub

Private Sub cmdLogout_Click()
Dim lintRetVal As Integer
   
    lintRetVal = MsgBox("You are about to logout and close the system!", _
        vbYesNo + vbDefaultButton1 + vbInformation, gconstrTitlPrefix & "System Exit")
    
    If lintRetVal = vbNo Then
        Exit Sub
    End If
    Busy True, Me
    gdatCentralDatabase.Close
    gdatLocalDatabase.Close
    Set gdatLocalDatabase = Nothing
    Set gdatCentralDatabase = Nothing
    
    UpdateLoader
    Busy False, Me
    Unload Me
    End
    
End Sub

Private Sub cmdMaxMin_Click()

    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
        cmdMaxMin.Caption = "&Maximize"
    Else
        Me.WindowState = vbMaximized
        cmdMaxMin.Caption = "&Normal"
    End If
    
End Sub

Private Sub cmdMinder_Click()
Dim lstrSourcePath As String
Dim lstrDestinationPath As String
Dim lstrSourceFile As String
Dim lstrDestinationFile As String
Dim lbooCopyDone As Boolean
Dim lintRetVal As Variant

    lintRetVal = MsgBox("Would you like to run Scandisk and Defrag?", vbYesNo)
    If lintRetVal = vbYes Then

        gdatLocalDatabase.Close
        gdatCentralDatabase.Close
        Set gdatLocalDatabase = Nothing
        Set gdatCentralDatabase = Nothing
        
        lstrSourcePath = gstrStatic.strServerPath '& "\"
        lstrDestinationPath = AppPath
        
        lstrSourceFile = lstrSourcePath & "Minder.exe"
        lstrDestinationFile = lstrDestinationPath & "Minder.exe"
        lbooCopyDone = FileCopyIfNewer(lstrSourceFile, lstrDestinationFile)
        
        'Run Minder.exe with APP Parameter

        RunNWait lstrDestinationFile & " APP"
    End If
    
    End

End Sub

Private Sub cmdOrderEnq_Click()

    Unload Me
    frmCustAcctSel.Route = gconstrEnquiry
    frmCustAcctSel.Show

End Sub

Private Sub cmdOrderEntry_Click()

    Unload Me
    frmCustAcctSel.Route = gconstrEntry
    frmCustAcctSel.Show
    
End Sub

Private Sub cmdPackers_Click()

    Unload Me
    frmPackaging.Route = gconstrPacking
    frmPackaging.Show
    
End Sub
Private Sub cmdQA_Click()

    Unload Me
    frmQAMisc.Route = gconstrOrdMaint
    frmQAMisc.Show
    
End Sub

Private Sub Form_Load()
Dim lstrSummaryString As String
Dim lstrSummaryLastUpdated As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    lblProductName = gconstrProductFullName
    
    NameForm Me
    ShowStatus 13

    Select Case gstrGenSysInfo.lngUserLevel
    Case Is < 20 'Distribution
        cmdOrderEntry.Enabled = False
        cmdOrderEnq.Enabled = False
        cmdAccountMaint.Enabled = False
        cmdPackers.Enabled = True
        cmdFinance.Enabled = False
        cmdConsignments.Enabled = False
        cmdQA.Enabled = False
    Case Is < 30 'Order Entry
        cmdOrderEntry.Enabled = True
        cmdOrderEnq.Enabled = True
        cmdAccountMaint.Enabled = True
        cmdPackers.Enabled = False
        cmdFinance.Enabled = False
        cmdConsignments.Enabled = True
        cmdQA.Enabled = False
    Case Is < 40 'Sales
        cmdOrderEntry.Enabled = False
        cmdOrderEnq.Enabled = True
        cmdAccountMaint.Enabled = True
        cmdPackers.Enabled = False
        cmdFinance.Enabled = False
        cmdConsignments.Enabled = False
        cmdQA.Enabled = False
    Case Is < 50 'Accounts
        cmdOrderEntry.Enabled = False
        cmdOrderEnq.Enabled = True
        cmdAccountMaint.Enabled = True
        cmdPackers.Enabled = False
        cmdFinance.Enabled = True
        cmdConsignments.Enabled = False
        cmdQA.Enabled = False
    Case Is < 99 ' General Managers
        cmdOrderEntry.Enabled = True
        cmdOrderEnq.Enabled = True
        cmdAccountMaint.Enabled = True
        cmdPackers.Enabled = True
        cmdFinance.Enabled = True
        cmdConsignments.Enabled = True
        cmdQA.Enabled = True
    Case Is < 100 ' Information Systems
        cmdOrderEntry.Enabled = True
        cmdOrderEnq.Enabled = True
        cmdAccountMaint.Enabled = True
        cmdPackers.Enabled = True
        cmdFinance.Enabled = True
        cmdConsignments.Enabled = True
        cmdQA.Enabled = True
    End Select

    'normal
    Me.Height = 4905
    If DebugVersion Then
        PopulateSummaryArray
        lstrSummaryLastUpdated = GetSummaryString(lstrSummaryString)
        If lstrSummaryString <> "" Then
            AddValuestoSummaryArray lstrSummaryString & Chr(182)
            ShowSummaryInfo True, lstrSummaryLastUpdated
            lbooIfSummaryInfoBuilt = True
            DebugSummary
        End If
    End If

End Sub

Private Sub Form_Resize()
Dim llngFormHalfWidth As Long

    llngFormHalfWidth = Me.Width / 2
    
    With lblProductName
        .Left = 0
        .Width = Me.Width
    End With
    
    With Shape2
        .Left = (llngFormHalfWidth - (Shape2.Width / 2)) - 60
        'Debug.Print "Shape2.Left = " & .Left
    End With
    
    With Shape1
        .Width = (Me.Width - (240)) - 120
    End With
    
    With fraSummary
        .Width = Shape1.Width
        .Height = ((Me.Height - .Top) - 700)
    End With
    
    With Image3
        .Left = (llngFormHalfWidth - (Image3.Width / 2)) - 60
        'Debug.Print "Image3.Left = " & .Left
    End With
    
    With cmdOrderEntry
        '.Left = 240
        .Width = (Shape2.Left - 120) - .Left
    End With
    
    With cmdOrderEnq
        .Width = cmdOrderEntry.Width
    End With
    
    With cmdAccountMaint
        .Width = cmdOrderEntry.Width
    End With
    
    With cmdFinance
        .Width = cmdOrderEntry.Width
    End With
    
    With cmdHelp
        .Width = cmdOrderEntry.Width
    End With
    
    With cmdPackers
        .Left = Shape2.Left + Shape2.Width + 120
        .Width = cmdOrderEntry.Width
    End With
    
    With cmdConsignments
        .Left = cmdPackers.Left
        .Width = cmdPackers.Width
    End With
    
    With cmdQA
        .Left = cmdPackers.Left
        .Width = cmdPackers.Width
    End With
    
    With cmdMinder
        .Left = cmdPackers.Left
        .Width = cmdPackers.Width
    End With
    
    With cmdLogout
        .Left = cmdPackers.Left
        .Width = cmdPackers.Width
    End With

    With fraKey
        .Left = (fraSummary.Width / 2) - (fraKey.Width / 2)
        .Top = fraSummary.Height - fraKey.Height
    End With
    
    If DebugVersion = True And lbooIfSummaryInfoBuilt = True Then
        ShowSummaryInfo False
    End If

End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

