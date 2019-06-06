VERSION 5.00
Begin VB.Form frmAbout 
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFeatures 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2460
      Left            =   2985
      TabIndex        =   5
      Top             =   5575
      Width           =   7555
      Begin VB.ListBox lstNewFeatures 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   2205
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   7550
      End
      Begin VB.CheckBox chkAllProgs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         Caption         =   "All programs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdFeatClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7290
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   20
         Width           =   260
      End
      Begin VB.Label lblNewFeatures 
         BackColor       =   &H80000002&
         Caption         =   " New Features: (Click for more information)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   7555
      End
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1852
   End
   Begin VB.Label lblConfigure 
      Alignment       =   2  'Center
      Caption         =   "You must complete all three steps, shown on the left!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   7455
   End
   Begin VB.Label lblTrainingCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Getting Started"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      MouseIcon       =   "About.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Click me to make contact"
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      X1              =   1005
      X2              =   1005
      Y1              =   2520
      Y2              =   3360
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   12
      X1              =   480
      X2              =   1560
      Y1              =   2160
      Y2              =   2760
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      X1              =   2040
      X2              =   2040
      Y1              =   2400
      Y2              =   3360
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      X1              =   1080
      X2              =   1920
      Y1              =   2400
      Y2              =   1800
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      X1              =   2520
      X2              =   2520
      Y1              =   2040
      Y2              =   3000
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      X1              =   1560
      X2              =   2520
      Y1              =   3720
      Y2              =   3000
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      X1              =   480
      X2              =   480
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      X1              =   480
      X2              =   1440
      Y1              =   2160
      Y2              =   1440
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      X1              =   1560
      X2              =   1560
      Y1              =   2760
      Y2              =   3720
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      X1              =   1560
      X2              =   2520
      Y1              =   2760
      Y2              =   2040
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      X1              =   960
      X2              =   2040
      Y1              =   1800
      Y2              =   2400
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      X1              =   1440
      X2              =   2520
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      X1              =   480
      X2              =   1560
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Shape shpBacking 
      BorderColor     =   &H00800000&
      BorderWidth     =   10
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   6795
      Left            =   105
      Top             =   1125
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "program or its sister programs, please email :-"
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   3960
      Width           =   7545
   End
   Begin VB.Label lblCover 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your cover has expired!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   7545
   End
   Begin VB.Label lblMCLContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "web address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4080
      MouseIcon       =   "About.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Click me to make contact"
      Top             =   4320
      Width           =   5475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If you have any comments or suggestions about this"
      Height          =   255
      Left            =   3060
      TabIndex        =   1
      Top             =   3600
      Width           =   7545
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lvarOrigBackcolor As Variant
Dim llngNewFeatures() As Long
Private Sub cmdFeatClose_Click()

    fraFeatures.Visible = False
    
End Sub

Private Sub Form_Activate()

    If gstrTempKeyFail <> "" Then
        MsgBox "Please be advised that your temporary license will expire on " & gstrTempKeyFail & " after this date, " & vbCrLf & _
            "this software will no longer function! For continued usage beyond this date please " & vbCrLf & _
            "ensure that you have purchased a full license. " & vbCrLf & vbCrLf & _
            "If you would like to discuss this matter, please Contact Mindwarp Consultancy Ltd.", vbInformation, gconstrTitlPrefix & "Warning!"
        gstrTempKeyFail = ""
    End If

    If WindowLoaded("Getting Started") Then
        glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, gstrHelpFileBase & "::/GettingStarted/GettingStarted.htm>GetStart", HH_DISPLAY_TOPIC, 0)
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (Shift And vbKeyControl) > 0 Then
        Select Case (KeyCode)
        Case vbKeyO ' Order Entry
            mdiMain.MenuCommands mnuClientGoOrderEntry
        Case vbKeyE ' Order Enquiry
            mdiMain.MenuCommands mnuClientGoEnquiry
        Case vbKeyA ' Account Maintenance
            mdiMain.MenuCommands mnuClientGoAcctMaint
        Case vbKeyP ' Packing
            mdiMain.MenuCommands mnuClientGoPacking
        Case vbKeyF ' Finance
            mdiMain.MenuCommands mnuClientGoFinance
        Case vbKeyM ' Order Maintenance
            mdiMain.MenuCommands mnuClientGoOrderMaint
        End Select
    End If
    
End Sub

Private Sub Form_Load()
Dim lstrShowFeatures As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    lvarOrigBackcolor = lblMCLContact.ForeColor
        
    lblMCLContact.Caption = gstrOurContactWeb
    
    ShowBanner frmAbout, ""
    gstrButtonRoute = gconstrMainMenu
    mdiMain.DrawButtonSet gstrButtonRoute
    
    If UCase$(App.ProductName) <> "LITE" Then
        lblTrainingCard.Visible = False
        lstrShowFeatures = GetSetting(gstrIniAppName, "UI", "ShowFeatures")
        If IsBlank(lstrShowFeatures) Then
            lstrShowFeatures = True
        ElseIf UCase$(lstrShowFeatures) <> "TRUE" Or UCase$(lstrShowFeatures) <> "FALSE" Then
            lstrShowFeatures = True
        End If

        If CBool(lstrShowFeatures) = True Then
            fraFeatures.Visible = True
        Else
            lstrShowFeatures = False
            fraFeatures.Visible = False
            SaveSetting gstrIniAppName, "UI", "ShowFeatures", lstrShowFeatures
        End If
        
        PopFeatList lstNewFeatures, False, llngNewFeatures()
    Else
        fraFeatures.Visible = False
    End If
    
    If UCase$(App.ProductName) <> "CONFIGURE" Then
        lblConfigure.Visible = False
    End If
    
    If gdatCoverDate < Date And gdatCoverDate <> "00:00:00" Then
        lblCover = "Your cover has expired! " & gdatCoverDate & " " & gstrStatic.strUnlockCode
        lblCover.Visible = True
    End If
    
    App.HelpFile = gstrHelpFileBase
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMCLContact.ForeColor = lvarOrigBackcolor
    lblTrainingCard.ForeColor = lvarOrigBackcolor

End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub
Public Sub Form_Resize()

    On Error Resume Next
    
    MakeVisible Me, False
    
    With lblCover
        .Width = Me.Width - 3060
    End With
    
    With Label2
        .Width = Me.Width - 3060
    End With
    
    With Label3
        .Width = Me.Width - 3060
    End With
    
    With lblConfigure
        .Width = Me.Width - 3060
    End With
    
    With shpBacking
        .Left = Me.Left + 160
        .Height = (Me.Height - (0 + 1080)) - 180
    End With
    
    With fraFeatures
        '.Top = ((Me.Height - 705) - 1755) - 100
        .Top = ((Me.Height - 705) - 1755) - 100
        .Left = shpBacking.Width + shpBacking.Left + 70 '50
        .Width = (Me.Width - .Left) - 110
    End With

    lstNewFeatures.Width = fraFeatures.Width
    lblNewFeatures.Width = fraFeatures.Width
    chkAllProgs.Left = fraFeatures.Width - 1795
    cmdFeatClose.Left = fraFeatures.Width - 265
    lblNewFeatures.BackColor = vbActiveTitleBar
    chkAllProgs.BackColor = vbActiveTitleBar
    
    With lblTrainingCard
        .Left = fraFeatures.Left + ((fraFeatures.Width / 2) - (.Width / 2))
    End With
    
    With lblMCLContact
        .Left = fraFeatures.Left + ((fraFeatures.Width / 2) - (.Width / 2))
    End With
    
    MakeVisible Me, True
    
End Sub
Private Sub lblMCLContact_Click()
Dim StartDoc As Long

     StartDoc = ShellExecute(Me.hwnd, "open", gstrOurContactWeb, _
       "", "C:\", 1)

End Sub
Private Sub lblMCLContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMCLContact.ForeColor = vbRed '&HFF8080

End Sub

Private Sub lblTrainingCard_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, gstrHelpFileBase & "::/GettingStarted/GettingStarted.htm>GetStart", HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub lblTrainingCard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblTrainingCard.ForeColor = vbRed

End Sub
Private Sub lstNewFeatures_Click()

    FeatureMsg (llngNewFeatures(lstNewFeatures.ListIndex))
    
End Sub
Private Sub chkAllProgs_Click()

    If chkAllProgs.Value = 0 Then ' unchecked
        PopFeatList lstNewFeatures, False, llngNewFeatures()
    Else
        PopFeatList lstNewFeatures, True, llngNewFeatures()
    End If
    
End Sub
