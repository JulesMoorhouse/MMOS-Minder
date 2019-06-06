VERSION 5.00
Begin VB.Form frmMakeIt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please enter your unlock code ..."
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMCL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration Details"
      Height          =   1815
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   4935
      Begin VB.Label lblName 
         Caption         =   "JULIAN MOORHOUSE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblCompanyTelephoneNum 
         Caption         =   "0123 456789"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblCompanyName 
         Caption         =   "MINDWARP CONSULTANCY LTD"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdUnlock 
      Caption         =   "Unlock"
      Height          =   360
      Left            =   2760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1305
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   3
      Left            =   4260
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   2
      Left            =   3480
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   1
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   1
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5100
      Top             =   4560
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You will automatically be moved to the next text box, while entering your code."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   3360
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your ""Unlock Code"" provided by Mindwarp Consultany Ltd."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome! Thank you, for purchasing Mindwarp Mail Order System!"
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
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lblBlastOff 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   3615
   End
End
Attribute VB_Name = "frmMakeIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lintCountDown As Integer
Dim lbooJustSelected As Boolean

Private Sub cmdUnlock_Click()
Dim lvarErrorStage

    If Len(Trim$(txt(0))) < 5 Or Len(Trim$(txt(1))) < 5 Or _
        Len(Trim$(txt(2))) < 5 Or Len(Trim$(txt(3))) < 5 Then
        MsgBox "You have entered an invalid code!", , gconstrTitlPrefix & "Startup"
        Unload Me
        Exit Sub
    End If
    
    With gstrKey
        .strUnlockKey = txt(0) & txt(1) & txt(2) & txt(3)
    End With
    
    If Decode(Trim$(lblCompanyName), Trim$(lblCompanyTelephoneNum), Trim$(lblName)) <> "21" Then
        MsgBox "You have entered an invalid code!", , gconstrTitlPrefix & "Startup"
        Unload Me
    Else
        gstrUserMode = gconstrLiveMode
        'ModeChange Me, gstrUserMode
        With gstrStatic
            If gstrSystemRoute = srCompanyRoute Then
                lvarErrorStage = 110
                Set gdatCentralDatabase = OpenDatabase(.strTrueLiveServerPath & .strShortCentralDBFile, , False)
            Else
                lvarErrorStage = 130
                Set gdatCentralDatabase = OpenDatabase(.strTrueLiveServerPath & .strShortCentralDBFile, _
                    dbDriverComplete, False, Trim$(gstrDBPasswords.strCentralDBPasswordString))
            End If
            
            DeployStaticInfo
            
            .strUnlockCode = gstrKey.strUnlockKey
            Encrypt gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, gconEncryptStatic
            .strUnlockCode = ""
            MsgBox "Unlock Code Accepted!" & vbCrLf & vbCrLf & "This program will now close, the program will then be ready for use!", , gconstrTitlPrefix & "Unlock"
            
            UpdateLoader
            Unload Me
            Unhook
            End
        End With
    End If

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    lintCountDown = 60
    
    With gstrReferenceInfo
        lblCompanyName = UCase$(Trim$(.strCompanyName))
        lblCompanyTelephoneNum = UCase$(Trim$(.strCompanyTelephone))
        lblName = UCase$(Trim$(.strCompanyContact))
    End With

    frmButtons.ImageList16Cols.UseMaskColor = True
    frmButtons.ImageList16Cols.MaskColor = vbRed 'vbGreen
    frmButtons.ImageList16Cols.BackColor = picMCL.BackColor
    frmButtons.ImageList16Cols.ListImages(11).Draw picMCL.hdc, 0, 0   ', imlNormal 'imlTransparent   'imlNormal

End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Timer1_Timer()

    lintCountDown = lintCountDown - 1
    If lintCountDown = 0 Then Unload Me
    lblBlastOff = lintCountDown & " Seconds left to enter your code!"
    
End Sub

Private Sub txt_GotFocus(Index As Integer)

    lbooJustSelected = True
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii <> Asc(vbBack) Then
        If lbooJustSelected = False Then
            If Len(txt(Index)) = txt(Index).MaxLength Then
                On Error Resume Next
                txt(Index + 1).SetFocus
                txt(Index + 1) = Chr(KeyAscii)
                txt(Index + 1).SelStart = Len(txt(Index + 1))
            End If
        Else
            lbooJustSelected = False
        End If
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    lbooJustSelected = False
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetSelected Me
    
End Sub
