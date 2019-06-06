VERSION 5.00
Begin VB.Form frmChildUserPass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Access Details"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRePassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1305
   End
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   4065
   End
   Begin VB.TextBox txtOtherPassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2160
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   288
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2325
   End
   Begin VB.Label lblRePassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Retype Password:"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Tag             =   "&Password:"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblEnterUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter a user name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblEnterFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your full name!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "For security reason, you password is not shown below"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblChangePassword 
      BackStyle       =   0  'Transparent
      Caption         =   "To Change password enter below"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Full Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblOtherPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Retype Password:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Tag             =   "&Password:"
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Tag             =   "&User Name:"
      Top             =   375
      Width           =   1080
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Tag             =   "&Password:"
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "frmChildUserPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrRoute As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute

End Property
Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    Select Case mstrRoute
    Case "ADD"
    Case "PASSCHANGE"
        With gstrGenSysInfo
            If Trim$(.strUserPassword) <> Hash(Trim$(txtPassword)) Then
                MsgBox "This is not your current password! Please try again!", vbInformation, gconstrTitlPrefix & "User Access Details"
                txtPassword.SetFocus
                txtPassword.SelStart = 0
                txtPassword.SelLength = Len(txtPassword.Text)
                Exit Sub
            End If
            If Trim$(txtOtherPassword) <> Trim$(txtRePassword) Then
                MsgBox "Your new password does not match the re-typed password! Please try again!", vbInformation, gconstrTitlPrefix & "User Access Details"
                txtOtherPassword.SetFocus
                txtOtherPassword.SelStart = 0
                txtOtherPassword.SelLength = Len(txtOtherPassword.Text)
                Exit Sub
            End If
            If Len(Trim$(txtOtherPassword)) < 6 Then ' will also cater for blank!
                MsgBox "Your password is too short, it must be at least 6 characters in length!", vbInformation, gconstrTitlPrefix & "User Access Details"
                txtOtherPassword.SetFocus
                txtOtherPassword.SelStart = 0
                txtOtherPassword.SelLength = Len(txtOtherPassword.Text)
                Exit Sub
            End If
            UpdateUser .strUserName, .strUserFullName, txtOtherPassword, .lngUserLevel, .strUserNotes
            MsgBox "Your Password has been updated!", vbInformation, gconstrTitlPrefix & "User Access Details"
        End With
    End Select
    
    Unload Me
    
End Sub

Private Sub Form_Activate()

    Select Case mstrRoute
    Case "PASSCHANGE"
        txtPassword.SetFocus
    End Select
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    Select Case mstrRoute
    Case "ADD"
        lblEnterUserName.Visible = True
        lblEnterFullName.Visible = True
        lblChangePassword.Visible = False
        lblNote.Visible = False
        txtRePassword.Visible = False
        lblRePassword.Visible = False
    Case "PASSCHANGE"
        With txtUserName
            .Enabled = False
            .BackColor = vbButtonFace
            .Text = Trim$(gstrGenSysInfo.strUserName)
        End With
        lblEnterUserName.Visible = False
        With txtFullName
            .Enabled = False
            .BackColor = vbButtonFace
            .Text = Trim$(gstrGenSysInfo.strUserFullName)
        End With
        lblEnterFullName.Visible = False
        lblChangePassword.Visible = True
        lblNote.Visible = True
        lblPassword.Caption = "&Current Password:"
        lblOtherPassword.Caption = "&New Password:"
        txtRePassword.Visible = True
        lblRePassword.Visible = True

    End Select
    
End Sub


Private Sub txtFullName_GotFocus()

    SetSelected Me
    
End Sub
Private Sub txtOtherPassword_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtPassword_GotFocus()

    SetSelected Me
    
End Sub
Private Sub txtRePassword_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtUserName_GotFocus()

    SetSelected Me
    
End Sub
