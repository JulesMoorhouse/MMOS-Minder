VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1875
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1425
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   288
      Left            =   1425
      MaxLength       =   20
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   1605
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3625
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "05/06/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "18:45"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   5
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public OK As Boolean

Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long
    
    NameForm Me, True

    sbStatusBar.Panels(1).Text = GreetingTime
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = ProperCase$(Left$(sBuffer, lSize))
    Else
        txtUserName.Text = vbNullString
    End If
    
End Sub
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    
    Dim lvarRetVal
    Dim lintUserCount As Long
    
    With gstrGenSysInfo
        .strUserName = ProperCase$(Trim$(txtUserName))
        
        If GetUser(Trim$(.strUserName), lintUserCount, True) = False Then
            lvarRetVal = MsgBox("Your UserID has not been found!" & vbCrLf & _
                "Would you like to add this new userID to the system?", vbYesNo + vbInformation, gconstrTitlPrefix & "User Login")
            If lvarRetVal = vbYes Then
                'by default set to Order Entry

                If Trim$(txtPassword.Text & "") = "" Then
                    MsgBox "Please enter a password you'd like to use with your new account!", , gconstrTitlPrefix & "User Login"
                    Exit Sub
                End If
                
                Dim lintUserLevel As Long: lintUserLevel = 20
                Dim lstrUserLevel As String: lstrUserLevel = "Order Entry level"
                
                If lintUserCount = 0 Then
                    lintUserLevel = 99
                    lstrUserLevel = "Information Systems level (as you're the first user!)"
                End If
                
                AddNewUser .strUserName, "My Full Name", lintUserLevel, txtPassword.Text, " ", True
                GetUser Trim$(.strUserName), lintUserCount, True
                
                MsgBox "Your new UserID has been added!" & vbCrLf & vbCrLf & _
                    "Your user level has been set to " & vbCrLf & _
                    lstrUserLevel & ".", vbInformation, gconstrTitlPrefix & "User Login"
            Else
                Exit Sub
            End If
        End If

        If Hash(Trim$(txtPassword.Text)) = Trim$(.strUserPassword) Then
            OK = True
            Me.Hide
            
        Else
            MsgBox "Invalid Password, try again!", , gconstrTitlPrefix & "User Login"
            txtPassword.SetFocus
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
        End If
    End With
End Sub

Private Sub txtPassword_GotFocus()
    
    SetSelected Me
    
End Sub

Private Sub txtUserName_GotFocus()

    SetSelected Me
    
End Sub
