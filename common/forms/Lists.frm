VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lists"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMaxMin 
      Caption         =   "&Maximize"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   6720
      Top             =   6240
   End
   Begin VB.CommandButton cmdPFSet 
      Caption         =   "&Reset PF"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeploy 
      Caption         =   "&Deploy"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   492
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1332
   End
   Begin MSDBGrid.DBGrid dbgListDetails 
      Bindings        =   "Lists.frx":0000
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "Lists.frx":001D
      TabIndex        =   1
      Top             =   3120
      Width           =   9855
   End
   Begin VB.Data datListDetails 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ListDetailsMaster"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgLists 
      Bindings        =   "Lists.frx":126B
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "Lists.frx":1282
      TabIndex        =   0
      Top             =   600
      Width           =   9855
   End
   Begin VB.Data datLists 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from ListsMaster where sysuse <> true"
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   492
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   6750
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12568
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "25/03/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "16:42"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "If your at all unsure about what to do on this screen please read the HELP first!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbooSysUser As Boolean

Private Sub cmdClose_Click()

    Unload Me
    frmMain.Show
    
End Sub

Private Sub cmdDeploy_Click()
Dim lintRetVal As Integer
Dim lvarErrorStage

    lintRetVal = MsgBox("This process will update the Server and File location settings used by all users." & _
        vbCrLf & vbCrLf & "The information used to make this change is stored only in the Live database.  Therefore" & _
        vbCrLf & "this process is independent of the testing environment." & vbCrLf & vbCrLf & _
        "If you wish to proceed click YES!", vbCritical + vbYesNo, gconstrTitlPrefix & "System Lists")

    If lintRetVal = vbYes Then
    
        Decrypt gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, gconEncryptStatic
        
        gstrUserMode = gconstrLiveMode
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
            
            Encrypt gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, gconEncryptStatic
            MsgBox "Static has been deployed!" & vbCrLf & vbCrLf & "You are about to logged out of the system!", , gconstrTitlPrefix & "System Lists"
            
            UpdateLoader
            Unload Me
            Unhook
            End
        End With
    End If
End Sub

Private Sub cmdHelp_Click()

    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1012.htm"

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

Private Sub cmdPFSet_Click()
Dim lstrPFStartConsign As String
Dim lstrSQL As String
Dim lintRetVal As Integer
    
    lstrPFStartConsign = GetListCodeDesc("PForce Consignment Range", "START")
    
    lstrPFStartConsign = Left$(lstrPFStartConsign, Len(lstrPFStartConsign) - 1)
    
    If Trim$(lstrPFStartConsign) = "" Then
        MsgBox "The Start Consignment value is blank!, " & vbCrLf & _
        "or you haven't depolyed it yet.  Try logging out, " & _
        "you may have the updated file on your PC!", , gconstrTitlPrefix & "PF Reset"
        Exit Sub
    End If
    
    If Val(lstrPFStartConsign) = 0 Then
        MsgBox "The Start consignment number isn't a number!", , gconstrTitlPrefix & "PF Reset"
        Exit Sub
    End If
    
    lintRetVal = MsgBox("Would you like to reset the starting" & vbCrLf & _
        "Parcel Force consignment number to " & lstrPFStartConsign & " ?" & vbCrLf & vbCrLf & _
        "WARNING: This should only be done if Parcel Force have asked you to " & vbCrLf & _
        "change the consignment range! or, when the system goes live for the first time!", vbYesNo, gconstrTitlPrefix & "PF Reset")
    If lintRetVal = vbYes Then
        lstrSQL = "UPDATE System SET System.[Value] = " & lstrPFStartConsign & " WHERE (((System.Item)='LastPFConsignNumIncr'));"
        gdatCentralDatabase.Execute lstrSQL
    End If
    
    lintRetVal = MsgBox("Would you like to reset the Batch number, this should only be done when the system goes live!", vbYesNo, gconstrTitlPrefix & "PF Reset")
    If lintRetVal = vbYes Then
        lstrSQL = "UPDATE System SET System.[Value] = '0001' WHERE (((System.Item)='BatchIncr'));"
        gdatCentralDatabase.Execute lstrSQL
    End If

End Sub

Private Sub datLists_Reposition()
Dim lstrSQL As String

    Me.Refresh
    
    On Error Resume Next
    If Not (datLists.Recordset.BOF = True And datLists.Recordset.EOF = True) Then
        If Not IsNull(datLists.Recordset("ListNum")) Then
            lstrSQL = "Select * from ListDetailsMaster where ListNum=" & _
                datLists.Recordset("ListNum") & " order by SequenceNum;"
    
                
            datListDetails.RecordSource = lstrSQL
                   
            datListDetails.Refresh
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datListDetails.DatabaseName = gstrStatic.strCentralTestingDBFile
        datLists.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datListDetails.DatabaseName = gstrStatic.strCentralDBFile
        datLists.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datListDetails.Connect = gstrDBPasswords.strCentralDBPasswordString
        datLists.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    Select Case mbooSysUser
    Case True
        datLists.RecordSource = "select * from ListsMaster " & _
            "where sysuse = True order by Listname "
        cmdDeploy.Visible = True
    Case False
        datLists.RecordSource = "select * from ListsMaster " & _
            "where sysuse <> true order by Listname"
        cmdDeploy.Visible = False
    End Select
    
    NameForm Me
End Sub
Public Property Let SysUse(pbooSysUser As Boolean)

    mbooSysUser = pbooSysUser

End Property
Public Property Get SysUse() As Boolean

    SysUse = mbooSysUser

End Property

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim lintHeightOnePercent As Integer

    lintHeightOnePercent = Me.Height / 100
    
    With cmdClose
        .Top = Me.Height - 1230
        .Left = Me.Width - 1545
    End With
    
    With cmdMaxMin
        .Top = Me.Height - 1230
        .Left = cmdClose.Left - (cmdMaxMin.Width + 120)
    End With
    
    With cmdHelp
        .Top = Me.Height - 1230
        .Left = 120
    End With
    
    With cmdDeploy
        .Top = Me.Height - 1230
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdPFSet
        .Top = Me.Height - 1230
        .Left = cmdDeploy.Width + cmdDeploy.Left + 120
    End With
    
    With dbgListDetails
        .Width = Me.Width - 360 '360 '240
        .Top = lintHeightOnePercent * 42
        .Height = (cmdHelp.Top - .Top) - 120
    End With
    
    With dbgLists
        .Width = Me.Width - 360 '360 '240
       .Height = (dbgListDetails.Top - .Top) - 240 ' 600
    End With
    
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
