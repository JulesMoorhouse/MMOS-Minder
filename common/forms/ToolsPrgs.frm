VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChildToolProgs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "External Programs"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1305
   End
   Begin VB.TextBox txtAddressingProgram 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4395
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   510
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Height          =   360
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Addressing Program:"
      Height          =   240
      Left            =   150
      TabIndex        =   4
      Top             =   360
      Width           =   1860
   End
End
Attribute VB_Name = "frmChildToolProgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()

    On Error Resume Next
    
    With CommonDialog1
        If txtAddressingProgram <> "" Then
            .FileName = txtAddressingProgram
        End If
        .Flags = cdlOFNHideReadOnly
        .Filter = "All Programs *.exe|*.exe|All Files *.*|*.*"
        .ShowOpen
        txtAddressingProgram = .FileName
    End With
    
End Sub

Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub cmdOK_Click()

    If Dir(txtAddressingProgram) <> "" And Trim$(txtAddressingProgram) <> "" Then
        SaveSetting gstrIniAppName, "QARAPID", "Location", txtAddressingProgram
        MsgBox "Program accepted!", vbInformation, gconstrTitlPrefix & "External Programs"
    Else
        MsgBox "Program not accepted!", vbCritical, gconstrTitlPrefix & "External Programs"
        
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim lstrQAProg As String
    
    If gbooJustPreLoading Then 
        Exit Sub
    End If
    
    NameForm Me 
    
    lstrQAProg = GetSetting(gstrIniAppName, "QARAPID", "Location")
    
    If Dir(lstrQAProg) <> "" Then
        txtAddressingProgram = lstrQAProg
    End If
    
End Sub
