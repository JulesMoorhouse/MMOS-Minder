VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelpAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFullAgreement 
      Height          =   375
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "HelpAbt.frx":0000
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLiteAgreement 
      Height          =   375
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "HelpAbt.frx":18A8
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDisagree 
      Caption         =   "I &Disagree"
      Height          =   360
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1305
   End
   Begin VB.TextBox txtLicense 
      Height          =   3855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   6495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "I &Agree"
      Height          =   360
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1305
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Software licence Agreement"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMCL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6120
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmHelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisagree_Click()

    MsgBox "Please uninstall this software and discontinue using it!", vbInformation, gconstrTitlPrefix & "Licence Agreement"
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
   frmButtons.ImageList16Cols.UseMaskColor = True
   frmButtons.ImageList16Cols.MaskColor = vbRed 'vbGreen
   frmButtons.ImageList16Cols.BackColor = picMCL.BackColor
   frmButtons.ImageList16Cols.ListImages(11).Draw picMCL.hdc, 0, 0

    If UCase$(App.ProductName) = "LITE" Then
        txtLicense = txtLiteAgreement
    Else
        txtLicense = txtFullAgreement
    End If
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

