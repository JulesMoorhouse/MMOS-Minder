VERSION 5.00
Begin VB.Form frmMinder 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Clean"
   ClientHeight    =   210
   ClientLeft      =   4125
   ClientTop       =   3045
   ClientWidth     =   1560
   Icon            =   "Minder.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   210
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3640
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minder..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblDDE 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This label is used for DDE connection to the Program Manager"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "frmMinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    Me.WindowState = vbNormal
    
    Select Case UCase$(Command)
    Case "APP"
        Timer1.Enabled = True
    End Select

End Sub

Private Sub Timer1_Timer()
Dim lRetVal As Variant
    
    DelUnauthorisedFiles
    DelWinTmp
    Select Case Left$(CheckWindowsVersion, 20)
    Case "Microsoft Windows NT"
        'no defrag
    Case Else
         RunNWait "Scandskw c: /n"
         RunNWait "Defrag c: /concise /noprompt /f"
    End Select
    
    gbooDoneTimeEvent = True
    Timer1.Enabled = False
    End
    
End Sub
