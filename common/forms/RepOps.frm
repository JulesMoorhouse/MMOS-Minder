VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Options"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Layout Report Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   4095
      Begin VB.Frame Frame4 
         Caption         =   "Margin Adjustments"
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3855
         Begin VB.TextBox txtLeftmargin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtTopMargin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   6
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "100 is approximately 2mm"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   2280
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Left Margin:"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Top Margin:"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detail Report Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4095
      Begin VB.Frame fraLineSpacing 
         Caption         =   "Line Spacing"
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton optLineSpacing 
            Caption         =   "Single"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optLineSpacing 
            Caption         =   "Double"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame fraMarings 
         Caption         =   "Margins"
         Height          =   1695
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton optMargin 
            Caption         =   "Narrow"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optMargin 
            Caption         =   "Wide"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   720
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame fraColourings 
         Caption         =   "Colouring"
         Height          =   1695
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton optColouring 
            Caption         =   "Bars on"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optColouring 
            Caption         =   "Bars off"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   1560
      TabIndex        =   8
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Size"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
      Begin VB.OptionButton optFontSize 
         Caption         =   "Small"
         Height          =   375
         Index           =   0
         Left            =   -480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFontSize 
         Caption         =   "Normal"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFontSize 
         Caption         =   "Large"
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   4380
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "27/06/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "09:05"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    SaveLocalFields
    Unload Me
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    GetLocalFields
    
End Sub
Sub GetLocalFields()

    With gstrReport
        If Printer.ColorMode = 1 Then
            optColouring(0).Enabled = False
            optColouring(1).Enabled = False
            fraColourings.Enabled = False
        Else
            optColouring(0).Enabled = True
            optColouring(1).Enabled = True
            fraColourings.Enabled = True
            If .booBarsOn = True Then
                optColouring(0).Value = True
            Else
                optColouring(1).Value = True
            End If
        End If
        
        Select Case .sngFontSize
        Case rpFontFactorSmall
            optFontSize(0).Value = True
        Case rpFontFactorNormal
            optFontSize(1).Value = True
        Case rpFontFactorLarge
            optFontSize(2).Value = True
        End Select
        
        Select Case .intSpacing
        Case rpSpacingSingle
            optLineSpacing(0).Value = True
        Case rpSpacingDouble
            optLineSpacing(1).Value = True
        End Select
        
        Select Case .lngMargins
        Case rpMarginNarrow
            optMargin(0).Value = True
        Case rpMarginWide
            optMargin(1).Value = True
        End Select
                
        
        If .booOptEnableBars = False Then
            optColouring(0).Enabled = False
            optColouring(1).Enabled = False
            fraColourings.Enabled = False
        End If
        
        If .booOptEnableLineSpace = False Then
            optLineSpacing(0).Enabled = False
            optLineSpacing(1).Enabled = False
            fraLineSpacing.Enabled = False
        End If
        
        If .booOptEnableMargins = False Then
            optFontSize(0).Enabled = False
            optFontSize(1).Enabled = False
            optFontSize(2).Enabled = False
            fraMarings.Enabled = False
        End If
        
        txtTopMargin = gstrReportLayout.lngTopMargAdj
        txtLeftmargin = gstrReportLayout.lngLeftMargAdj
        
    End With

End Sub

Sub SaveLocalFields()

    With gstrReport
        If Printer.ColorMode = 2 Then
            If optColouring(0).Value = True Then
                .booBarsOn = True
            Else
                .booBarsOn = False
            End If
        End If

        If optFontSize(0).Value = True Then
            .sngFontSize = rpFontFactorSmall
        ElseIf optFontSize(1).Value = True Then
            .sngFontSize = rpFontFactorNormal
        ElseIf optFontSize(2).Value = True Then
            .sngFontSize = rpFontFactorLarge
        End If

        If optLineSpacing(0).Value = True Then
            .intSpacing = rpSpacingSingle
        ElseIf optLineSpacing(1).Value = True Then
            .intSpacing = rpSpacingDouble
        End If

        If optMargin(0).Value = True Then
            .lngMargins = rpMarginNarrow
        ElseIf optMargin(1).Value = True Then
            .lngMargins = rpMarginWide
        End If
        
        gstrReportLayout.lngTopMargAdj = txtTopMargin
        gstrReportLayout.lngLeftMargAdj = txtLeftmargin
    End With

End Sub
Private Sub txtLeftmargin_LostFocus()

    txtLeftmargin = Val(txtLeftmargin)
    
End Sub

Private Sub txtTopMargin_LostFocus()

    txtTopMargin = Val(txtTopMargin)

End Sub
