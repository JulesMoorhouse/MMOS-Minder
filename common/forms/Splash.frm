VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Splash.frx":0000
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4320
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMCL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6720
      Picture         =   "Splash.frx":07D7
      ScaleHeight     =   735
      ScaleWidth      =   750
      TabIndex        =   5
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mindwarp Mail Order System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Tag             =   "CompanyProduct"
      Top             =   600
      Width           =   4965
   End
   Begin VB.Label lblProductName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MMOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   25.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   600
      Left            =   2760
      TabIndex        =   2
      Tag             =   "Product"
      Top             =   0
      Width           =   4845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved. Using this program constitutes acceptance of the license terms and conditions described in Help About."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   3480
      Width           =   3195
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©  Mindwarp Consultancy Ltd 2002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Tag             =   "Company"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   16
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   17
      Visible         =   0   'False
      X1              =   1680
      X2              =   1920
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   18
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      Visible         =   0   'False
      X1              =   1920
      X2              =   1680
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      Visible         =   0   'False
      X1              =   1920
      X2              =   2040
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      Visible         =   0   'False
      X1              =   2040
      X2              =   2280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      Visible         =   0   'False
      X1              =   2400
      X2              =   2400
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      Visible         =   0   'False
      X1              =   2640
      X2              =   2400
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      Visible         =   0   'False
      X1              =   2400
      X2              =   2640
      Y1              =   1440
      Y2              =   1680
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   2280
      Y1              =   1440
      Y2              =   1800
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   13
      Visible         =   0   'False
      X1              =   1680
      X2              =   1920
      Y1              =   3600
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   14
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   15
      Visible         =   0   'False
      X1              =   1920
      X2              =   1680
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      Visible         =   0   'False
      X1              =   2280
      X2              =   2400
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      Visible         =   0   'False
      X1              =   2040
      X2              =   2280
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      Visible         =   0   'False
      X1              =   2640
      X2              =   2400
      Y1              =   3600
      Y2              =   3840
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   19
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   3360
      Y2              =   3600
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      Visible         =   0   'False
      X1              =   2400
      X2              =   2640
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line lblSpanner 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   12
      Visible         =   0   'False
      X1              =   2040
      X2              =   1920
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   2520
      Y2              =   3360
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   12
      Visible         =   0   'False
      X1              =   1155
      X2              =   2235
      Y1              =   2160
      Y2              =   2760
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      Visible         =   0   'False
      X1              =   2715
      X2              =   2715
      Y1              =   2400
      Y2              =   3360
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      Visible         =   0   'False
      X1              =   1755
      X2              =   2595
      Y1              =   2400
      Y2              =   1800
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      Visible         =   0   'False
      X1              =   3195
      X2              =   3195
      Y1              =   2040
      Y2              =   3000
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      Visible         =   0   'False
      X1              =   2235
      X2              =   3195
      Y1              =   3720
      Y2              =   3000
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      Visible         =   0   'False
      X1              =   1155
      X2              =   1155
      Y1              =   2160
      Y2              =   3120
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      Visible         =   0   'False
      X1              =   1155
      X2              =   2115
      Y1              =   2160
      Y2              =   1440
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      Visible         =   0   'False
      X1              =   2235
      X2              =   2235
      Y1              =   2760
      Y2              =   3720
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      Visible         =   0   'False
      X1              =   2235
      X2              =   3195
      Y1              =   2760
      Y2              =   2040
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      Visible         =   0   'False
      X1              =   1635
      X2              =   2715
      Y1              =   1800
      Y2              =   2400
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      Visible         =   0   'False
      X1              =   2115
      X2              =   3195
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Line lblBox 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      Visible         =   0   'False
      X1              =   1155
      X2              =   2235
      Y1              =   3120
      Y2              =   3720
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   0
      Visible         =   0   'False
      X1              =   2205
      X2              =   2805
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   1
      Visible         =   0   'False
      X1              =   1365
      X2              =   2205
      Y1              =   2400
      Y2              =   1680
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   2
      Visible         =   0   'False
      X1              =   1365
      X2              =   1965
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   3
      Visible         =   0   'False
      X1              =   1965
      X2              =   2805
      Y1              =   2760
      Y2              =   2040
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   5
      Visible         =   0   'False
      X1              =   1965
      X2              =   2805
      Y1              =   3240
      Y2              =   2520
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   6
      Visible         =   0   'False
      X1              =   1365
      X2              =   1965
      Y1              =   2880
      Y2              =   3240
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   7
      Visible         =   0   'False
      X1              =   1365
      X2              =   1605
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   4
      Visible         =   0   'False
      X1              =   2565
      X2              =   2805
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   10
      Visible         =   0   'False
      X1              =   1965
      X2              =   2805
      Y1              =   3720
      Y2              =   3000
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   9
      Visible         =   0   'False
      X1              =   1365
      X2              =   1965
      Y1              =   3360
      Y2              =   3720
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   8
      Visible         =   0   'False
      X1              =   1365
      X2              =   1605
      Y1              =   3360
      Y2              =   3120
   End
   Begin VB.Line lblPaper 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      Index           =   11
      Visible         =   0   'False
      X1              =   2565
      X2              =   2805
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   10
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

    Sleep 2000
    
End Sub

Private Sub Form_Load()
Dim lintArrInc As Integer

    On Error Resume Next
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.WindowState = vbNormal
    
    lblProductName = gconstrProductShortName & " " & App.ProductName
    
    lblCompanyProduct = gconstrProductFullName
    
    Select Case UCase$(App.ProductName)
    Case "MAINTENANCE" '"ADMIN"
        For lintArrInc = 0 To 19
            lblSpanner(lintArrInc).Visible = True
        Next lintArrInc
    Case "LITE"
        
        lblProductName = "Lite / Demo Version"
        lblProductName.Visible = True
        For lintArrInc = 0 To 12
            lblBox(lintArrInc).Visible = True
        Next lintArrInc
    Case "CLIENT"
        lblProductName.Visible = False
        For lintArrInc = 0 To 12
            lblBox(lintArrInc).Visible = True
        Next lintArrInc
    Case Else '"MANAGER" '"REPORTING"
        For lintArrInc = 0 To 11
            lblPaper(lintArrInc).Visible = True
        Next lintArrInc
    End Select
    
    Const PLANES = 14
    Const BITSPIXEL = 12
    glngNumOfColours = GetDeviceCaps(hdc, PLANES) * 2 ^ GetDeviceCaps(hdc, BITSPIXEL)

    Select Case glngNumOfColours
    Case 8, 16
        Me.BackColor = vbCyan
        picMCL.BackColor = Me.BackColor
    Case 256
        Me.BackColor = RGB(192, 255, 255)
        picMCL.BackColor = Me.BackColor
    Case Else
        picMCL.BackColor = RGB(255, 255, 239)
        ShadeMe frmSplash, 600, 28, "BLUEISH"
        
    End Select

End Sub
Sub ShadeMe(pobjObject As Object, plngUnit As Long, plngDrawWidth As Long, pstrColourScheme As String)
Dim lintArrInc As Integer
Dim llngDefR As Long
Dim llngDefG As Long
Dim llngDefB As Long
Dim lstrRGorB As String
Dim llngColourInc As Long

    pobjObject.Cls
    
    Select Case pstrColourScheme
    Case "BLUEISH"
        llngDefR = 174:           llngDefG = 200:           llngDefB = 243
        lstrRGorB = "RDGR"
        llngColourInc = 5
    Case "DARKBLUEISH"
        llngDefR = 0:           llngDefG = 0:           llngDefB = 100
        lstrRGorB = "BLUE"
        llngColourInc = 15
    Case "YELLOWISH"
        llngDefR = 255:         llngDefG = 255:         llngDefB = 0
        lstrRGorB = "BLUE"
        llngColourInc = 2
    Case "LILAC"
        llngDefR = 194:         llngDefG = 190:         llngDefB = 244
        lstrRGorB = "GREEN"
        llngColourInc = -5
    End Select

    pobjObject.DrawWidth = plngDrawWidth
    pobjObject.ForeColor = SetColour(lstrRGorB, llngColourInc, llngDefR, llngDefG, llngDefB)
    
    For lintArrInc = 0 To 20
        pobjObject.ForeColor = SetColour(lstrRGorB, llngColourInc * lintArrInc, llngDefR, llngDefG, llngDefB)
        pobjObject.Line (lintArrInc * plngUnit, -200)-(-200, lintArrInc * plngUnit)
    Next lintArrInc
    
End Sub
Function SetColour(pstrRGorB As String, plngColourInc As Long, plngDefR As Long, plngDefG As Long, plngDefB As Long) As Long

On Error Resume Next
    Select Case pstrRGorB
    Case "RED"
        SetColour = RGB((plngColourInc) + plngDefR, plngDefG, plngDefB)
    Case "GREEN"
        SetColour = RGB(plngDefR, (plngColourInc) + plngDefG, plngDefB)
    Case "BLUE"
        SetColour = RGB(plngDefR, plngDefG, (plngColourInc) + plngDefB)
    Case "BLGR"
        SetColour = RGB(plngDefR, (plngColourInc) + plngDefG, (plngColourInc) + plngDefB)
    Case "RDGR"
        SetColour = RGB((plngColourInc) + plngDefR, (plngColourInc) + plngDefG, plngDefB)
    End Select
    
End Function

Private Sub fraMainFrame_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

