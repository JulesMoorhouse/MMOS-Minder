VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmChildCalendar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Date..."
   ClientHeight    =   2985
   ClientLeft      =   30
   ClientTop       =   210
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1305
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2532
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3612
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4471
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1999
      Month           =   12
      Day             =   23
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChildCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcalDate As Date

Private Sub cmdSelect_Click()

    mcalDate = Calendar1.Value
    Unload Me
End Sub

Private Sub Form_Load()
    
    If gbooJustPreLoading Then
        Exit Sub
    End If

    Calendar1.Value = mcalDate
    
End Sub
Public Property Let CalDate(pcalDate As String)

    If IsDate(pcalDate) Then
        mcalDate = CDate(pcalDate)
    End If

End Property
Public Property Get CalDate() As String

    CalDate = Format$(mcalDate, "dd/mmm/yyyy")

End Property
