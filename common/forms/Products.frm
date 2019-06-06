VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildProducts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please select products"
   ClientHeight    =   4830
   ClientLeft      =   30
   ClientTop       =   225
   ClientWidth     =   11085
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   4080
      Width           =   375
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   2040
      Top             =   4080
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1305
   End
   Begin VB.TextBox txtSearchCriteria 
      Height          =   288
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   2412
   End
   Begin VB.Frame fraSearchBy 
      Caption         =   "Search By"
      Height          =   1092
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2412
      Begin VB.OptionButton optSearchField 
         Caption         =   "&Catalogue Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1692
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Item &Description"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1452
      End
      Begin VB.OptionButton optSearchField 
         Caption         =   "Class &Item"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1812
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1305
   End
   Begin VB.Data datProducts 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2292
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14367
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "4/27/2019"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "9:37 AM"
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid dbgProducts 
      Bindings        =   "Products.frx":0000
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "Products.frx":001A
      TabIndex        =   11
      Top             =   1320
      Width           =   10815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Products.frx":15BD
      Height          =   375
      Left            =   7800
      OleObjectBlob   =   "Products.frx":15D7
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   $"Products.frx":305A
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   7110
      TabIndex        =   15
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblSubstitute 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label lblFoundNumber 
      Caption         =   "Found 0 records"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Choose a product, then use either Y or N or Space to select it."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   6255
   End
End
Attribute VB_Name = "frmChildProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrScreenHelpFile As String
Function FillProductsGrid(pobjList As Object, pdatControl As Object, pstrCiteria As String, pobjOptionList As Object)

Dim lstrSQL As String

    If Trim$(pstrCiteria) = "" Then
        MsgBox "You must enter some criteria to search upon", , gconstrTitlPrefix & "Searching"
        
        Exit Function
    End If
    
    lstrSQL = "select * from " & gtblProducts & " where "
    
    If pobjOptionList(0).Value Then 'Catalogue Number
        sbStatusBar.Panels(1).Text = StatusText(4)
        lstrSQL = lstrSQL & "CatNum like '*" & pstrCiteria & "*'"
    ElseIf pobjOptionList(1).Value Then 'Item Description
        sbStatusBar.Panels(1).Text = StatusText(5)
        lstrSQL = lstrSQL & "ItemDescription like '*" & pstrCiteria & "*'"
    ElseIf pobjOptionList(2).Value Then 'Class item
        sbStatusBar.Panels(1).Text = StatusText(6)
        lstrSQL = lstrSQL & "ClassItem like '*" & pstrCiteria & "*'"
    End If
    
    lstrSQL = lstrSQL & " order by ItemDescription, ClassItem"
    
    pdatControl.RecordSource = lstrSQL
    pdatControl.Refresh
    pobjList.Refresh
 
    sbStatusBar.Panels(1).Text = StatusText(0)
Exit Function
Err_Hand:
Select Case Err.Number
Case 3261
    MsgBox "Someone has the Central Database exclusively Locked!" & vbCrLf & vbCrLf & _
        "Please inform IT!", , gconstrTitlPrefix & "Fill Products grid"
    
    Exit Function
Case Else
    MsgBox "Please report this error!" & vbCrLf & vbCrLf & _
        Err.Number & " " & Err.Description, , gconstrTitlPrefix & "Fill Products Grid"
    
    Resume Next
End Select
End Function

Private Sub cmdFind_Click()
Dim lstrSeacrhCriteria As String
Dim lstrRetValCriteria As String
Dim lstrOutAutoCatNum As String

    lblSubstitute = ""
    
    If optSearchField(0).Value Then 'Catalogue Number
        lstrRetValCriteria = CheckForSubstitutions(txtSearchCriteria)
        If UCase$(Trim$(lstrRetValCriteria)) <> UCase$(Trim$(txtSearchCriteria)) Then
            'substitution found
            If Left$(lstrRetValCriteria, 10) = "STOCKAUTO#" Then
                lstrOutAutoCatNum = Mid$(lstrRetValCriteria, 11, Len(lstrRetValCriteria) - 10)
                'Check stock of primary product
                If CheckStockSpecificProduct(txtSearchCriteria) > 0 Then
                    lstrSeacrhCriteria = txtSearchCriteria
                Else
                    lstrSeacrhCriteria = lstrOutAutoCatNum
                    lblSubstitute = "Selected item " & Trim$(txtSearchCriteria) & " is out of stock, use this item instead!"
                End If
            Else ' e.g Auto
                lstrSeacrhCriteria = lstrRetValCriteria
                lblSubstitute = "Automatic Substitution made for " & Trim$(txtSearchCriteria) & ""
            End If
        Else
            lstrSeacrhCriteria = txtSearchCriteria
        End If
    Else
        lstrSeacrhCriteria = txtSearchCriteria
    End If
    
    FillProductsGrid dbgProducts, datProducts, lstrSeacrhCriteria, optSearchField
    
    If dbgProducts.ApproxCount > 0 Then
        dbgProducts.SetFocus
        sbStatusBar.Panels(1).Text = StatusText(0)
    Else
        sbStatusBar.Panels(1).Text = StatusText(15)
    End If
    
    lblFoundNumber = "Found " & datProducts.Recordset.RecordCount & " records."
    
End Sub

Private Sub cmdHelp_Click()
    
    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdSelect_Click()

    Unload Me
    
End Sub

Private Sub dbgProducts_AfterColEdit(ByVal ColIndex As Integer)
Dim lstrRetValCriteria As String
Dim lstrOutAutoCatNum As String

    Select Case ColIndex
    Case 3 'Qty Field
        If Val(dbgProducts.Text) > 0 Then
            If Val(dbgProducts.Columns(7).Value) < 5 And gstrReferenceInfo.booStockThreashold Then
                'check for substitue item.
                lstrRetValCriteria = CheckForSubstitutions(dbgProducts.Columns(0).Value)
                If Left$(lstrRetValCriteria, 10) = "STOCKAUTO#" Then
                    lstrOutAutoCatNum = Mid$(lstrRetValCriteria, 11, Len(lstrRetValCriteria) - 10)
                    'Add substitue item, linenum & qty to orderlines table
                    gintOrderLineNumber = gintOrderLineNumber + 10
                    AddProductSpecific lstrOutAutoCatNum, CLng(dbgProducts.Text), CLng(gintOrderLineNumber)
                    
                    dbgProducts.Text = 0
                    lblSubstitute = "Substitue item used, as item " & Trim$(dbgProducts.Columns(0).Value) & " is out of stock!"
                    Exit Sub
                End If
                MsgBox "This item is out of stock, or in the threshold.", , gconstrTitlPrefix & "Product Selection"
                dbgProducts.Col = 3
                dbgProducts.Text = 0
                dbgProducts.Col = 2
                dbgProducts.Text = 1
                gintOrderLineNumber = gintOrderLineNumber + 10
                dbgProducts.Col = 5
                dbgProducts.Text = gintOrderLineNumber
                dbgProducts.Col = 3
            ElseIf Val(dbgProducts.Columns(7).Value) < Val(dbgProducts.Text) Then
                'check for substitue item.
                lstrRetValCriteria = CheckForSubstitutions(dbgProducts.Columns(0).Value)
                If Left$(lstrRetValCriteria, 10) = "STOCKAUTO#" Then
                    lstrOutAutoCatNum = Mid$(lstrRetValCriteria, 11, Len(lstrRetValCriteria) - 10)
                    'Add substitue item, linenum & qty to orderlines table
                    gintOrderLineNumber = gintOrderLineNumber + 10
                    AddProductSpecific lstrOutAutoCatNum, CLng(dbgProducts.Text), CLng(gintOrderLineNumber)
                    
                    dbgProducts.Text = 0
                    lblSubstitute = "Substitue item used, as item " & Trim$(dbgProducts.Columns(0).Value) & " is out of stock!"
                    Exit Sub
                End If
                
                MsgBox "There are not enough items in stock to order this amount!", , gconstrTitlPrefix & "Product Selection"
                dbgProducts.Col = 3
                dbgProducts.Text = 0
                
                gintOrderLineNumber = gintOrderLineNumber + 10
                dbgProducts.Col = 5
                dbgProducts.Text = gintOrderLineNumber
                dbgProducts.Col = 3
            Else
                dbgProducts.Col = 2
                dbgProducts.Text = 1
                gintOrderLineNumber = gintOrderLineNumber + 10
                dbgProducts.Col = 5
                dbgProducts.Text = gintOrderLineNumber
                dbgProducts.Col = 3
            End If
        End If
        dbgProducts.Col = 3
    End Select
    
    sbStatusBar.Panels(1).Text = StatusText(0)
    
End Sub

Private Sub dbgProducts_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case dbgProducts.Col
    Case 3
    Case Else
        dbgProducts.Col = 2
        Select Case KeyCode
        Case 78 'N
            dbgProducts.Text = 0 '"No"
            dbgProducts.Col = 3: dbgProducts.Text = 0
        Case 89 'Y
            dbgProducts.Text = 1 '"Yes"
            dbgProducts.Col = 3: dbgProducts.Text = 1
            
            gintOrderLineNumber = gintOrderLineNumber + 10
            dbgProducts.Col = 5
            dbgProducts.Text = gintOrderLineNumber
            
            dbgProducts.Col = 2
        Case 32 ' Space - Toggle
            If dbgProducts.Text = "No" Then
                dbgProducts.Text = 1
                dbgProducts.Col = 3: dbgProducts.Text = 1
                
                gintOrderLineNumber = gintOrderLineNumber + 10
                dbgProducts.Col = 5
                dbgProducts.Text = gintOrderLineNumber
                
                dbgProducts.Col = 2
            Else
                dbgProducts.Text = 0
                dbgProducts.Col = 3: dbgProducts.Text = 0: dbgProducts.Col = 2
            End If
        Case Else
             Exit Sub
        End Select
    End Select
    
    sbStatusBar.Panels(1).Text = StatusText(0)
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    UpdateProductQtyFromMaster
    
    Select Case gstrUserMode
    Case gconstrTestingMode
        datProducts.DatabaseName = gstrStatic.strLocalTestingDBFile
    Case gconstrLiveMode
        datProducts.DatabaseName = gstrStatic.strLocalDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datProducts.Connect = gstrDBPasswords.strLocalDBPasswordString
    End If
    
    datProducts.RecordSource = "select * from " & gtblProducts & " where selected = 1 order by ItemDescription"
    
    SetupHelpFileReqs
    
End Sub

Private Sub optSearchField_Click(Index As Integer)

    txtSearchCriteria.SetFocus

End Sub

Private Sub optSearchField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    txtSearchCriteria.SetFocus

End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub txtSearchCriteria_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdFind_Click
    End If

End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/ChildProducts.xml>WhatsScreen"
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdSelect.WhatsThisHelpID = IDH_STANDARD_NEXT
    cmdFind.WhatsThisHelpID = IDH_STANDARD_FIND
    
    dbgProducts.WhatsThisHelpID = IDH_CHIPRODS_GRIDPRODS
    
    fraSearchBy.WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(0).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(1).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    optSearchField(2).WhatsThisHelpID = IDH_STANDARD_SEARCHBY
    
    txtSearchCriteria.WhatsThisHelpID = IDH_CHIPRODS_SEARCHCRIT

End Sub
