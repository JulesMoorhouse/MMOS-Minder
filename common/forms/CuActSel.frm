VERSION 5.00
Begin VB.Form frmCustAcctSel 
   Caption         =   "Please Select an Account..."
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   10485
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   10485
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.TextBox txtPostCode 
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   2412
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1305
   End
   Begin VB.TextBox txtSurname 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2412
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1305
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   360
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      WhatsThisHelpID =   200
      Width           =   1305
   End
   Begin VB.CommandButton cmdOverideNSelect 
      Caption         =   "&Add New"
      Height          =   360
      Left            =   7608
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   2160
      Top             =   6120
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton xcmdOldFlag 
      Caption         =   "&Remove Old Account"
      Height          =   492
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   6153
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   2412
      Begin VB.OptionButton xoptSearchField 
         Caption         =   "&BPCS Number"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Tag             =   "BPCS"
         Top             =   960
         Width           =   1812
      End
      Begin VB.OptionButton xoptSearchField 
         Caption         =   "Customer &Name"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Tag             =   "Name"
         Top             =   720
         Width           =   1812
      End
      Begin VB.OptionButton xoptSearchField 
         Caption         =   "&Post Code"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Tag             =   "Postcode"
         Top             =   480
         Width           =   1452
      End
      Begin VB.OptionButton xoptSearchField 
         Caption         =   "&Customer Number"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Tag             =   "CustNumber"
         Top             =   240
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.TextBox txtCustNum 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2412
   End
   Begin VB.ListBox lstAccounts 
      Height          =   5610
      IntegralHeight  =   0   'False
      Left            =   2850
      TabIndex        =   10
      Top             =   1200
      Width           =   7542
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   19
      Top             =   7050
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   1244
   End
   Begin VB.Label lblCustomerCount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You have created 50 Customer Accounts and have therefore reached the end of demo version evaulation period"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Postcode :-"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Surname :-"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Number :-"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "frmCustAcctSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim llngCustNumber() As Long
Dim lstrCustName() As String
Dim lstrInUseBy() As String
Dim lstrDBIndicator() As String

Dim mstrRoute As String
Dim lstrSearchType As String
Dim lstrCriteria As String
Dim lstrScreenHelpFile As String

Private Sub cmdBack_Click()

    Unload Me
    frmAbout.Show
    
End Sub

Private Sub cmdFind_Click(Index As Integer)

    'FillCustAccountsList lstAccounts, optSearchField, txtSearchCriteria, llngCustNumber(), lstrCustName(), lstrInUseBy(), lstrDBIndicator()
    Select Case Index
    Case 0 'custnum
        FillCustAccountsListCustNumIndexed lstAccounts, txtCustNum, llngCustNumber(), lstrCustName(), lstrInUseBy(), lstrDBIndicator()
        lstrSearchType = "CustNumber"
        lstrCriteria = txtCustNum
    Case 1 'surname
        FillCustAccountsList lstAccounts, txtSurname, llngCustNumber(), lstrCustName(), lstrInUseBy(), lstrDBIndicator(), "SURNAME"
        lstrSearchType = "Name"
        lstrCriteria = txtSurname
    Case 2 'postcode
        FillCustAccountsList lstAccounts, txtPostcode, llngCustNumber(), lstrCustName(), lstrInUseBy(), lstrDBIndicator(), "POSTCODE"
        lstrSearchType = "Postcode"
        lstrCriteria = txtPostcode
    End Select
    'PaintFormMenu Me
    
End Sub

Private Sub cmdHelp_Click()

    'If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
    '    'RunxNWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1002.htm"
    '    RunNDontWait FindProgram("IEXPLORE") & " " & gstrStatic.strServerPath & "Help\h1002.htm"
    'Else
        glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    'End If
    
End Sub
Private Sub cmdOldFlag_Click()
Dim lintArrInc As Integer
Dim lstrFoundItem As Boolean

    On Error GoTo ErrHandler
    lstrFoundItem = False
    If gstrUserMode = gconstrLiveMode Then
        For lintArrInc = 0 To lstAccounts.ListCount
            If lstAccounts.Selected(lintArrInc) Then
                If Right$(Trim$(lstAccounts.List(lintArrInc)), 5) = "(OLD)" Then
                    UpdateOldAcIndicator llngCustNumber(lintArrInc)
                    lstrFoundItem = True
                    Exit Sub
                Else
                    MsgBox "This feature only works on OLD (introduced from another system) Accounts." & vbCrLf & _
                        "Please use Account maintenance and select take off mailing " & vbCrLf & _
                        "List within the Accounts screen, for NEW accounts.", , gconstrTitlPrefix & "System Integration"
                End If
            End If
        Next lintArrInc
        If lstrFoundItem = False Then
            MsgBox "you must select an OLD item from the list!", , gconstrTitlPrefix & "System Integration"
        End If
    Else
        MsgBox "You may Not remove items from the Old database, in Training Mode.", , gconstrTitlPrefix & "System Integration"
    End If

Exit Sub
ErrHandler:
Exit Sub
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdOverideNSelect_Click()

    lstAccounts.Clear
    Call cmdSelect_Click
    
End Sub

Private Sub cmdSelect_Click()
Dim lintArrInc As Integer
Dim lintMsgRetVal As Integer
Dim lstrInUseByFlag As String
Dim lintArrInc2 As Integer
Const lconstrEvalUp = "You have created 50 Customer Accounts and have therefore reached the end of demo version evaulation period!"

    On Error GoTo Exitsub
    
    If lstAccounts.ListCount > 0 Then
        For lintArrInc = 0 To lstAccounts.ListCount - 1
            If lstAccounts.Selected(lintArrInc) Then
                lintMsgRetVal = MsgBox("Select Customer Account : " & Chr(34) & lstrCustName(lintArrInc) & Chr(34) & " (M" & llngCustNumber(lintArrInc) & ") ?", vbYesNo + vbInformation, gconstrTitlPrefix & "Account Select")
                If lintMsgRetVal = vbYes Then
                   
                    'If Trim$(lstrInUseBy(lintArrInc)) <> "" Then
                    '    'remind user AC is flagged in use
                    '    lintMsgRetVal = MsgBox("This Customer Account " & Chr(34) & lstrCustName(lintArrInc) & Chr(34) & " is in use by " & lstrInUseBy(lintArrInc) & vbCrLf & vbCrLf & _
                    '        "WARNING: Please ensure that user " & lstrInUseBy(lintArrInc) & " has finished with this account!" & vbCrLf & _
                    '        "Proceed?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Account Select")
                    'End If
                    
                    'If lintMsgRetVal = vbYes Or Trim$(lstrInUseBy(lintArrInc)) = "" Then
                    If CheckAcctInUseAvail(llngCustNumber(lintArrInc)) = True Then
                    
                        If lstrDBIndicator(lintArrInc) = "N" Then
                            ShowStatus 7 ' 
                            'sbStatusBar.Panels(1).Text = StatusText(7)
                            gstrCustomerAccount.lngCustNum = llngCustNumber(lintArrInc)
                            GetCustomerAccount gstrCustomerAccount.lngCustNum, True
                            Unload Me
                           
                            Set frmCustAcctSel = Nothing
                            
                            'mdiMain.DrawButtonSet mstrRoute ' 
                            
                            Select Case Me.Route
                            Case gconstrEnquiry
                                Set gstrCurrentLoadedForm = frmOrdHistory ' 
                                frmOrdHistory.Route = Me.Route 'only has route order enquiry
                                frmOrdHistory.Show
                            Case Else
                               
                                gstrOrderEntryOrderStatus = ""
                                        
                                Set gstrCurrentLoadedForm = frmAccount ' 
                                frmAccount.Route = Me.Route
                                frmAccount.Show
                            End Select
                            
                            Exit Sub

                        End If
                    Else
                        'PaintFormMenu Me
                        Exit Sub
                    End If
                    'End If
                Else
                    'PaintFormMenu Me
                    Exit Sub
                End If
            End If
        Next lintArrInc
    ElseIf lstAccounts.ListCount = lintArrInc Or lstAccounts.ListCount = 0 Then
        If mstrRoute = gconstrEntry Then

           
            If UCase$(App.ProductName) = "LITE" Then
                If CountCustAccounts >= 50 Then
                    MsgBox lconstrEvalUp, vbInformation, gconstrTitlPrefix & "Account Select"
                    Exit Sub
                End If
            End If
                
           
            If UCase$(App.ProductName) = "LITE" Then
                If cmdOverideNSelect.Enabled = False Then
                    MsgBox lconstrEvalUp, vbInformation, gconstrTitlPrefix & "Account Select"
                End If
            End If
            
            lintMsgRetVal = MsgBox("You need to Select or create a Customer Account" & vbCrLf & vbCrLf & _
                "Do you want to create a new account?", vbInformation + vbYesNo, gconstrTitlPrefix & "Account Select")
            If lintMsgRetVal = vbYes Then
                'Create new account
               
                'lstrInUseByFlag = Trim$(gstrGenSysInfo.strUserName) & " " & Now()
                lstrInUseByFlag = LockingPhaseGen(True)
                
                AddNewCustomerAccount lstrInUseByFlag
                GetCustomerAccountNum lstrInUseByFlag
                
               
                'frmAccount.Criteria = txtSearchCriteria
                frmAccount.Criteria = lstrCriteria
                
               
                'For lintArrInc2 = 0 To 3
                '    If optSearchField(lintArrInc2).Value = True Then
                '        frmAccount.CriteriaType = optSearchField(lintArrInc2).Tag
                '    End If
                'Next lintArrInc2
                frmAccount.CriteriaType = lstrSearchType
                
               
                gstrOrderEntryOrderStatus = ""
                
                'mdiMain.DrawButtonSet mstrRoute ' 
                Set gstrCurrentLoadedForm = frmAccount ' 
                
                frmAccount.Route = Me.Route
                
                Unload Me
                frmAccount.Show
                Exit Sub
            End If
        ElseIf mstrRoute = gconstrEnquiry Then
            MsgBox "You may not create New accounts in Enquiries!", , gconstrTitlPrefix & "Account Select"
        End If
    End If
    
    'PaintFormMenu Me
    
    ''txtSearchCriteria.SetFocus
Exit Sub
Exitsub:
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Select Case (KeyCode)
    Case vbKeyF1
        'Call cmdHelp_Click
    End Select
    
End Sub

Private Sub Form_Load()
    
    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me

    ShowBanner Me, Me.Route
    
   
    If UCase$(App.ProductName) = "LITE" Then
        If CountCustAccounts >= 50 Then
            lblCustomerCount.Visible = True
            cmdOverideNSelect.Enabled = False
        End If
    End If
    
    SetupHelpFileReqs
    
End Sub


Function FillCustAccountsList(pobjList As Object, pstrCiteria As String, plngCustNumber() As Long, pstrCustName() As String, pstrInUseBy() As String, pstrDBIndicator() As String, pstrParam As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    If Trim$(pstrCiteria) = "" Then
        MsgBox "You must enter some criteria to search upon", , gconstrTitlPrefix & "Searching"
        'PaintFormMenu Me
        Exit Function
    End If
    
    Busy True, Me
    
    Select Case pstrParam
    Case "CUSTNUM" ' can't happen
    'If pobjOptionList(0).Value Then 'Customer Number
        If UCase$(Left$(pstrCiteria, 1)) = "M" Then
            pstrCiteria = Right$(pstrCiteria, Len(pstrCiteria) - 1)
        End If
        If Val(pstrCiteria) = 0 Then
            Busy False, Me
            Exit Function
        End If
        'sbStatusBar.Panels(1).Text = StatusText(8)
        ShowStatus 8 ' 
        'Converted table names to constants
        lstrSQL = "select * from " & gtblCustAccounts & " where "
'        lstrSQL = lstrSQL & "CustNum like '*" & pstrCiteria & "*'"
        lstrSQL = lstrSQL & "CustNum = " & pstrCiteria & ""
        lstrSQL = lstrSQL & " order by surname, postcode"
        
    Case "POSTCODE"
    'ElseIf pobjOptionList(1).Value Then ' Post Code
        'sbStatusBar.Panels(1).Text = StatusText(9)
        ShowStatus 9 ' 
'        lstrSQL = "Select surname, Initials, Add1, PostCode, AcctinUseByFlag " & _
            "from CustAccounts Where "
        lstrSQL = "Select * " & _
            "from " & gtblCustAccounts & " Where "
        lstrSQL = lstrSQL & "Postcode like '*" & pstrCiteria & "*'"
        lstrSQL = lstrSQL & " and DBIndicator = 'N'"
        lstrSQL = lstrSQL & " order by surname, postcode "

    Case "SURNAME"
    'ElseIf pobjOptionList(2).Value Then 'Customer name
        'sbStatusBar.Panels(1).Text = StatusText(10)
        ShowStatus 10 ' 
        ' lstrSQL = "Select surname, Initials, Add1, PostCode, AcctinUseByFlag " & _
            "from CustAccounts Where "
        'Converted table names to constants
         lstrSQL = "Select * " & _
            "from " & gtblCustAccounts & " Where "
        lstrSQL = lstrSQL & "Surname like '*" & pstrCiteria & "*'"
        lstrSQL = lstrSQL & " and DBIndicator = 'N'"
        lstrSQL = lstrSQL & " order by surname, postcode "
        'lstrSQL = lstrSQL & "Union Select surname, Initials, Add1, PostCode, " & _
            "AcctinUseByFlag from OldCustAccounts where "
       '' lstrSQL = lstrSQL & "Union Select * from OldCustAccounts where "
       '' lstrSQL = lstrSQL & "Surname like '*" & pstrCiteria & "*'"
       '' lstrSQL = lstrSQL & " and DBIndicator = 'O'"
       '' lstrSQL = lstrSQL & " order by surname, postcode"
       
    Case "BPCS" ' can't happen
    'ElseIf pobjOptionList(3).Value Then 'BPCS NUmber
        'sbStatusBar.Panels(1).Text = StatusText(18)
        ShowStatus 18 ' 
        'Converted table names to constants
        lstrSQL = "select * from " & gtblCustAccounts & " where "
        lstrSQL = lstrSQL & "BPCSCusNum like '*" & pstrCiteria & "*'"
        lstrSQL = lstrSQL & " order by surname, postcode"
    End Select
    'End If
    
    
    pobjList.Clear
    pobjList.BackColor = vbWindowBackground
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            ReDim Preserve plngCustNumber(llngRecCount)
            ReDim Preserve pstrCustName(llngRecCount)
            ReDim Preserve pstrInUseBy(llngRecCount)
            ReDim Preserve pstrDBIndicator(llngRecCount)
                        
            'If pobjOptionList(0).Value Or pobjOptionList(3).Value Then
            If Trim$(.Fields("DBIndicator") & "") = "O" Then
                pobjList.AddItem Trim$(.Fields("Surname") & "") & " " & _
                    Trim$(.Fields("Initials") & "") & ", " & Trim$(.Fields("Add1") & "") & _
                    ", " & Trim$(.Fields("Postcode") & "") & _
                    " (OLD)" '& _
                    IIf(Trim$(.Fields("AcctInUseByFlag") & _
                    "") = "", "", " - " & UCase$(Trim$(.Fields("AcctInUseByFlag") & "")))
                    'Last bit commented 27/12/01
            Else
                pobjList.AddItem Trim$(.Fields("Surname") & "") & " " & _
                    Trim$(.Fields("Initials") & "") & ", " & Trim$(.Fields("Add1") & "") & _
                    ", " & Trim$(.Fields("Postcode") & "") & _
                    " (M" & Trim$(.Fields("CustNum") & "") & ")" '& _
                    IIf(Trim$(.Fields("AcctInUseByFlag") & _
                    "") = "", "", " - " & UCase$(Trim$(.Fields("AcctInUseByFlag") & "")))
                    'Last bit commented 27/12/01
            End If
            plngCustNumber(llngRecCount - 1) = ValNull(.Fields("CustNum"))
            pstrCustName(llngRecCount - 1) = Trim$(.Fields("Surname") & "") & " " & .Fields("Initials") & ""
            pstrInUseBy(llngRecCount - 1) = Trim$(.Fields("AcctInUseByFlag") & "")
            pstrDBIndicator(llngRecCount - 1) = Trim$(.Fields("DBIndicator") & "")
    
            .MoveNext
        Loop
        'pobjList.AddItem ""
    End With
    
    If llngRecCount = 0 Then
        'pobjList.AddItem ""
        'pobjList.BackColor = vbActiveBorder
        ReDim plngCustNumber(0)
        plngCustNumber(0) = 0
        Busy False, Me
    ElseIf llngRecCount = 1 Then
        pobjList.Selected(0) = True
        Busy False, Me
        DoEvents
        cmdSelect_Click
        DoEvents
    Else
        Busy False, Me
        ' On Error Resume Next
        ' pobjList.SetFocus
        ' On Error GoTo ErrHandler
    End If
    
    lsnaLists.Close
    'sbStatusBar.Panels(1).Text = StatusText(0)
    ShowStatus 0 ' 
Exit Function
ErrHandler:
    
    Busy False, Me
    
    Select Case GlobalErrorHandler(Err.Number, "frmCustAcctSel.FillCustAccountsList", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        'sbStatusBar.Panels(1).Text = StatusText(0)
        ShowStatus 0 ' 
        Exit Function
    Case Else
        Resume Next
    End Select


End Function
Function FillCustAccountsListCustNumIndexed(pobjList As Object, pstrCiteria As String, plngCustNumber() As Long, pstrCustName() As String, pstrInUseBy() As String, pstrDBIndicator() As String)
Dim ltabCustAccount As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    If Trim$(pstrCiteria) = "" Then
        MsgBox "You must enter some criteria to search upon", , gconstrTitlPrefix & "Searching"
        ShowStatus 0
        'PaintFormMenu Me
        Exit Function
    End If
    
    Busy True, Me
    
    If UCase$(Left$(pstrCiteria, 1)) = "M" Then
        pstrCiteria = Right$(pstrCiteria, Len(pstrCiteria) - 1)
    End If
    If Val(pstrCiteria) = 0 Then
        Busy False, Me
        ShowStatus 0
        Exit Function
    End If
    ShowStatus 8
        
    pobjList.Clear
    pobjList.BackColor = vbWindowBackground
    
    'Converted table names to constants
    Set ltabCustAccount = gdatCentralDatabase.OpenRecordset(gtblCustAccounts)
    'ltabCustAccount.Index = "CustAcctSel"
    ltabCustAccount.Index = "PrimaryKey"
    
    With ltabCustAccount
        .Seek "=", pstrCiteria
        If .NoMatch Then
           
            'MsgBox "No match"
            ltabCustAccount.Close
            Busy False, Me
            ShowStatus 0
            Exit Function
        End If
        llngRecCount = 0
        
        If Not .NoMatch Then
            llngRecCount = llngRecCount + 1
            ReDim Preserve plngCustNumber(llngRecCount)
            ReDim Preserve pstrCustName(llngRecCount)
            ReDim Preserve pstrInUseBy(llngRecCount)
            ReDim Preserve pstrDBIndicator(llngRecCount)
                        
            'If pobjOptionList(0).Value Or pobjOptionList(3).Value Then
            If Trim$(.Fields("DBIndicator") & "") = "O" Then
                pobjList.AddItem Trim$(.Fields("Surname") & "") & " " & _
                    Trim$(.Fields("Initials") & "") & ", " & Trim$(.Fields("Add1") & "") & _
                    ", " & Trim$(.Fields("Postcode") & "") & _
                    " (OLD)" & _
                    IIf(Trim$(.Fields("AcctInUseByFlag") & _
                    "") = "", "", " - " & UCase$(Trim$(.Fields("AcctInUseByFlag") & "")))
            Else
                pobjList.AddItem Trim$(.Fields("Surname") & "") & " " & _
                    Trim$(.Fields("Initials") & "") & ", " & Trim$(.Fields("Add1") & "") & _
                    ", " & Trim$(.Fields("Postcode") & "") & _
                    " (M" & Trim$(.Fields("CustNum") & "") & ")" & _
                    IIf(Trim$(.Fields("AcctInUseByFlag") & _
                    "") = "", "", " - " & UCase$(Trim$(.Fields("AcctInUseByFlag") & "")))
            End If
            plngCustNumber(llngRecCount - 1) = ValNull(.Fields("CustNum"))
            pstrCustName(llngRecCount - 1) = Trim$(.Fields("Surname") & "") & " " & .Fields("Initials") & ""
            pstrInUseBy(llngRecCount - 1) = Trim$(.Fields("AcctInUseByFlag") & "")
            pstrDBIndicator(llngRecCount - 1) = Trim$(.Fields("DBIndicator") & "")
        End If
        'pobjList.AddItem ""
    End With
    
    If llngRecCount = 0 Then
        'pobjList.AddItem ""
        'pobjList.BackColor = vbActiveBorder
        ReDim plngCustNumber(0)
        plngCustNumber(0) = 0
        Busy False, Me
    ElseIf llngRecCount = 1 Then
        pobjList.Selected(0) = True
        Busy False, Me
        DoEvents
        cmdSelect_Click
        DoEvents
    Else
        Busy False, Me
        ' On Error Resume Next
        ' pobjList.SetFocus
        ' On Error GoTo ErrHandler
    End If
    
    ltabCustAccount.Close
    'sbStatusBar.Panels(1).Text = StatusText(0)
    ShowStatus 0 ' 
Exit Function
ErrHandler:
    
    Busy False, Me
    
    Select Case GlobalErrorHandler(Err.Number, "frmCustAcctSel.FillCustAccountsListCustNumIndexed", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        'sbStatusBar.Panels(1).Text = StatusText(0)
        ShowStatus 0 ' 
        Exit Function
    Case Else
        Resume Next
    End Select


End Function

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With cmdSelect
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdOverideNSelect
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdSelect.Left - (cmdOverideNSelect.Width + 120)
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdOverideNSelect.Left - (cmdBack.Width + 120)
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With lstAccounts
        If Me.Width > .Left Then
            .Width = (Me.Width - .Left) - 213
        Else
            .Width = (.Left - Me.Width) - 213
        End If
        '.Height = ((Me.Height - cmdSelect.Top) - .Top) - 120
        If .Top < cmdHelp.Top Then
            .Height = (cmdHelp.Top - .Top) - 420 '305 '120
        Else
            .Height = (.Top - cmdHelp.Top) - 420 '305 '120
        End If
    End With
    
    With lblCustomerCount
        .Top = cmdBack.Top - 395
        .Width = (Me.Width - .Left) - 213
    End With
    
End Sub

Private Sub lstAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdSelect_Click
    End If

End Sub

Private Sub optSearchField_Click(Index As Integer)

    'txtSearchCriteria.SetFocus

End Sub

Private Sub optSearchField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'txtSearchCriteria.SetFocus

End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub txtSearchCriteria_KeyDown(KeyCode As Integer, Shift As Integer)

    'If KeyCode = 13 Then 'Carriage return
    '    DoEvents
    '    cmdFind_Click
    '    Me.Refresh
    '    DoEvents
    'End If

End Sub
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Private Sub txtCustNum_GotFocus()

    SetSelected Me
    
End Sub
Private Sub txtCustNum_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        DoEvents
        cmdFind_Click (0)
        Me.Refresh
        DoEvents
    End If
    
End Sub
Private Sub txtPostCode_GotFocus()

    SetSelected Me

End Sub
Private Sub txtPostcode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        DoEvents
        cmdFind_Click (2)
        Me.Refresh
        DoEvents
    End If
    
End Sub
Private Sub txtSurname_GotFocus()

    SetSelected Me

End Sub
Private Sub txtSurname_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        DoEvents
        cmdFind_Click (1)
        Me.Refresh
        DoEvents
    End If
    
End Sub
Sub SetupHelpFileReqs()

    'If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
    '    cmdHelpWhat.Visible = False
    '    Exit Sub
    'End If
    
    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/CustAcctSel.xml>WhatsScreen"
    
    'Me.HelpContextID = IDH_CAS_MAIN
    ctlBanner1.WhatsThisHelpID = IDH_CAS_MAIN
    ctlBanner1.WhatIsID = IDH_CAS_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_CAS_MAIN
    
    txtCustNum.WhatsThisHelpID = IDH_CAS_TXTCNUM
    cmdFind(0).WhatsThisHelpID = IDH_CAS_FIND
    txtSurname.WhatsThisHelpID = IDH_CAS_TXTSURN
    cmdFind(1).WhatsThisHelpID = IDH_CAS_FIND
    txtPostcode.WhatsThisHelpID = IDH_CAS_TXTPOST
    cmdFind(2).WhatsThisHelpID = IDH_CAS_FIND
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdOverideNSelect.WhatsThisHelpID = IDH_CAS_ADDNEW
    cmdSelect.WhatsThisHelpID = IDH_CAS_SELECT
    lstAccounts.WhatsThisHelpID = IDH_CAS_LSTACCT
    
End Sub

