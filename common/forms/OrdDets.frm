VERSION 5.00
Begin VB.Form frmOrdDetails 
   Caption         =   "Order Details"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdNotePad 
      Caption         =   "&Notepad"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort Order"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   6120
      Top             =   6720
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Order Details"
      Height          =   4455
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   10305
      Begin VB.Frame fraOrderDets2 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   5302
         TabIndex        =   32
         Top             =   360
         Width           =   3615
         Begin VB.ComboBox cboYearStart 
            Height          =   315
            Left            =   2640
            TabIndex        =   13
            Top             =   3120
            Width           =   975
         End
         Begin VB.ComboBox cboMonthsStart 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   3120
            Width           =   1215
         End
         Begin VB.ComboBox cboCourier 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Left            =   2640
            TabIndex        =   8
            Text            =   "cboYear"
            Top             =   2160
            Width           =   975
         End
         Begin VB.ComboBox cboMonths 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox chkOverSeas 
            Alignment       =   1  'Right Justify
            Caption         =   "Overseas Flag"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   3600
            Width           =   1455
         End
         Begin VB.ComboBox cboCardType 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label lblValidFrom 
            Alignment       =   1  'Right Justify
            Caption         =   "Valid From"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Delivery Type"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblExpiryDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblCardType 
            Alignment       =   1  'Right Justify
            Caption         =   "Card Type"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   2640
            Width           =   975
         End
      End
      Begin VB.TextBox txtIssueNumber 
         Height          =   288
         Left            =   1800
         TabIndex        =   11
         Top             =   3480
         Width           =   1812
      End
      Begin VB.TextBox txtCardHoldersName 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3000
         Width           =   3015
      End
      Begin VB.ComboBox cboOrderStyle 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cboOrderCode 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3960
         Width           =   3012
      End
      Begin VB.ComboBox cboPaymentType2 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   3012
      End
      Begin VB.ComboBox cboPaymentType 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   3012
      End
      Begin VB.TextBox txtCCNum 
         Height          =   288
         Left            =   1800
         TabIndex        =   6
         Top             =   2520
         Width           =   3015
      End
      Begin VB.ComboBox cboMedia 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3012
      End
      Begin VB.TextBox txtDeliveryDate 
         Height          =   288
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lblIssueNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Issue Number"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblCardname 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Holders Name   (as appears on card)"
         Height          =   495
         Left            =   0
         TabIndex        =   30
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Order Type"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Order Code"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblPaymentType2 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Type 2"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblPaymentType1 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Type 1"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblCCNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card Number"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Media Code"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Delivery Date"
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   360
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7235
      Width           =   1305
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   37
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
End
Attribute VB_Name = "frmOrdDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrOrderStyle() As String

Dim lstrMediaCode() As String
Dim lstrCourierCode() As String

Dim lstrMonthsCode() As String
Dim lstrYearCode() As String
Dim lstrOrderCode() As String
Dim lstrCardType() As String

Dim mstrRoute As String
Dim lstrScreenHelpFile As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
'Dim lstrTextBox As String


Private Sub cboCardType_Click()

    If UCase$(Trim$(cboCardType)) = "SWITCH" Then
        lblIssueNumber.Visible = True
        txtIssueNumber.Visible = True

        lblValidFrom.Visible = True
        cboYearStart.Visible = True
        cboMonthsStart.Visible = True
    Else
        lblIssueNumber.Visible = False
        txtIssueNumber.Visible = False

        lblValidFrom.Visible = False
        cboYearStart.Visible = False
        cboMonthsStart.Visible = False
    End If
    
End Sub

Private Sub cboCardType_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        If txtIssueNumber.Visible Then
            txtIssueNumber.SetFocus
        Else
            cboOrderCode.SetFocus
        End If
    End If
    
End Sub

Private Sub cboCourier_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        If cboPaymentType.Visible = True Then
            cboPaymentType.SetFocus
        Else
            cboOrderCode.SetFocus
        End If
    End If
    
End Sub

Private Sub cboMedia_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliveryDate.SetFocus
    End If
    
End Sub

Private Sub cboMonths_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboYear.SetFocus
    End If
    
End Sub

Private Sub cboMonthsStart_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboYearStart.SetFocus
    End If
    
End Sub

Private Sub cboOrderCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        chkOverSeas.SetFocus
    End If
    
End Sub

Private Sub cboOrderStyle_Click()
Dim lbooVisible As Boolean

    If cboOrderStyle = "Replacement, No charge" Then
        gbooFreeOrder = True
        lbooVisible = False
    Else
        gbooFreeOrder = False
        lbooVisible = True
    End If
    
    cboPaymentType.Visible = lbooVisible
    cboPaymentType2.Visible = lbooVisible
    
    lblPaymentType1.Visible = lbooVisible
    lblPaymentType2.Visible = lbooVisible
    
    If cboPaymentType.Visible Then
        Call cboPaymentType_Click
    Else
    
        txtCardHoldersName.Visible = lbooVisible
        txtCCNum.Visible = lbooVisible
        cboCardType.Visible = lbooVisible
        cboMonths.Visible = lbooVisible
        cboYear.Visible = lbooVisible
        lblExpiryDate.Visible = lbooVisible
        lblCardType.Visible = lbooVisible
        lblPaymentType1.Visible = lbooVisible
        lblPaymentType2.Visible = lbooVisible
        lblCCNum.Visible = lbooVisible
        lblCardname.Visible = lbooVisible
        
        If cboCardType.Visible = True Then
            Call cboCardType_Click
        Else
            lblIssueNumber.Visible = lbooVisible
            txtIssueNumber.Visible = lbooVisible
            lblValidFrom.Visible = lbooVisible
            cboMonthsStart.Visible = lbooVisible
            cboYearStart.Visible = lbooVisible
        End If
    End If
    
End Sub

Private Sub cboOrderStyle_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboMedia.SetFocus
    End If
    
End Sub

Private Sub cboPaymentType_Click()
Dim lbooVisible As Boolean

    If Trim$(UCase$(cboPaymentType)) = "CREDIT CARD" Or Trim$(UCase$(cboPaymentType2)) = "CREDIT CARD" Then
        lbooVisible = True
    Else
        lbooVisible = False
    End If
    
    txtCardHoldersName.Visible = lbooVisible
    txtCCNum.Visible = lbooVisible
    cboCardType.Visible = lbooVisible
    cboMonths.Visible = lbooVisible
    cboYear.Visible = lbooVisible
    lblExpiryDate.Visible = lbooVisible
    lblCardType.Visible = lbooVisible
    lblCCNum.Visible = lbooVisible
    lblCardname.Visible = lbooVisible
    
    If cboCardType.Visible = True Then
        Call cboCardType_Click
    Else
        lblIssueNumber.Visible = lbooVisible
        txtIssueNumber.Visible = lbooVisible
        lblValidFrom.Visible = lbooVisible
        cboMonthsStart.Visible = lbooVisible
        cboYearStart.Visible = lbooVisible
    End If
End Sub

Private Sub cboPaymentType_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboPaymentType2.SetFocus
    End If
    
End Sub

Private Sub cboPaymentType2_Click()

    Call cboPaymentType_Click
    
End Sub

Private Sub cboPaymentType2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtCCNum.SetFocus
    End If
    
End Sub

Private Sub cboYear_GotFocus()

    SetSelected Me
    
End Sub

Private Sub cboYear_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtCardHoldersName.SetFocus
    End If
    
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)
    
End Sub

Private Sub cboYear_LostFocus()

    cboYear = Year("01/01/" & cboYear)
    
End Sub

Private Sub cboYearStart_GotFocus()

    SetSelected Me
    
End Sub

Private Sub cboYearStart_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboOrderCode.SetFocus
    End If
    
End Sub

Private Sub cboYearStart_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)
    
End Sub

Private Sub cboYearStart_LostFocus()

    If cboYearStart <> "" Then
        cboYearStart = Year("01/01/" & cboYearStart)
    End If
    
End Sub

Private Sub chkOverSeas_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdNext.SetFocus
    End If
    
End Sub

Private Sub cmdAbort_Click()
Dim lintRetVal As Variant
Dim lstrRetAbort As String
Dim lstrInUseByFlag As String

    frmChildAbortOptions.Style = 0
    frmChildAbortOptions.Show vbModal
    lstrRetAbort = frmChildAbortOptions.AbortOption
    Unload frmChildAbortOptions
    
    Select Case lstrRetAbort
    Case "BACK"
        Exit Sub
    Case "SAVE"
        'Will never see Abort button unless order entry
        lintRetVal = MsgBox("Do you wish to Save this order so far and set to cancelled?", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")
        If lintRetVal = vbYes Then
            SaveLocalFields
            lstrInUseByFlag = LockingPhaseGen(True)
            AddAdviceNote lstrInUseByFlag, "X"
            GetAdviceOrderNum lstrInUseByFlag, gstrAdviceNoteOrder.lngCustNum
                
            MsgBox "Your Order number is " & gstrAdviceNoteOrder.lngOrderNum, , gconstrTitlPrefix & "Order Incomplete!"
            
            UpdateSalesCode
        
            AppendOrderLinesToMaster gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
            UpdateMasterStockRecords gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum, Now()
            With gstrAdviceNoteOrder
                AddNewFileHistoryItem .lngCustNum, .lngOrderNum, _
                    Trim$(Trim$(gstrCustomerAccount.strSalutation) & " " & _
                    Trim$(gstrCustomerAccount.strInitials) & " " & Trim$(gstrCustomerAccount.strSurname)), _
                    "Aborted"
            End With
            
            ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
            ToggleAdviceInUseBy gstrAdviceNoteOrder.lngOrderNum, False
        Else
            Exit Sub
        End If
    Case "ABORT"
        lintRetVal = MsgBox("WARNING: By aborting this order no information will be saved!", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")
        If lintRetVal = vbYes Then
            ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
        Else
            Exit Sub
        End If
    End Select
    
    ClearCustomerAcount
    ClearAdviceNote
    ClearGen
    
    gintOrderLineNumber = 0
    
    Unload Me
    frmAbout.Show
    
End Sub

Private Sub cmdBack_Click()
Dim lintRetVal As Variant

    SaveLocalFields
    
    If Me.Route = gconstrOrderModify Then
        lintRetVal = MsgBox("WARNING! Information changed on this screen will not be save! " & vbCrLf & _
            "Proceed?", vbYesNo, gconstrTitlPrefix & "Warning")
        If lintRetVal = vbYes Then
            Set gstrCurrentLoadedForm = frmOrdHistory
            frmOrdHistory.Route = Me.Route
            frmOrdHistory.Show
        Else
            Exit Sub
        End If
    Else
        Set gstrCurrentLoadedForm = frmAccount
        frmAccount.Route = Me.Route
        frmAccount.Show
    End If
    
    Unload Me
    
End Sub

Private Sub cmdHelp_Click()
    
    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Public Sub cmdNext_Click()
    
    If Not IsDate(txtDeliveryDate) And txtDeliveryDate <> "" Then
        MsgBox "The Delivery Date you have entered does not appear to a date!" _
            , , gconstrTitlPrefix & "Discretionary Field"
        txtDeliveryDate.SetFocus
        Exit Sub
    End If
    
    If Trim$(txtCCNum) <> "" Then
        If Trim$(UCase$(cboPaymentType)) <> "CREDIT CARD" And _
            Trim$(UCase$(cboPaymentType2)) <> "CREDIT CARD" Then
            MsgBox "You have entered a Credit Card number.  You have not selected Credit card as the payment method?" _
            , , gconstrTitlPrefix & "Mandatory Field"
            cboPaymentType.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim$(UCase$(cboPaymentType)) = "CREDIT CARD" Or _
        Trim$(UCase$(cboPaymentType2)) = "CREDIT CARD" Then
        If Trim$(txtCCNum) = "" Then
            MsgBox "You have selected a payment type of credit card, without entering a credit card number!", , gconstrTitlPrefix & "Mandatory Field"
            txtCCNum.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim$(UCase$(cboPaymentType)) = "CREDIT CARD" And _
        Trim$(UCase$(cboPaymentType2)) = "CREDIT CARD" Then
            cboPaymentType2.ListIndex = -1
    End If
    
    If Trim$(cboOrderStyle.Text) <> "Replacement, No charge" Then
        If Trim$(UCase$(cboPaymentType)) = "" And _
            Trim$(UCase$(cboPaymentType2)) = "" Then
            MsgBox "You must enter at least one payment type!", , gconstrTitlPrefix & "Mandatory Field"
            cboPaymentType.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim$(cboMonthsStart) = "" Then
        If Trim$(cboYearStart) <> "" Then
            MsgBox "You have enter part of a Valid From date!", , gconstrTitlPrefix & "Mandatory Field"
            If cboMonthsStart.Visible = True Then
                cboMonthsStart.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    If Trim$(cboYearStart) = "" Then
        If Trim$(cboMonthsStart) <> "" Then
            MsgBox "You have enter part of a Valid From date!", , gconstrTitlPrefix & "Mandatory Field"
            If cboYearStart.Visible = True Then
                cboYearStart.SetFocus
            End If
            Exit Sub
        End If
    End If
    
    SaveLocalFields
    If Me.Route = gconstrOrderModify Then
        GetOrderLinesFromMaster
    End If
    Set gstrCurrentLoadedForm = frmOrder
    frmOrder.Route = Me.Route
    frmOrder.Show
    Unload Me
    
End Sub

Private Sub cmdNotePad_Click()

    frmChildCuNotes.Show vbModal
    
End Sub

Private Sub Form_Load()
Dim lstrCrap() As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    If gstrAdviceNoteOrder.lngCustNum = 0 Then
        gstrAdviceNoteOrder.lngCustNum = gstrCustomerAccount.lngCustNum
    End If
          
    With cboYear
        'Years changed
        .AddItem Year(DateAdd("YYYY", -5, Now()))
        .AddItem Year(DateAdd("YYYY", -4, Now()))
        .AddItem Year(DateAdd("YYYY", -3, Now()))
        .AddItem Year(DateAdd("YYYY", -2, Now()))
        .AddItem Year(DateAdd("YYYY", -1, Now()))
        .AddItem Year(Now)
        .AddItem Year(DateAdd("YYYY", 1, Now()))
        .AddItem Year(DateAdd("YYYY", 2, Now()))
        .AddItem Year(DateAdd("YYYY", 3, Now()))
        .AddItem Year(DateAdd("YYYY", 4, Now()))
        .AddItem Year(DateAdd("YYYY", 5, Now()))
    End With
            
    With cboYearStart
        'Years changed
        .AddItem ""
        .AddItem Year(DateAdd("YYYY", -5, Now()))
        .AddItem Year(DateAdd("YYYY", -4, Now()))
        .AddItem Year(DateAdd("YYYY", -3, Now()))
        .AddItem Year(DateAdd("YYYY", -2, Now()))
        .AddItem Year(DateAdd("YYYY", -1, Now()))
        .AddItem Year(Now)
        .AddItem Year(DateAdd("YYYY", 1, Now()))
        '.AddItem Year(DateAdd("YYYY", 2, Now()))
        '.AddItem Year(DateAdd("YYYY", 3, Now()))
        '.AddItem Year(DateAdd("YYYY", 4, Now()))
        '.AddItem Year(DateAdd("YYYY", 5, Now()))
    End With
            
    FillList "Order Type", cboOrderStyle, lstrOrderStyle()
    
    cboMonths.AddItem ""
    FillList "Months", cboMonths, lstrMonthsCode(), , , , True

    cboMonthsStart.AddItem ""
    FillList "Months", cboMonthsStart, lstrMonthsCode(), , , , True
    
    
    FillList "Media Codes", cboMedia, lstrMediaCode(), , , "CODE&DESC"
    FillList "PForce Service Indicator", cboCourier, lstrCourierCode()
    FillList "Payment Method", cboPaymentType, gstrPaymentTypeCode()
    FillList "Payment Method", cboPaymentType2, gstrPaymentTypeCode()

    FillList "Order Code", cboOrderCode, lstrOrderCode()

    FillList "Credit Card Type", cboCardType, lstrCardType()
    
    Select Case Me.Route
    Case gconstrAccount
        cmdAbort.Visible = False
        lblCardname = "Card Holders Name / Refund Cheque Name"
    Case gconstrEntry
        cmdAbort.Visible = True
    Case gconstrOrderModify, gconstrEnquiry
        cmdAbort.Visible = False
        lblCardname = "Card Holders Name / Refund Cheque Name"
    End Select

    ShowBanner Me, Me.Route
    
    GetLocalFields

End Sub
Sub SaveLocalFields()
    
    With gstrAdviceNoteOrder
    
        .strOrderStyle = Trim$(NotNull(cboOrderStyle, lstrOrderStyle))
        .strMediaCode = Trim$(NotNull(cboMedia, lstrMediaCode))
        If IsDate(txtDeliveryDate) Then
            .datDeliveryDate = CDate(Format$(txtDeliveryDate, "DD/Mmm/YYYY"))
        End If
        .strCourierCode = Trim$(NotNull(cboCourier, lstrCourierCode))
        If gbooFreeOrder = False Then
            .strPaymentType1 = Trim$(NotNull(cboPaymentType, gstrPaymentTypeCode))
            .strPaymentType2 = Trim$(NotNull(cboPaymentType2, gstrPaymentTypeCode))
            .strCardNumber = Trim$(txtCCNum)
            
            .lngIssueNumber = CLng(txtIssueNumber)
            .strCardType = Trim$(NotNull(cboCardType, lstrCardType))
            
            .datExpiryDate = Format("01/" & Val(Left$(cboMonths, 2)) & _
                " / " & Trim$(cboYear), "DD/MMM/YYYY")
            
            If Trim$(cboMonthsStart) <> "" Then
                .datCardStartDate = Format("01/" & Val(Left$(cboMonthsStart, 2)) & _
                    " / " & Trim$(cboYearStart), "DD/MMM/YYYY")
            Else
                .datCardStartDate = CDate(0)
            End If
            
            .strCardName = Trim$(txtCardHoldersName)
        ElseIf gbooFreeOrder = True Then
            .strPaymentType1 = "V"
            .strPaymentType2 = ""
            .strCardNumber = ""
            
            .lngIssueNumber = 0
            .strCardType = ""
            
            .datExpiryDate = CDate(0)
            .datCardStartDate = CDate(0)
            .strCardName = ""
        
        End If
        .strOrderCode = Trim$(NotNull(cboOrderCode, lstrOrderCode()))
        
        Select Case chkOverSeas.Value
        Case 0
            .strOverSeasFlag = "N"
            gstrVATRate = gstrReferenceInfo.strVATRate175
        Case 1
            .strOverSeasFlag = "Y"
            gstrVATRate = 0
        End Select

    End With

End Sub
Sub GetLocalFields()

    With gstrAdviceNoteOrder
        If Val(.strOrderStyle) = 0 Then
            .strOrderStyle = "1"
        End If
        SelectListItem Trim$(.strOrderStyle), cboOrderStyle, lstrOrderStyle()
        SelectListItem Trim$(.strMediaCode), cboMedia, lstrMediaCode()
        If Trim$(.datDeliveryDate) <> "00:00:00" Then
            txtDeliveryDate = Format(Trim$(.datDeliveryDate), "dd/Mmm/yyyy")
        Else
            txtDeliveryDate = ""
        End If
        If Trim$(.strCourierCode) = "" Or IsBlank(.strCourierCode) Or Trim$(.strCourierCode) = "PF 48" Then
            cboCourier.ListIndex = 0
        Else
            SelectListItem Trim$(.strCourierCode), cboCourier, lstrCourierCode()
        End If
        
        SelectListItem Trim$(.strPaymentType1), cboPaymentType, gstrPaymentTypeCode()
        
        SelectListItem Trim$(.strPaymentType2), cboPaymentType2, gstrPaymentTypeCode()
        
        txtCCNum = Trim$(.strCardNumber)
        
        If Trim$(.datExpiryDate) = "00:00:00" Then
            .datExpiryDate = Date
        End If
                
        SelectListItem Trim$(Month(.datExpiryDate)), cboMonths, lstrMonthsCode()
        
        On Error Resume Next
        If Year(Trim$(.datExpiryDate)) = "1899" Then
            cboYear = "2000"
        Else
            cboYear = Year(Trim$(.datExpiryDate))
        End If
        On Error GoTo 0
        
        If Trim$(.datCardStartDate) = "00:00:00" Then
        
        Else
            SelectListItem Trim$(Month(.datCardStartDate)), cboMonthsStart, lstrMonthsCode()
        End If
        
        On Error Resume Next
        If Year(Trim$(.datCardStartDate)) = "1899" Then
           
        Else
            cboYearStart = Year(Trim$(.datCardStartDate))
        End If
        On Error GoTo 0
        
        If Trim$(.strOrderCode) <> "" And Not IsBlank(.strOrderCode) Then
            SelectListItem Trim$(.strOrderCode), cboOrderCode, lstrOrderCode()
        Else
            SelectListItem "P", cboOrderCode, lstrOrderCode()
        End If

        If Trim$(.strCardName) = "" Or IsBlank(.strCardName) Then
            txtCardHoldersName = Trim$(Trim$(gstrCustomerAccount.strSalutation) & " " & Trim$(gstrCustomerAccount.strInitials) & " " & Trim$(gstrCustomerAccount.strSurname))
        Else
            txtCardHoldersName = Trim$(.strCardName)
        End If

        txtIssueNumber = CLng(.lngIssueNumber)
        SelectListItem Trim$(.strCardType), cboCardType, lstrCardType()
        
        If UCase$(Trim$(cboCardType)) = "SWITCH" Then
            lblIssueNumber.Visible = True
            txtIssueNumber.Visible = True

            lblValidFrom.Visible = True
            cboYearStart.Visible = True
            cboMonthsStart.Visible = True
        Else
            lblIssueNumber.Visible = False
            txtIssueNumber.Visible = False

            lblValidFrom.Visible = False
            cboYearStart.Visible = False
            cboMonthsStart.Visible = False
        End If

        Select Case UCase$(Trim$(.strOverSeasFlag))
        Case "Y"
            chkOverSeas.Value = 1 'Checked
        Case "N"
            chkOverSeas.Value = 0 'UnChecked
        End Select

    End With
    
End Sub

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim llngFormHalfWidth As Long

    On Error Resume Next
    
    llngFormHalfWidth = Me.Width / 2
    
    With cmdNext
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNext.Left - (cmdBack.Width + 120)
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdAbort
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdNotePad
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdAbort.Width + cmdAbort.Left + 120
    End With
        
    With Frame1
        .Width = Me.Width - (.Left * 2) - 60
    End With
    Dim llngMinLeft As Long
    llngMinLeft = cboOrderStyle.Left + cboOrderStyle.Width + 240
    With fraOrderDets2
        If llngFormHalfWidth < llngMinLeft Then
            .Left = llngMinLeft
        Else
            .Left = llngFormHalfWidth
        End If
    End With
        
End Sub

Private Sub Label3_Click()

End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub

Private Sub txtCardHoldersName_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtCardHoldersName_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboCardType.SetFocus
    End If
    
End Sub

Private Sub txtCardHoldersName_KeyPress(KeyAscii As Integer)
    
    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub

Private Sub txtCardHoldersName_LostFocus()

    txtCardHoldersName = ProperCase(txtCardHoldersName)

End Sub

Private Sub txtCCNum_GotFocus()

    SetSelected Me

End Sub

Private Sub txtCCNum_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboMonths.SetFocus
    End If
    
End Sub

Private Sub txtCCNum_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)

End Sub

Private Sub txtCCNum_LostFocus()

    txtCCNum = Trim$(txtCCNum)
    
End Sub


Private Sub txtDeliveryDate_GotFocus()

    ShowStatus 11
    
    SetSelected Me

End Sub

Private Sub txtDeliveryDate_KeyDown(KeyCode As Integer, Shift As Integer)

    txtDeliveryDate = CheckCalendar(KeyCode, txtDeliveryDate)

    If KeyCode = 13 Then 'Carriage return
        cboCourier.SetFocus
    End If
    
End Sub

Private Sub txtDeliveryDate_LostFocus()

    ShowStatus 0
    If IsDate(txtDeliveryDate) Then
        txtDeliveryDate = Format(Trim$(txtDeliveryDate), "Dd/Mmm/yyyy")
    End If

End Sub

Private Sub txtIssueNumber_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtIssueNumber_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboMonthsStart.SetFocus
    End If
    
End Sub

Private Sub txtIssueNumber_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidNum(KeyAscii)
    
End Sub

Private Sub txtIssueNumber_LostFocus()

    If Not IsNumeric(txtIssueNumber) Then
        txtIssueNumber = 0
    End If
    txtIssueNumber = Trim$(txtIssueNumber)
    
End Sub
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/OrderDetails.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_ORDDETS_MAIN
    ctlBanner1.WhatIsID = IDH_ORDDETS_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_ORDDETS_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdNext.WhatsThisHelpID = IDH_STANDARD_NEXT
    cmdAbort.WhatsThisHelpID = IDH_STANDARD_ABORT
    cmdNotePad.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    
    cboOrderStyle.WhatsThisHelpID = IDH_ORDDETS_ORDSTY
    cboMedia.WhatsThisHelpID = IDH_ORDDETS_MEDIA
    txtDeliveryDate.WhatsThisHelpID = IDH_ORDDETS_DELDATE
    cboCourier.WhatsThisHelpID = IDH_ORDDETS_COURIER
    cboPaymentType.WhatsThisHelpID = IDH_ORDDETS_PAYTY1
    cboPaymentType2.WhatsThisHelpID = IDH_ORDDETS_PAYTY2
    txtCCNum.WhatsThisHelpID = IDH_ORDDETS_CCNUM
    cboMonths.WhatsThisHelpID = IDH_ORDDETS_EXPMONTH
    cboYear.WhatsThisHelpID = IDH_ORDDETS_EXPYEAR
    txtCardHoldersName.WhatsThisHelpID = IDH_ORDDETS_CARDNAME
    cboCardType.WhatsThisHelpID = IDH_ORDDETS_CARDTYPE
    txtIssueNumber.WhatsThisHelpID = IDH_ORDDETS_CARDISS
    cboMonthsStart.WhatsThisHelpID = IDH_ORDDETS_VALMONTH
    cboYearStart.WhatsThisHelpID = IDH_ORDDETS_VALYEAR
    cboOrderCode.WhatsThisHelpID = IDH_ORDDETS_ORDCODE
    chkOverSeas.WhatsThisHelpID = IDH_ORDDETS_OVERSEAS
    
End Sub


