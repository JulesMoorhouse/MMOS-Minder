VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmOrdHistory 
   Caption         =   "Select an order"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefundInfo 
      Caption         =   "&Refund Info"
      Height          =   360
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.CommandButton cmdViewConsign 
      Caption         =   "&View Consignment"
      Height          =   360
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewAdviceNote 
      Caption         =   "&View Advice Note"
      Height          =   360
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1543
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10515
      _extentx        =   18547
      _extenty        =   1852
   End
   Begin VB.CommandButton cmdCashBook 
      Caption         =   "&Cash Book"
      Height          =   360
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1305
   End
   Begin VB.CommandButton cmdInternalNote 
      Caption         =   "&Internal Note"
      Height          =   360
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7235
      Visible         =   0   'False
      Width           =   1543
   End
   Begin VB.CommandButton cmdConsignNote 
      Caption         =   "&Consignment Note"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7235
      Width           =   1543
   End
   Begin VB.CommandButton cmdNotePad 
      Caption         =   "&Notepad"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7235
      Width           =   1305
   End
   Begin VB.Data datCheques 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   7320
      Top             =   7080
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1305
   End
   Begin MSDBGrid.DBGrid dbgOrderLines 
      Bindings        =   "OrdHist.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "OrdHist.frx":001C
      TabIndex        =   6
      Top             =   4200
      Width           =   10215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Close"
      Height          =   360
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7235
      Width           =   1305
   End
   Begin MSDBGrid.DBGrid dbgAdviceNotes 
      Bindings        =   "OrdHist.frx":1749
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "OrdHist.frx":1766
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
   End
   Begin VB.Data datAdviceNotes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   13
      Top             =   7080
      Width           =   10515
      _extentx        =   18547
      _extenty        =   1244
   End
   Begin VB.Data datOrderLines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblBalance 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8040
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
   End
End
Attribute VB_Name = "frmOrdHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrRoute As String
Dim lstrScreenHelpFile As String

Public Sub cmdBack_Click()

    ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    Unload Me
    ClearAdviceNote
    ClearCustomerAcount
    ClearGen
    frmAbout.Show
    
End Sub
Private Sub cmdClose_Click()
    
    ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    Unload Me
    frmAbout.Show
    
End Sub


Private Sub cmdCashBook_Click()
    
    Busy True, Me
    frmChildCashbook.Route = gconstrCashbookSpecificCustomer
    Busy False, Me
    frmChildCashbook.Show vbModal
    
End Sub

Private Sub cmdConsignNote_Click()
Dim lstrLockingFlag As String

    lstrLockingFlag = LockingPhaseGen(True)
    
    gstrConsignmentNote.strType = "Consignment"

    On Error GoTo 0
    gstrAdviceNoteOrder.lngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
    gstrAdviceNoteOrder.lngConsignRemarkNum = CLng(datAdviceNotes.Recordset("ConsignRemarkNum"))
    On Error Resume Next
    
    If gstrAdviceNoteOrder.lngOrderNum = 0 Then
        MsgBox "You must select an order!", vbInformation, gconstrTitlPrefix & "Consignment Note"
        Exit Sub
    End If
    
    If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 Then
        If gstrConsignmentNote.lngRemarkNumber <> 0 Then
            frmChildNote.NoteText = gstrConsignmentNote.strText
            frmChildNote.NoteType = "Consignment Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        Else
            gstrConsignmentNote.lngRemarkNumber = gstrAdviceNoteOrder.lngConsignRemarkNum
            GetRemark gstrConsignmentNote.lngRemarkNumber, gstrConsignmentNote
            gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
            frmChildNote.NoteText = gstrConsignmentNote.strText
            frmChildNote.NoteType = "Consignment Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        End If
    Else
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrConsignmentNote
        gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngConsignRemarkNum, "Consignment"
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False
        frmChildNote.NoteText = gstrConsignmentNote.strText
        frmChildNote.NoteType = "Consignment Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrConsignmentNote.strText = frmChildNote.NoteText
    gstrConsignmentNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText

    gstrAdviceNoteOrder.lngConsignRemarkNum = 0
    gstrConsignmentNote.lngRemarkNumber = 0

End Sub

Private Sub cmdHelp_Click()
    
    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdInternalNote_Click()
Dim lstrLockingFlag As String
    
    lstrLockingFlag = LockingPhaseGen(True)
    
    gstrInternalNote.strType = "Internal"

    gstrAdviceNoteOrder.lngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))

    gstrAdviceNoteOrder.lngAdviceRemarkNum = CLng(datAdviceNotes.Recordset("AdviceRemarkNum"))

    
    If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 Then
        If gstrInternalNote.lngRemarkNumber <> 0 Then
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        Else
            gstrInternalNote.lngRemarkNumber = gstrAdviceNoteOrder.lngAdviceRemarkNum
            GetRemark gstrInternalNote.lngRemarkNumber, gstrInternalNote
            frmChildNote.NoteText = gstrInternalNote.strText
            frmChildNote.NoteType = "Internal Note Comments"
            Load frmChildNote
            frmChildNote.Show vbModal
        End If
    Else
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrInternalNote
        gstrAdviceNoteOrder.lngAdviceRemarkNum = gstrInternalNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngAdviceRemarkNum, "Internal"
        frmChildNote.NoteText = gstrInternalNote.strText
        frmChildNote.NoteType = "Internal Note Comments"
        Load frmChildNote
        frmChildNote.Show vbModal
    End If
    
    gstrInternalNote.strText = frmChildNote.NoteText
    gstrInternalNote.strType = frmChildNote.NoteType
    UpdateRemark gstrAdviceNoteOrder.lngAdviceRemarkNum, gstrInternalNote.strType, gstrInternalNote.strText

    gstrAdviceNoteOrder.lngAdviceRemarkNum = 0
    gstrInternalNote.lngRemarkNumber = 0
    
End Sub
Private Sub cmdModify_Click()
Dim lintRetVal As Variant
Dim lstrListCodeVars As ListVars

    If Val(datAdviceNotes.Recordset("OrderNum")) = 0 Then
        MsgBox "You Must Select and AdviceNote Line", , gconstrTitlPrefix & "Modify Order"
        Exit Sub
    End If
    
    If Val(datAdviceNotes.Recordset("CustNum")) = 0 Then
        MsgBox "You Must Select and AdviceNote Line", , gconstrTitlPrefix & "Modify Order"
        Exit Sub
    End If
    
    lstrListCodeVars.strListName = "Order Status"
    lstrListCodeVars.strListCode = Trim$(dbgAdviceNotes.Columns(1).Value)
    GetListVarsAll lstrListCodeVars

    If Trim$(lstrListCodeVars.strUserDef1) <> "CAN CANCEL" Then
        MsgBox "You may not modify this order!  This order has been processed or despatched!", , gconstrTitlPrefix & "Order Status Change"
        Exit Sub
    End If

    lintRetVal = MsgBox("Do you wish to modify Order Number " & _
        datAdviceNotes.Recordset("OrderNum") & _
        " Account Number " & datAdviceNotes.Recordset("CustNum"), vbYesNo, gconstrTitlPrefix & "Modify Order")
        
    If lintRetVal = vbYes Then
        gstrCustomerAccount.lngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
        gstrAdviceNoteOrder.lngCustNum = CLng(datAdviceNotes.Recordset("CustNum"))
        gstrAdviceNoteOrder.lngOrderNum = CLng(datAdviceNotes.Recordset("OrderNum"))
        GetAdviceNote gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum
                
        gstrOrderEntryOrderStatus = ""
        
        If gstrReferenceInfo.strDenomination <> gstrAdviceNoteOrder.strDenom Then
            MsgBox "You may not modify this order, as it was entered using a currency that is not in use!", , gconstrTitlPrefix & "Regional Settings"
            ClearCustomerAcount
            ClearAdviceNote
            ClearGen
            Exit Sub
        End If
        
        Unload Me
        
        Set gstrCurrentLoadedForm = frmAccount
        frmAccount.Route = gconstrOrderModify
        frmAccount.Show
    End If

End Sub

Private Sub cmdNotePad_Click()

    frmChildCuNotes.Show vbModal
    
End Sub

Private Sub cmdRefundInfo_Click()
Dim llngOrderNum As Long


    On Error Resume Next
    llngOrderNum = datAdviceNotes.Recordset("OrderNum")
    On Error GoTo 0
    If llngOrderNum = 0 Then
        MsgBox "This does not appear to be an Order Number, please select from grid!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
    
    GetRefundInfo llngOrderNum
    
End Sub

Private Sub cmdViewAdviceNote_Click()
Dim llngOrderNum As Long

    Debug.Print "START: " & Now()
    
    On Error Resume Next
    llngOrderNum = datAdviceNotes.Recordset("OrderNum")
    On Error GoTo 0
    If llngOrderNum = 0 Then
        MsgBox "This does not appear to be an Order Number, please select from grid!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If

    If UCase$(App.ProductName) = "LITE" Then
        ChooseLayout ltAdviceWithAddress, Me
    Else
        ChooseLayout ltAdviceNote, Me
    End If
    
    
    Busy True, Me
    
    Debug.Print "B4DAT: " & Now
    
    PrintObjAdviceNotesGeneral 0, 0, "S", llngOrderNum, , datAdviceNotes.Recordset("OrderStatus")
    Debug.Print "AFDAT: " & Now
    
    Busy False, Me
    
    Debug.Print "B4PRE: " & Now()
    
    ShowPlotReport
    Debug.Print "AFPRE: " & Now
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdViewConsign_Click()
Dim llngOrderNum As Long


    On Error Resume Next
    llngOrderNum = datAdviceNotes.Recordset("OrderNum")
    On Error GoTo 0
    If llngOrderNum = 0 Then
        MsgBox "This does not appear to be an Order Number, please select from grid!", , gconstrTitlPrefix & "Mandatory Field"
        Exit Sub
    End If
    
    GetConsignmentInfo llngOrderNum
    
End Sub

Private Sub datAdviceNotes_Reposition()
Dim lstrSQL As String

    Me.Refresh
    
    On Error Resume Next
    If Not (datAdviceNotes.Recordset.BOF = True And datAdviceNotes.Recordset.EOF = True) Then
        If Not IsNull(datAdviceNotes.Recordset("OrderNum")) Then
            lstrSQL = "Select CatNum, ItemDescription as Description, Qty, DespQty, " & _
                "trim(Denom) & format(Price,'0.00') as UnitPrice, " & _
                "trim(Denom) & format(Vat,'0.00') as Vat2, TaxCode, " & _
                "trim(Denom) & format(TotalPrice,'0.00') as TP, " & _
                "OrderLineNum as LineNum, ParcelNumber "

            lstrSQL = lstrSQL & " from " & gtblMasterOrderLines & " where CustNum=" & datAdviceNotes.Recordset("CustNum")
    
            lstrSQL = lstrSQL & " and OrderNum=" & datAdviceNotes.Recordset("OrderNum") & _
                " order by OrderLineNum, OrderNum"
                
            datOrderLines.RecordSource = lstrSQL
            datOrderLines.Refresh
            
            lstrSQL = "select * from " & gtblCashBook & " where OrderNum=" & datAdviceNotes.Recordset("OrderNum")
            datCheques.RecordSource = lstrSQL
            datCheques.Refresh
            
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub dbgAdviceNotes_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

    'And Check to see if anyone else is using it, if there is deny this update

End Sub

Private Sub dbgAdviceNotes_ButtonClick(ByVal ColIndex As Integer)
Dim lstrListCodeVars As ListVars

    frmChildOptions.List = ""
    
    Select Case ColIndex
    Case 1
        'Check current order status features
        lstrListCodeVars.strListName = "Order Status"
        lstrListCodeVars.strListCode = dbgAdviceNotes.Columns(ColIndex).Value
        GetListVarsAll lstrListCodeVars
        frmChildOptions.List = "Order Status"
    End Select
    
    If frmChildOptions.List <> "" Then
        frmChildOptions.Code = dbgAdviceNotes.Columns(ColIndex).Value
        frmChildOptions.Show vbModal

        If ColIndex = 1 And Trim$(lstrListCodeVars.strUserDef1) <> "CAN CANCEL" Then
            MsgBox "You may not change the order status of this order!" & vbCrLf & vbCrLf & _
            "This order has processed or despatched!", , gconstrTitlPrefix & "Order Status Change"
            Exit Sub
        End If
        
        If ColIndex = 1 Then
            lstrListCodeVars.strListName = "Order Status"
            lstrListCodeVars.strListCode = frmChildOptions.Code
            GetListVarsAll lstrListCodeVars
            If Trim$(lstrListCodeVars.strUserDef2) <> "USER USE" Then
                MsgBox "You may not change the order status of this order to status '" & _
                frmChildOptions.Code & "'!" & vbCrLf & vbCrLf & _
                "This order status is system assignable only!", , gconstrTitlPrefix & "Order Status Change"
                Exit Sub
            End If
        End If
        
        dbgAdviceNotes.Columns(ColIndex).Value = frmChildOptions.Code
    End If

End Sub

Private Sub Form_Load()
Dim lstrSQL As String

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
        
    Select Case gstrUserMode
    Case gconstrTestingMode
        datAdviceNotes.DatabaseName = gstrStatic.strCentralTestingDBFile
        datOrderLines.DatabaseName = gstrStatic.strCentralTestingDBFile
    Case gconstrLiveMode
        datAdviceNotes.DatabaseName = gstrStatic.strCentralDBFile
        datOrderLines.DatabaseName = gstrStatic.strCentralDBFile
    End Select
    
    If gstrSystemRoute <> srCompanyRoute Then
        datAdviceNotes.Connect = gstrDBPasswords.strCentralDBPasswordString
        datOrderLines.Connect = gstrDBPasswords.strCentralDBPasswordString
    End If
    
    lstrSQL = "SELECT CustNum, OrderNum, OrderStatus, DespatchDate, " & _
        "DeliveryAdd1, DeliveryPostcode, Trim(format([Denom],'0.00')) & " & _
        "format([Donation],'0.00') AS Donat, Trim(format([Denom],'0.00')) & " & _
        "format([Payment],'0.00') AS Pay1, Trim(format([Denom],'0.00')) & " & _
        "format([Payment2],'0.00') AS Pay2, Trim(format([Denom],'0.00')) & " & _
        "format([Underpayment],'0.00') AS UndP, Trim(format([Denom],'0.00')) & " & _
        "format([Reconcilliation],'0.00') AS Recon, Trim(format([Denom],'0.00')) & " & _
        "format([Postage],'0.00') AS Post, Trim(format([Denom],'0.00')) & " & _
        "format([Vat],'0.00') AS TaxVat, Trim(format([Denom],'0.00')) & " & _
        "format([TotalIncVat],'0.00') AS Total, ProcessedBy, CreationDate, " & _
        "ConsignRemarkNum, OrderType, PaymentType2, TelephoneNum, CardNumber, CardType "

    If gstrCustomerAccount.lngCustNum = 0 Then
        lstrSQL = lstrSQL & "FROM " & gtblAdviceNotes & " order by CreationDate, orderNum"
    Else
        lstrSQL = lstrSQL & "FROM " & gtblAdviceNotes & " where CustNum = " & _
            gstrCustomerAccount.lngCustNum & " order by CreationDate, orderNum"
    End If
    
    datAdviceNotes.RecordSource = lstrSQL
    
    If gstrReferenceInfo.booDonationAvail = False Then
        dbgAdviceNotes.Columns(5).Visible = False
    End If
    
    If UCase$(App.ProductName) = "LITE" Then
        'Doesn't really make much sense having this in the lite version!
        cmdViewConsign.Visible = False
    End If
    
    lblBalance = "Account balance = " & AccountBalance(gstrCustomerAccount.lngCustNum)
    ShowBanner Me, Me.Route
    
    SetupHelpFileReqs
    
End Sub

Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property

Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub

Private Sub Form_Resize()
Dim lintHeightOnePercent As Integer

    lintHeightOnePercent = Me.Height / 100
    
    On Error Resume Next
    With cmdBack
        .Top = Me.Height - gconlongButtonTop
        .Left = Me.Width - 1545
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdNotePad
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdConsignNote
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNotePad.Width + cmdNotePad.Left + 120
    End With
    
    With cmdInternalNote
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdConsignNote.Width + cmdConsignNote.Left + 120
    End With
    
    With dbgOrderLines
        .Width = Me.Width - 360
        .Top = cmdModify.Top + cmdModify.Height + 120
        .Height = ((cmdHelp.Top - .Top) - 120) - 305
    End With
    
    With cmdModify
        .Top = (dbgOrderLines.Top - .Height) - 120
    End With

    With cmdCashBook
        .Top = (dbgOrderLines.Top - .Height) - 120
    End With
    
    With cmdViewAdviceNote
        .Top = (dbgOrderLines.Top - .Height) - 120
    End With
    
    With dbgAdviceNotes
        .Width = Me.Width - 360
        .Height = (cmdModify.Top - .Top) - 120
    End With
End Sub

Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Sub GetConsignmentInfo(plngOrderNumber As Long)
Dim lstrSQL As String
Dim lstrMsg As String
Dim lsnaLists As Recordset
Dim lstrPFStatusCodes() As ListDetails
Dim lstrPFServiceCodes() As ListDetails
Dim lstrPFNotifyCodes() As ListDetails
Dim lstrPFPrepaidCodes() As ListDetails
Dim lstrPFWeekendCodes() As ListDetails
Dim lstrVarious As String
Dim lstrSpecial As String

    GetListDetailToArray "Consignment Status", lstrPFStatusCodes()
    GetListDetailToArray "PForce Service Indicator", lstrPFServiceCodes()
    GetListDetailToArray "PForce Notification Code", lstrPFNotifyCodes()
    GetListDetailToArray "PForce Prepaid Indicator", lstrPFPrepaidCodes()
    GetListDetailToArray "PForce Weekend Handling Code", lstrPFWeekendCodes()
    
    lstrSQL = "SELECT * From " & gtblPForce & " WHERE (((CustNum)=" & _
        gstrCustomerAccount.lngCustNum & ") AND ((OrderNum)=" & plngOrderNumber & "));"

    On Error GoTo ErrHandler
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        Do Until .EOF
            If lstrMsg <> "" Then
                lstrMsg = lstrMsg & vbTab & vbTab & vbTab & vbTab & "==================" & vbCrLf & vbCrLf
            End If
            lstrMsg = lstrMsg & Pad(Me, 35, "CONSIGNMENT: " & .Fields("ConsignNum") & "    ")
            lstrMsg = lstrMsg & Pad(Me, 35, "STATUS: " & MatchListDetailArray(.Fields("Status"), lstrPFStatusCodes()))
            lstrMsg = lstrMsg & "SERVICE: " & MatchListDetailArray(.Fields("ServiceID"), lstrPFServiceCodes()) & vbCrLf
            
            lstrMsg = lstrMsg & Pad(Me, 35, "DESPACTH DATE: " & .Fields("DespatchDate"))
            lstrMsg = lstrMsg & Pad(Me, 35, "PARCEL ITEMS: " & .Fields("ParcelItems"))
            lstrMsg = lstrMsg & Pad(Me, 35, "GROSS WEIGHT: " & .Fields("GrossWeight")) & vbCrLf & vbCrLf
            
            If Trim$(.Fields("WeekendHandCode")) <> "" Then
                lstrVarious = Pad(Me, 35, "WKND HDLG: " & MatchListDetailArray(.Fields("WeekendHandCode"), lstrPFWeekendCodes()))
            End If
            If Trim$(.Fields("PrepaidInd")) <> "" Then
                lstrVarious = lstrVarious & Pad(Me, 35, "PREP IND: " & MatchListDetailArray(.Fields("PrepaidInd"), lstrPFPrepaidCodes()))
            End If

            If lstrVarious <> "" Then
                lstrMsg = lstrMsg & lstrVarious & vbCrLf
            End If
            lstrVarious = ""
            
            lstrMsg = lstrMsg & Pad(Me, 20, "DELIVER TO: ") & Trim$(Trim$(.Fields("DeliverySalutation")) & _
                " " & Trim$(.Fields("DeliverySurname")) & " " & _
                Trim$(.Fields("DeliveryInitials"))) & vbCrLf
                
            lstrMsg = lstrMsg & vbTab & .Fields("DeliveryAdd1")
            If Trim$(.Fields("DeliveryAdd2")) <> "" Then
                lstrMsg = lstrMsg & ",  " & .Fields("DeliveryAdd2")
            End If
            If Trim$(.Fields("DeliveryAdd3")) <> "" Then
                lstrMsg = lstrMsg & ",  " & .Fields("DeliveryAdd3")
            End If
            If Trim$(.Fields("DeliveryAdd4")) <> "" Then
                lstrMsg = lstrMsg & ",  " & .Fields("DeliveryAdd4")
            End If
            If Trim$(.Fields("DeliveryAdd5")) <> "" Then
                lstrMsg = lstrMsg & ",  " & .Fields("DeliveryAdd5")
            End If
            lstrMsg = lstrMsg & ",  " & .Fields("DeliveryPostcode") & vbCrLf '& vbCrLf
            
            If .Fields("SpecialSatDel") = "Y" Then
                lstrSpecial = "Saturday Delivery = " & .Fields("SpecialSatDel") & vbTab
            End If
            If .Fields("SpecialBookIn") = "Y" Then
                lstrSpecial = lstrSpecial & " Book In = " & .Fields("SpecialBookIn") & vbTab
            End If
            If .Fields("SpecialProof") = "Y" Then
                lstrSpecial = lstrSpecial & " Proof = " & .Fields("SpecialProof")
            End If
            If lstrSpecial <> "" Then
                lstrMsg = lstrMsg & vbCrLf & "SPECIAL: " & lstrSpecial
            End If
            lstrSpecial = ""
            
            lstrMsg = lstrMsg & vbCrLf '& vbCrLf
            .MoveNext
        Loop
    End With
    lsnaLists.Close

    If Trim$(lstrMsg) = "" Then
        lstrMsg = "No consignments found!"
    End If
    
    MsgBox lstrMsg, , gconstrTitlPrefix & "Consignment(s) Details.  Cust Num:  M" & gstrCustomerAccount.lngCustNum & " Order Num: " & plngOrderNumber

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetConsignmentInfo", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Sub GetRefundInfo(plngOrderNumber As Long)
Dim lstrSQL As String
Dim lstrMsg As String
Dim lsnaLists As Recordset
Dim lstrAmount As String
Dim lstrChequeRequestDate As String
Dim lstrChequePrintedDate As String
Dim lstrBankRepPrintDate As String

    lstrSQL = "SELECT OrderNum, TotalIncVat, RefundReason, ChequeRequestDate, OrderType, " & _
        "PaymentType2, Denom, ChequePrintedDate, BankRepPrintDate " & _
        "From AdviceNotes WHERE RefundOrignNum=" & plngOrderNumber & ";"

    On Error GoTo ErrHandler
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        Do Until .EOF
        
            If lstrMsg = "" Then
                lstrMsg = "The following refund(s) relate to order number " & plngOrderNumber & vbCrLf & vbCrLf

                lstrMsg = lstrMsg & ColLeveller(Me, 7, "OrderNum") & "" & _
                    ColLeveller(Me, 12, "RefundReason") & " " & _
                    ColLeveller(Me, 11, "RequestDate") & " " & _
                    ColLeveller(Me, 11, "ChqPrintedDate") & " " & _
                    ColLeveller(Me, 11, "BankRepPrintDate") & " " & _
                    ColLeveller(Me, 12, "Amount  ") & vbCrLf & _
                    ColLeveller(Me, 7, "=======") & "" & _
                    ColLeveller(Me, 12, "==========") & "" & _
                    ColLeveller(Me, 11, " ==========") & " " & _
                    ColLeveller(Me, 11, "===========") & " " & _
                    ColLeveller(Me, 11, "===========") & " " & _
                    ColLeveller(Me, 12, "======") & vbCrLf

            End If
            lstrChequeRequestDate = " ~ BLANK ~ "
            lstrChequePrintedDate = " ~ BLANK ~ "
            lstrBankRepPrintDate = " ~ BLANK ~ "
            
            lstrAmount = .Fields("Denom") & Format$(.Fields("TotalIncVat"), "0.00")

            If Not IsNull(.Fields("ChequeRequestDate")) Then
                lstrChequeRequestDate = .Fields("ChequeRequestDate")
            End If
            If Not IsNull(.Fields("ChequePrintedDate")) Then
                lstrChequePrintedDate = .Fields("ChequePrintedDate")
            End If
            If Not IsNull(.Fields("BankRepPrintDate")) And .Fields("BankRepPrintDate") <> "00:00:00" Then
                lstrBankRepPrintDate = .Fields("BankRepPrintDate")
            End If
            
            lstrMsg = lstrMsg & ColLeveller(Me, 7, .Fields("OrderNum")) & vbTab & _
                ColLeveller(Me, 12, .Fields("RefundReason")) & " " & _
                ColLeveller(Me, 11, lstrChequeRequestDate) & " " & _
                ColLeveller(Me, 11, lstrChequePrintedDate) & " " & _
                ColLeveller(Me, 11, lstrBankRepPrintDate) & " " & _
                ColLeveller(Me, 14, lstrAmount) & vbCrLf

            .MoveNext
        Loop
    End With
    lsnaLists.Close

    If Trim$(lstrMsg) = "" Then
        lstrMsg = "No refunds found!"
    End If
    
    MsgBox lstrMsg, , gconstrTitlPrefix & "Refund(s) Details.  Cust Num:  M" & gstrCustomerAccount.lngCustNum & " Order Num: " & plngOrderNumber

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetRefundInfo", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub

Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/OrderHistory.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_ORDMNT_MAIN
    ctlBanner1.WhatIsID = IDH_ORDMNT_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_ORDMNT_MAIN
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdBack.WhatsThisHelpID = IDH_STANDARD_BACK
    cmdNotePad.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    cmdConsignNote.WhatsThisHelpID = IDH_STANDARD_CONSNOTE
    
    cmdCashBook.WhatsThisHelpID = IDH_STANDARD_CHICASHBOOK
    cmdViewAdviceNote.WhatsThisHelpID = IDH_STANDARD_VWADNOT
    
    cmdModify.WhatsThisHelpID = IDH_ORDHIST_MODIFY
    cmdViewConsign.WhatsThisHelpID = IDH_ORDHIST_CONSIGNDETSVIEW
    dbgAdviceNotes.WhatsThisHelpID = IDH_ORDHIST_GRIDADV
    dbgOrderLines.WhatsThisHelpID = IDH_ORDHIST_GRIDOLM

End Sub

