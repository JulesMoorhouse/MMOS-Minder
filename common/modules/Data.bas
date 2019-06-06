Attribute VB_Name = "modData"
Option Explicit
Sub AddProductSpecific(pstrCatNum As String, plngQty As Long, plngOrderLineNum As Long)
Dim lstrSQL As String

    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblProducts & " SET Selected = True, Qty = " & plngQty & _
        ", OrderLineNum = " & plngOrderLineNum & _
        " WHERE (((trim(ucase(CatNum)))='" & Trim$(UCase$(pstrCatNum)) & "'));"

    gdatLocalDatabase.Execute lstrSQL
    
Exit Sub
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "AddProductSpecific", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub
Sub ApndOrdLinesMstToLocal(plngCustomerNum As Long, plngOrderNum As Long)
Dim lstrSQL As String

    ShowStatus 26
    
    On Error GoTo ErrHandler

    lstrSQL = "Delete * from " & gtblOrderLines & ";"
    gdatLocalDatabase.Execute lstrSQL
        
    lstrSQL = "INSERT INTO " & gtblOrderLines & " ( CustNum, CatNum, ItemDescription, " & _
        "BinLocation, Qty, Price, Vat, Weight, TaxCode, TotalPrice, " & _
        "TotalWeight, Class, SalesCode, OrderLineNum, ParcelNumber ) SELECT " & gtblMasterOrderLines & _
        ".CustNum , " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & _
        ".ItemDescription, " & gtblMasterOrderLines & ".BinLocation, " & _
        gtblMasterOrderLines & ".Qty, " & gtblMasterOrderLines & ".Price, " & _
        gtblMasterOrderLines & ".Vat, " & gtblMasterOrderLines & _
        ".Weight, " & gtblMasterOrderLines & ".TaxCode, " & gtblMasterOrderLines & _
        ".TotalPrice, " & gtblMasterOrderLines & ".TotalWeight, " & _
        gtblMasterOrderLines & ".Class, " & gtblMasterOrderLines & _
        ".SalesCode, " & gtblMasterOrderLines & ".OrderLineNum, " & gtblMasterOrderLines & _
        ".ParcelNumber From " & gtblMasterOrderLines & _
        " WHERE (((" & gtblMasterOrderLines & ".CustNum)=" & plngCustomerNum & _
        ") AND ((" & gtblMasterOrderLines & ".OrderNum)=" & plngOrderNum & "));"

    gdatLocalDatabase.Execute lstrSQL
Exit Sub
ErrHandler:
    Select Case GlobalErrorHandler(Err.Number, "ApndOrdLinesMstToLocal", "Local", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub
Function GetCustomerAccount(plngCustNumber As Long, pbooLockRecord As Boolean) As Boolean
Dim ltabCustAccount As Recordset

    ShowStatus 19
    On Error GoTo ErrHandler

    If pbooLockRecord = True Then
        ToggleAcctInUseBy plngCustNumber, True
    End If

    Set ltabCustAccount = gdatCentralDatabase.OpenRecordset(gtblCustAccounts)
    ltabCustAccount.Index = "PrimaryKey"
    
    With ltabCustAccount
        .Seek "=", plngCustNumber
        
        If .NoMatch Then
            GetCustomerAccount = False
            Exit Function
        End If
        GetCustomerAccount = True
        
        gstrCustomerAccount.lngCustNum = Trim$(!CustNum)
        gstrCustomerAccount.strSurname = Trim$(!Surname & "")
        gstrCustomerAccount.strSalutation = Trim$(!Salutation & "")
        gstrCustomerAccount.strInitials = Trim$(!Initials & "")
        gstrCustomerAccount.strAdd1 = Trim$(!Add1 & "")
        gstrCustomerAccount.strAdd2 = Trim$(!Add2 & "")
        gstrCustomerAccount.strAdd3 = Trim$(!Add3 & "")
        gstrCustomerAccount.strAdd4 = Trim$(!Add4 & "")
        gstrCustomerAccount.strAdd5 = Trim$(!Add5 & "")
        gstrCustomerAccount.strPostcode = Trim$(!Postcode & "")
        gstrCustomerAccount.strDeliverySalutation = Trim$(!DeliverySalutation & "")
        gstrCustomerAccount.strDeliverySurname = Trim$(!DeliverySurname & "")
        gstrCustomerAccount.strDeliveryInitials = Trim$(!DeliveryInitials & "")
        
        gstrCustomerAccount.strTelephoneNum = Trim$(!TelephoneNum & "")
        gstrCustomerAccount.strEveTelephoneNum = Trim$(!EveTelephoneNum & "")
        
        gstrCustomerAccount.strDeliveryAdd1 = Trim$(!DeliveryAdd1 & "")
        gstrCustomerAccount.strDeliveryAdd2 = Trim$(!DeliveryAdd2 & "")
        gstrCustomerAccount.strDeliveryAdd3 = Trim$(!DeliveryAdd3 & "")
        gstrCustomerAccount.strDeliveryAdd4 = Trim$(!DeliveryAdd4 & "")
        gstrCustomerAccount.strDeliveryAdd5 = Trim$(!DeliveryAdd5 & "")
        gstrCustomerAccount.strDeliveryPostcode = Trim$(!DeliveryPostcode & "")
        gstrCustomerAccount.strAccountType = Trim$(!AccountType & "")
        gstrCustomerAccount.strReceiveMailings = Trim$(!ReceiveMailings & "")
        gstrCustomerAccount.strAcctInUseByFlag = Trim$(!AcctInUseByFlag & "")
        gstrCustomerAccount.strAccountStatus = Trim$(!AcctStatus & "")
        gstrCustomerAccount.strEmail = Trim$(!email & "")
                
        .Close

    End With

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetCustomerAccount", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub GetOrderLinesFromMaster()
Dim lstrSQL As String

    ShowStatus 42
    
    On Error GoTo ErrHandler
    
    lstrSQL = "Delete * from " & gtblOrderLines & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    lstrSQL = "INSERT INTO " & gtblOrderLines & " ( CustNum, CatNum, ItemDescription, " & _
        "BinLocation, Qty, Price, Vat, Weight, TaxCode, TotalPrice, " & _
        "TotalWeight, Class, SalesCode, OrderLineNum, ParcelNumber ) SELECT " & _
        gtblMasterOrderLines & ".CustNum, " & _
        gtblMasterOrderLines & ".CatNum, " & _
        gtblMasterOrderLines & ".ItemDescription, " & _
        gtblMasterOrderLines & ".BinLocation, " & _
        gtblMasterOrderLines & ".Qty, " & _
        gtblMasterOrderLines & ".Price, " & _
        gtblMasterOrderLines & ".Vat, " & _
        gtblMasterOrderLines & ".Weight, " & _
        gtblMasterOrderLines & ".TaxCode, " & _
        gtblMasterOrderLines & ".TotalPrice, " & _
        gtblMasterOrderLines & ".TotalWeight, " & _
        gtblMasterOrderLines & ".Class, " & _
        gtblMasterOrderLines & ".SalesCode, " & _
        gtblMasterOrderLines & ".OrderLineNum, " & _
        gtblMasterOrderLines & ".ParcelNumber " & _
        "From " & gtblMasterOrderLines & " " & _
        "WHERE (((" & gtblMasterOrderLines & ".CustNum)=" & gstrAdviceNoteOrder.lngCustNum & _
        ") AND ((" & gtblMasterOrderLines & ".OrderNum)=" & gstrAdviceNoteOrder.lngOrderNum & "));"
        
        gdatLocalDatabase.Execute lstrSQL
    
Exit Sub
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "GetOrderLinesFromMaster", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function UpdateLists(pobjForm As Form, Optional pbooNoStatus As Variant) As Boolean
Dim lstrSQL As String

    UpdateLists = True
    
    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 1
        DoEvents
    End If

    On Error GoTo ErrHandler
    
    lstrSQL = "Delete * from " & gtblListDetails & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    lstrSQL = "INSERT INTO " & gtblListDetails & "  SELECT " & gtblMasterListDetails & ".* " & _
        "FROM " & gtblMasterListDetails & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    If pbooNoStatus = False Then
        ShowStatus 2
        DoEvents
    End If
    
    lstrSQL = "Delete * from " & gtblLists & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    lstrSQL = "INSERT INTO " & gtblLists & " SELECT " & gtblMasterLists & ".* " & _
        "FROM " & gtblMasterLists & ";"
    gdatLocalDatabase.Execute lstrSQL
    
    If pbooNoStatus = False Then
        ShowStatus 3
        DoEvents
    End If
    
    'Update constant info previously in static, other that file info
    GetReferenceInfo True
    
    If Trim$(gstrReferenceInfo.strDenomination) <> Left$(FormatCurrency("0"), 1) Then
        MsgBox "Please contact your technical support advisor!" & vbCrLf & vbCrLf & _
            "Your 'Regional Currency Settings' do not match those setup for " & vbCrLf & _
            "the system as a whole.  " & vbCrLf & vbCrLf & "Your system Denomination is set to '" & _
            gstrReferenceInfo.strDenomination & "' and your PC is set to '" & _
            Left$(FormatCurrency("0"), 1) & "'", , gconstrTitlPrefix & "Regional Settings"
        
        UpdateLists = False
        Exit Function
    End If
        
    If NewStockAvailable(True) Then
        lstrSQL = "Delete * from " & gtblProducts & ";"
        gdatLocalDatabase.Execute lstrSQL
        
        lstrSQL = "INSERT INTO " & gtblProducts & " (CatNum, ItemDescription, " & _
            "BinLocation, Class, ClassItem, ClassGroup, Price, Weight, TaxCode, " & _
            "NumInStock) SELECT " & gtblMasterProducts & ".CatNum, " & _
            "" & gtblMasterProducts & ".ItemDescription, " & gtblMasterProducts & ".BinLocation, " & _
            "" & gtblMasterProducts & ".Class, " & gtblMasterProducts & ".ClassItem, " & gtblMasterProducts & ".ClassGroup, " & _
            "" & gtblMasterProducts & ".Price, " & gtblMasterProducts & ".Weight, " & gtblMasterProducts & ".TaxCode, " & _
            "" & gtblMasterProducts & ".NumInStock " & _
            "FROM " & gtblMasterProducts & ";"
        
        gdatLocalDatabase.Execute lstrSQL
        
        UpdateMachineStockFlag True
    End If
    
    If UCase$(App.ProductName) = "CLIENT" Then
        If False Then ' DEV NOTE 2019: Add Post Office Address file data NewPADAvailable(True) Then
            If pbooNoStatus = False Then ShowStatus 130
            gdatLocalDatabase.Execute "DELETE * FROM " & gtblPADAvailable & ";"
            gdatLocalDatabase.Execute "DELETE * FROM " & gtblPADOffice & ";"
            gdatLocalDatabase.Execute "DELETE * FROM " & gtblPADOpeningTimes & ";"
            
            If pbooNoStatus = False Then ShowStatus 131
            DoEvents
            gdatLocalDatabase.Execute "INSERT INTO " & gtblPADAvailable & _
                " SELECT * FROM " & gtblMasterPADAvailable & ";"
            If pbooNoStatus = False Then ShowStatus 132
            DoEvents
            gdatLocalDatabase.Execute "INSERT INTO " & gtblPADOffice & _
                " SELECT * FROM " & gtblMasterPADOffice & ";"
            If pbooNoStatus = False Then ShowStatus 133
            DoEvents
            gdatLocalDatabase.Execute "INSERT INTO " & gtblPADOpeningTimes & _
                " SELECT * FROM " & gtblMasterPADOpeningTimes & ";"
            DoEvents
            
            lstrSQL = "UPDATE " & gtblSystem & ", " & gtblMachine & " SET " & gtblMachine & ".[Value] = [" & gtblSystem & "].[Value] " & _
                "WHERE (((" & gtblSystem & ".Item)='PADUploaded') AND ((" & gtblMachine & ".Item)='PADDownloaded'));"
            gdatLocalDatabase.Execute lstrSQL
        End If
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 0
        DoEvents
    End If

Exit Function
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "UpdateLists", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Function
Sub UpdateOrderLinesToMaster(plngCustomerNum As Long, plngOrderNum As Long)
Dim lstrSQL As String

    ShowStatus 26
    On Error GoTo ErrHandler

    lstrSQL = "UPDATE " & gtblMasterOrderLines & " INNER JOIN " & _
        "" & gtblOrderLines & " ON (" & gtblMasterOrderLines & ".CatNum = " & _
        "" & gtblOrderLines & ".CatNum) AND (" & gtblMasterOrderLines & ".CustNum = " & _
        "" & gtblOrderLines & ".CustNum) SET " & gtblMasterOrderLines & ".Qty = " & _
        "[" & gtblOrderLines & "].[Qty], " & gtblMasterOrderLines & ".Price = " & _
        "[" & gtblOrderLines & "].[Price], " & gtblMasterOrderLines & ".Vat = " & _
        "[" & gtblOrderLines & "].[Vat], " & gtblMasterOrderLines & ".OrderLineNum =" & _
        "[" & gtblOrderLines & "].[OrderLineNum], " & gtblMasterOrderLines & ".ParcelNumber =" & _
        "[" & gtblOrderLines & "].[ParcelNumber], " & gtblMasterOrderLines & ".TotalPrice =" & _
        "[" & gtblOrderLines & "].[TotalPrice] WHERE (((" & gtblMasterOrderLines & _
        ".CustNum)=" & plngCustomerNum & _
        ") AND ((" & gtblMasterOrderLines & ".OrderNum)=" & plngOrderNum & "));"
        
    gdatLocalDatabase.Execute lstrSQL

    lstrSQL = "Delete * from " & gtblOrderLines & ";"
    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOrderLinesToMaster", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateOrdLinesWithProds(plngCustNumber As Long)
Dim lstrSQL As String

    ShowStatus 43

    On Error GoTo ErrHandler
    lstrSQL = "INSERT INTO " & gtblOrderLines & " ( CustNum, CatNum, ItemDescription, " & _
        "BinLocation, Qty, Price, Weight, TaxCode, Class, OrderLineNum ) " & _
        "SELECT " & plngCustNumber & " AS X, " & _
        "" & gtblProducts & ".CatNum, " & gtblProducts & ".ItemDescription, " & gtblProducts & ".BinLocation, " & _
        "" & gtblProducts & ".Qty, " & gtblProducts & ".Price, " & gtblProducts & ".Weight, " & gtblProducts & ".TaxCode , " & gtblProducts & ".Class, " & gtblProducts & ".OrderLineNum " & _
        "From " & gtblProducts & " WHERE (((" & gtblProducts & ".Selected)=True));"

    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOrdLinesWithProds", "Local", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateOrderLinesTotals()
Dim lstrSQL As String

    ShowStatus 44
    On Error GoTo ErrHandler

    lstrSQL = "UPDATE " & gtblOrderLines & " SET " & gtblOrderLines & ".TotalPrice = " & _
        "IIf(Trim$([TaxCode])='Z',([Qty]*[Price]),([Qty]*[Price])+" & _
        "([Qty]*[Price])*" & gstrVATRate / 100 & "), " & gtblOrderLines & ".TotalWeight = [Qty]*[Weight], " & _
        "" & gtblOrderLines & ".Vat = IIf(Trim$([TaxCode])='Z',0,([Qty]*[Price])*" & gstrVATRate / 100 & ");"

    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOrderLinesTotals", "Local", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateOrderLinesTotalsDespQty(plngOrderNum As Long, pbooOverSeas As Boolean)
Dim lstrSQL As String
Dim lstrKeptVAT As String

    ShowStatus 45
    
    On Error GoTo ErrHandler

    If pbooOverSeas = True Then
        lstrKeptVAT = 0
    ElseIf pbooOverSeas = False Then
        lstrKeptVAT = gstrVATRate
    End If
    
    lstrSQL = "UPDATE " & gtblMasterOrderLines & " SET " & gtblMasterOrderLines & ".TotalPrice = " & _
        "IIf(Trim$([TaxCode])='Z',([DespQty]*[Price]),([DespQty]*[Price])+" & _
        "([DespQty]*[Price])*" & lstrKeptVAT / 100 & "), " & gtblMasterOrderLines & ".TotalWeight = [DespQty]*[Weight], " & _
        "" & gtblMasterOrderLines & ".Vat = IIf(Trim$([TaxCode])='Z',0,([DespQty]*[Price])*" & _
        lstrKeptVAT / 100 & ") WHERE OrderNum = " & plngOrderNum & ";"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOrderLinesTotalsDespQty", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub

Sub AddNewCustomerAccount(pstrInUseByFlag As String)
Dim lstrSQL As String

    ShowStatus 20
    
    On Error GoTo ErrHandler

    lstrSQL = "INSERT INTO " & gtblCustAccounts & " ( AcctInUseByFlag ) select ('" & pstrInUseByFlag & "');"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddNewCustomerAccount", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub GetCustomerAccountNum(pstrInUseByFlag As String)
Dim ltabCustAccount As Recordset

    ShowStatus 21
    On Error GoTo ErrHandler

    Set ltabCustAccount = gdatCentralDatabase.OpenRecordset(gtblCustAccounts)
    ltabCustAccount.Index = "AcctInUseBy"
    
    With ltabCustAccount
        .Seek "=", pstrInUseByFlag
        
        gstrCustomerAccount.lngCustNum = Trim$(!CustNum)
        gstrCustomerAccount.strAcctInUseByFlag = Trim$(!AcctInUseByFlag & "")
            
        .Close

    End With

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetCustomerAccountNum", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select


End Sub
Sub GetAdviceOrderNum(pstrInUseByFlag As String, plngCustomerNum As Long)
Dim ltabAdviceNote As Recordset

    ShowStatus 22
    On Error GoTo ErrHandler

    Set ltabAdviceNote = gdatCentralDatabase.OpenRecordset(gtblAdviceNotes)
    ltabAdviceNote.Index = "AcctInUseBy"
    
    With ltabAdviceNote
        .Seek "=", pstrInUseByFlag, plngCustomerNum
        
        gstrAdviceNoteOrder.lngOrderNum = Trim$(!OrderNum)
        gstrAdviceNoteOrder.lngCustNum = Trim$(!CustNum)
        gstrAdviceNoteOrder.strAcctInUseByFlag = Trim$(!LockingFlag & "")
            
        .Close

    End With

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetAdviceOrderNum", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select


End Sub
Function GetAdviceCustName(plngOrderNum As Long) As String
Dim lstrSQL As String
Dim lsnaLists As Recordset

    ShowStatus 46
    On Error GoTo ErrHandler

    GetAdviceCustName = ""
    lstrSQL = "SELECT CallerSalutation, CallerSurname from " & gtblAdviceNotes & " where OrderNum = " & plngOrderNum & ";"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        Do Until .EOF
            GetAdviceCustName = Trim$(!CallerSalutation & "") & " " & Trim$(!CallerSurname & "")
            .MoveNext ' Won't happen
        Loop
    End With
    
    lsnaLists.Close


Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetAdviceCustName", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select


End Function
Function GetTotalCustomerUnderpayment(plngCustomerNum As Long) As Currency
Dim lstrSQL As String
Dim lsnaLists As Recordset
Dim lstrCashbook As String
Dim lvarTotalCustomerUnderpayment

    GetTotalCustomerUnderpayment = 0
    
    ShowStatus 47
    On Error GoTo ErrHandler
               
    lstrSQL = "SELECT Sum(UnderPayment) AS SumOfAmount From AdviceNotes " & _
        "WHERE (((CustNum)=" & plngCustomerNum & ") AND ((RefundReason)='UNDERPAY') " & _
        "AND ((BankRepPrintDate)<>0 And (BankRepPrintDate) Is Not Null));"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            lvarTotalCustomerUnderpayment = .Fields("SumOfAmount") & ""
        End If
    End With
    
    lsnaLists.Close
    If IsNull(lvarTotalCustomerUnderpayment) Or lvarTotalCustomerUnderpayment = "" Then
        GetTotalCustomerUnderpayment = 0
    Else
        GetTotalCustomerUnderpayment = CCur(lvarTotalCustomerUnderpayment)
    End If

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetTotalCustomerUnderpayment", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select


End Function
Sub UpdateAccount(plngCustomerNum As Long)
Dim lstrSQL As String

    ShowStatus 23
    On Error GoTo ErrHandler
    
    With gstrCustomerAccount
        lstrSQL = "UPDATE " & gtblCustAccounts & " SET " & _
            "Salutation = '" & OneSpace(JetSQLFixup(.strSalutation)) & "', " & _
            "Surname = '" & OneSpace(JetSQLFixup(.strSurname)) & "', " & _
            "Initials = '" & OneSpace(JetSQLFixup(.strInitials)) & "', " & _
            "Add1 = '" & OneSpace(JetSQLFixup(.strAdd1)) & "', " & "Add2 = '" & OneSpace(JetSQLFixup(.strAdd2)) & "', " & _
            "Add3 = '" & OneSpace(JetSQLFixup(.strAdd3)) & "', " & _
            "Add4 = '" & OneSpace(JetSQLFixup(.strAdd4)) & "', " & _
            "Add5 = '" & OneSpace(JetSQLFixup(.strAdd5)) & "'," & _
            "Postcode = '" & OneSpace(.strPostcode) & "', " & _
            "TelephoneNum = '" & OneSpace(.strTelephoneNum) & "', " & _
            "EveTelephoneNum = '" & OneSpace(.strEveTelephoneNum) & "', " & _
            "DeliverySalutation = '" & OneSpace(JetSQLFixup(.strDeliverySalutation)) & "', " & _
            "DeliverySurname = '" & OneSpace(JetSQLFixup(.strDeliverySurname)) & "', " & _
            "DeliveryInitials = '" & OneSpace(JetSQLFixup(.strDeliveryInitials)) & "', " & _
            "DeliveryAdd1 = '" & OneSpace(JetSQLFixup(.strDeliveryAdd1)) & "', " & _
            "DeliveryAdd2 = '" & OneSpace(JetSQLFixup(.strDeliveryAdd2)) & "', " & _
            "DeliveryAdd3 = '" & OneSpace(JetSQLFixup(.strDeliveryAdd3)) & "', " & _
            "DeliveryAdd4 = '" & OneSpace(JetSQLFixup(.strDeliveryAdd4)) & "', " & _
            "DeliveryAdd5 = '" & OneSpace(JetSQLFixup(.strDeliveryAdd5)) & "', " & _
            "DeliveryPostcode = '" & OneSpace(.strDeliveryPostcode) & "', " & _
            "AccountType = '" & OneSpace(.strAccountType) & "', " & _
            "AcctStatus = '" & OneSpace(.strAccountStatus) & "', " & _
            "Email = '" & OneSpace(.strEmail) & "', " & _
            "ReceiveMailings = '" & OneSpace(.strReceiveMailings) & "' " & _
            "WHERE (((CustNum)=" & plngCustomerNum & "));"

    End With
    
    gdatCentralDatabase.Execute lstrSQL
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateAccount", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Sub
Sub ToggleAcctInUseBy(plngCustomerNum As Long, pbooLockAccount As Boolean)
Dim lstrSQL As String
Dim lstrInUseByFlag As String

    On Error GoTo ErrHandler
    
    If pbooLockAccount = True Then
        ShowStatus 24
        lstrInUseByFlag = LockingPhaseGen(True)
    Else
        lstrInUseByFlag = " "
    End If

    lstrSQL = "UPDATE " & gtblCustAccounts & " SET AcctInUseByFlag = '" & lstrInUseByFlag & _
        "' WHERE (((" & gtblCustAccounts & ".CustNum)=" & plngCustomerNum & _
        "));"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "ToggleAcctInUseBy", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub ToggleAdviceInUseBy(plngOrderNum As Long, pbooLockAdvice As Boolean)
Dim lstrSQL As String
Dim lstrInUseByFlag As String

    On Error GoTo ErrHandler
    
    If pbooLockAdvice = True Then
        lstrInUseByFlag = LockingPhaseGen(True)
    Else
        lstrInUseByFlag = " "
    End If

    lstrSQL = "UPDATE " & gtblAdviceNotes & " SET LockingFlag = '" & lstrInUseByFlag & _
        "' WHERE (((OrderNum)=" & plngOrderNum & "));"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "ToggleAdviceInUseBy", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub ToggleRemarkInUseBy(plngRemarkNum As Long, pbooLockAccount As Boolean)
Dim lstrSQL As String
Dim lstrInUseByFlag As String
    
    On Error GoTo ErrHandler
    
    If pbooLockAccount = True Then
        ShowStatus 24
        lstrInUseByFlag = LockingPhaseGen(True)
    Else
        lstrInUseByFlag = " "
    End If

    lstrSQL = "UPDATE " & gtblRemarks & " SET LockingFlag = '" & lstrInUseByFlag & _
        "' WHERE (((RemarkNum)=" & plngRemarkNum & _
        "));"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "ToggleRemarkInUseBy", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub AddAdviceNote(pstrInUseByFlag As String, pstrOrderStatus As String)
Dim lstrSQL As String

    ShowStatus 25
    On Error GoTo ErrHandler

    With gstrCustomerAccount
        lstrSQL = "INSERT INTO " & gtblAdviceNotes & " ( CustNum, CallerSalutation, CallerSurname, CallerInitials," & _
            "AdviceAdd1, AdviceAdd2, AdviceAdd3, AdviceAdd4, AdviceAdd5, AdvicePostcode, " & _
            "TelephoneNum, DeliverySalutation, DeliverySurname, DeliveryInitials, DeliveryAdd1, DeliveryAdd2, DeliveryAdd3, DeliveryAdd4, " & _
            "DeliveryAdd5, DeliveryPostcode, MediaCode, DeliveryDate, CourierCode, " & _
            "OrderType, OrderCode, CardNumber, ExpiryDate, Donation, Payment, " & _
            "Underpayment, Reconcilliation, AdviceRemarkNum, ConsignRemarkNum, " & _
            "LockingFlag, PaymentType2, Payment2, TotalIncVat, VAT, Postage, " & _
            "OverSeasFlag, ProcessedBy, OrderStatus, OOSRefund, " & _
            "OrderStyle, CardName, BankRepPrintDate, DespatchDate, AuthorisationCode, CardType, CardIssueNumber, CardStartDate, Denom ) " & _
            "SELECT " & .lngCustNum & ", '" & OneSpace(JetSQLFixup(.strSalutation)) & _
            "', '" & OneSpace(JetSQLFixup(.strSurname)) & _
            "', '" & OneSpace(JetSQLFixup(.strInitials)) & "' , '" & OneSpace(JetSQLFixup(.strAdd1)) & _
            "' , '" & OneSpace(JetSQLFixup(.strAdd2)) & "', '" & OneSpace(JetSQLFixup(.strAdd3)) & "', '" & OneSpace(JetSQLFixup(.strAdd4)) & _
            "', '" & OneSpace(JetSQLFixup(.strAdd5)) & "', '" & OneSpace(.strPostcode) & "', '" & OneSpace(.strTelephoneNum) & _
            "', '" & OneSpace(JetSQLFixup(.strDeliverySalutation)) & _
            "', '" & OneSpace(JetSQLFixup(.strDeliverySurname)) & _
            "', '" & OneSpace(JetSQLFixup(.strDeliveryInitials)) & "', '" & OneSpace(JetSQLFixup(.strDeliveryAdd1)) & "', '" & OneSpace(JetSQLFixup(.strDeliveryAdd2)) & _
            "', '" & OneSpace(JetSQLFixup(.strDeliveryAdd3)) & "', '" & OneSpace(JetSQLFixup(.strDeliveryAdd4)) & _
            "', '" & OneSpace(JetSQLFixup(.strDeliveryAdd5)) & "', '" & OneSpace(.strDeliveryPostcode)
    End With
    
    With gstrAdviceNoteOrder
        lstrSQL = lstrSQL & _
            "', '" & OneSpace(.strMediaCode) & "', '" & CDate(.datDeliveryDate) & "', '" & _
            OneSpace(.strCourierCode) & "', '" & OneSpace(.strPaymentType1) & "', '" & _
            OneSpace(.strOrderCode) & "', '" & OneSpace(.strCardNumber) & "', '" & _
            CDate(.datExpiryDate) & "', '" & SystemPrice(.strDonation) & "', '" & _
            SystemPrice(.strPayment) & "', '" & SystemPrice(.strUnderpayment) & "', '" & _
            SystemPrice(.strReconcilliation) & "',  " & CLng(.lngAdviceRemarkNum) & ", " & _
            CLng(.lngConsignRemarkNum) & ", '" & pstrInUseByFlag & _
            "', '" & OneSpace(.strPaymentType2) & "', '" & SystemPrice(.strPayment2) & _
            "', '" & SystemPrice(.strTotalIncVat) & "', '" & SystemPrice(.strVAT) & _
            "', '" & SystemPrice(.strPostage) & "', '" & OneSpace(.strOverSeasFlag) & _
            "', '" & OneSpace(Trim$(gstrGenSysInfo.strUserName)) & _
            "', '" & pstrOrderStatus & "', '" & SystemPrice(.strOOSRefund) & _
            "', '" & OneSpace(.strOrderStyle) & "', '" & OneSpace(JetSQLFixup(.strCardName)) & _
            "', '" & CDate(.datBankRepPrintDate) & "', '" & CDate(.datDespatchDate) & _
            "', '" & OneSpace(.strAuthorisationCode) & _
            "', '" & OneSpace(.strCardType) & _
            "', " & CLng(.lngIssueNumber) & ", '" & CDate(.datCardStartDate) & "', '" & gstrReferenceInfo.strDenomination & "';"

    End With
        
    gdatCentralDatabase.Execute lstrSQL
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddAdviceNote", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub AppendOrderLinesToMaster(plngCustomerNum As Long, plngAdviceOrderNum As Long)
Dim lstrSQL As String

    ShowStatus 26
    On Error GoTo ErrHandler

    lstrSQL = "INSERT INTO " & gtblMasterOrderLines & " "
    
    lstrSQL = lstrSQL & "( CustNum, CatNum, ItemDescription, " & _
        "BinLocation, Qty, Price, Weight, TaxCode, TotalPrice, " & _
        "TotalWeight, OrderNum, Vat, Class, SalesCode, DespQty, OrderLineNum, ParcelNumber, Denom ) " & _
        "SELECT " & gtblOrderLines & ".CustNum, " & gtblOrderLines & ".CatNum, " & gtblOrderLines & ".ItemDescription, " & _
        "" & gtblOrderLines & ".BinLocation, " & gtblOrderLines & ".Qty, " & gtblOrderLines & ".Price, " & _
        "" & gtblOrderLines & ".Weight, " & gtblOrderLines & ".TaxCode, " & gtblOrderLines & ".TotalPrice, " & _
        "" & gtblOrderLines & ".TotalWeight, " & plngAdviceOrderNum & _
        " , " & gtblOrderLines & ".Vat, " & gtblOrderLines & ".Class, " & gtblOrderLines & ".SalesCode, " & _
        "" & gtblOrderLines & ".Qty, " & gtblOrderLines & ".OrderLineNum, " & gtblOrderLines & ".ParcelNumber, '" & _
        gstrReferenceInfo.strDenomination & "' " & _
        " FROM " & gtblOrderLines & " Where CustNum=" & plngCustomerNum & ";"

    gdatLocalDatabase.Execute lstrSQL

    lstrSQL = "Delete * from " & gtblOrderLines & ";"
    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AppendOrderLinesToMaster", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub
Sub AppendNewOrderLinesToMaster(plngCustomerNum As Long, plngAdviceOrderNum As Long)
Dim lstrSQL As String

    ShowStatus 26
    On Error GoTo ErrHandler
    
    lstrSQL = lstrSQL & "( CustNum, CatNum, ItemDescription, " & _
        "BinLocation, Qty, Price, Weight, TaxCode, TotalPrice, TotalWeight, " & _
        "OrderNum, Vat, Class, SalesCode, DespQty, ParcelNumber, OrderLineNum, Denom ) " & _
        "SELECT " & gtblOrderLines & ".CustNum, " & gtblOrderLines & ".CatNum, " & gtblOrderLines & ".ItemDescription, " & _
        "" & gtblOrderLines & ".BinLocation, " & gtblOrderLines & ".Qty, " & gtblOrderLines & ".Price, " & _
        "" & gtblOrderLines & ".Weight, " & gtblOrderLines & ".TaxCode, " & gtblOrderLines & ".TotalPrice, " & _
        "" & gtblOrderLines & ".TotalWeight, " & plngAdviceOrderNum & _
        " , " & gtblOrderLines & ".Vat, " & gtblOrderLines & ".Class, " & _
        "" & gtblOrderLines & ".SalesCode, " & gtblOrderLines & ".Qty, " & gtblOrderLines & ".ParcelNumber, " & gtblOrderLines & ".OrderLineNum, '" & gstrReferenceInfo.strDenomination & "' " & _
        " FROM " & gtblOrderLines & " Where CustNum=" & plngCustomerNum
    lstrSQL = lstrSQL & ";"

    gdatLocalDatabase.Execute lstrSQL
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AppendNewOrderLinesToMaster", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
    
End Sub

Sub UpdateMasterStockRecords(plngCustomerNum As Long, plngAdviceOrderNum As Long, pdateUpdateTime As Date)
Dim lstrSQL As String
Dim lvarErrorStage As Variant

    ShowStatus 31
    On Error GoTo ErrHandler
    
    lvarErrorStage = 180
    lstrSQL = "UPDATE " & gtblMasterOrderLines & " INNER JOIN " & gtblMasterProducts & " ON " & _
        "" & gtblMasterOrderLines & ".CatNum = " & gtblMasterProducts & ".CatNum SET " & _
        "" & gtblMasterProducts & ".NumInStock = " & _
        "[" & gtblMasterProducts & "].[NumInStock]-[" & gtblMasterOrderLines & "].[qty], " & _
        "" & gtblMasterProducts & ".UserUpdated = '" & pdateUpdateTime & "' " & _
        "WHERE (((" & gtblMasterOrderLines & ".CustNum)=" & plngCustomerNum & _
        ") AND ((" & gtblMasterOrderLines & ".OrderNum)=" & plngAdviceOrderNum & "));"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateMasterStockRecords", "Central", True, lvarErrorStage)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub

Function NewStockAvailable(Optional pbooNoStatus As Variant) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 27
    End If
    On Error GoTo ErrHandler
            
    lstrSQL = "SELECT " & gtblSystem & ".Item, " & gtblMachine & ".Item, " & gtblSystem & ".Value as SVal, " & gtblMachine & ".Value as MVal " & _
        "From " & gtblSystem & ", " & gtblMachine & " WHERE (((" & gtblSystem & ".Item)='StockUploaded') " & _
        "AND ((" & gtblMachine & ".Item)='StockDownloaded'));"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            If CDate(.Fields("SVal")) > CDate(.Fields("MVal")) Then
                NewStockAvailable = True
            Else
                NewStockAvailable = False
            End If
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "NewStockAvailable", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
 
End Function
Sub UpdateMachineStockFlag(Optional pbooNoStatus As Variant)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 48
    End If
    
    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblSystem & ", " & gtblMachine & " SET " & gtblMachine & ".[Value] = [" & gtblSystem & "].[Value] " & _
        "WHERE (((" & gtblSystem & ".Item)='StockUploaded') AND ((" & gtblMachine & ".Item)='StockDownloaded'));"
    
    gdatLocalDatabase.Execute lstrSQL
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateMachineStockFlag", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select


End Sub
Sub UpdateProductQtyFromMaster()
Dim lstrSQL As String
Dim lvarErrorStage As Variant

    ShowStatus 49
    On Error GoTo ErrHandler
    lvarErrorStage = 160
    lstrSQL = "UPDATE " & gtblMasterProducts & " INNER JOIN " & gtblProducts & " ON " & _
        "" & gtblMasterProducts & ".CatNum = " & gtblProducts & ".CatNum SET " & gtblProducts & ".NumInStock " & _
        "= [" & gtblMasterProducts & "].[NumInStock] WHERE (((" & gtblMasterProducts & ".UserUpdated) Is Not Null));"

    gdatLocalDatabase.Execute lstrSQL
    lstrSQL = "UPDATE " & gtblProducts & " SET " & gtblProducts & ".Selected = False, " & _
        "" & gtblProducts & ".Qty = 0, " & gtblProducts & ".OrderLineNum = 0 WHERE " & gtblProducts & ".Qty <>0 or " & gtblProducts & ".Selected <> False;"

    lvarErrorStage = 170
    gdatLocalDatabase.Execute lstrSQL
    
Exit Sub
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "UpdateProductQtyFromMaster", "local", True, lvarErrorStage)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub AddNewRemark(pstrLockingFlag As String)
Dim lstrSQL As String

    ShowStatus 28
    On Error GoTo ErrHandler

    lstrSQL = "INSERT INTO " & gtblRemarks & " ( LockingFlag ) select ('" & pstrLockingFlag & "');"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddNewRemark", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub AddNewCustomerNote(plngCustomerNum, pstrNotes As String)
Dim lstrSQL As String
 
    ShowStatus 50
    On Error GoTo ErrHandler
    
    If Trim$(pstrNotes) = "" Then
        Exit Sub
    End If
        
    lstrSQL = "INSERT INTO " & gtblCustNotes & " ( CustNum, Notes ) select " & _
        plngCustomerNum & " as CN, '" & JetSQLFixup(Trim$(pstrNotes)) & "' as N;"

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddNewCustomerNote", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub GetRemarkNum(pstrLockingFlag As String, pstrNoteVar As Remarks)
Dim ltabRemark As Recordset

    ShowStatus 52
    On Error GoTo ErrHandler
    
    Set ltabRemark = gdatCentralDatabase.OpenRecordset(gtblRemarks)
    ltabRemark.Index = "LockingFlag"
    
    With ltabRemark
        .Seek "=", pstrLockingFlag
        
        pstrNoteVar.lngRemarkNumber = .Fields("RemarkNum")
        pstrNoteVar.strText = .Fields("Remark") & ""
        pstrNoteVar.strType = .Fields("Type") & ""
        .Close

    End With

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetRemarkNum", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub

Sub GetRemark(plngRemarkNum As Long, pstrNoteVar As Remarks)
Dim lsnaLists As Recordset
Dim lstrSQL As String
  
    ShowStatus 51
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT * from " & gtblRemarks & " where RemarkNum = " & plngRemarkNum & ";"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        Do Until .EOF
            pstrNoteVar.lngRemarkNumber = .Fields("RemarkNum")
            pstrNoteVar.strText = .Fields("Remark") & ""
            pstrNoteVar.strType = .Fields("Type") & ""
            .MoveNext ' Won't happen
        Loop
    End With
    
    lsnaLists.Close

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetRemark", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function GetCustomerNote(plngCustomerNum As Long, ByRef pstrNotes As String) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim lintRecordInidicator As Integer

    lintRecordInidicator = 0
    ShowStatus 53
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT * from " & gtblCustNotes & " where CustNum = " & plngCustomerNum & ";"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        Do Until .EOF
            lintRecordInidicator = 1
            pstrNotes = .Fields("Notes")
            .MoveNext ' Won't happen
        Loop
    End With
    
    lsnaLists.Close

    If lintRecordInidicator = 0 Then
        GetCustomerNote = False
    Else
        GetCustomerNote = True
    End If
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetCustomerNote", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub UpdateRemark(plngRemarkNum As Long, pstrRemarkType As String, pstrRemarkText As String)
Dim lstrSQL As String
 
    ShowStatus 54
    On Error GoTo ErrHandler
    
    With gstrCustomerAccount
        
        lstrSQL = "UPDATE " & gtblRemarks & " SET " & _
            "Type = '" & pstrRemarkType & "', " & _
            "Remark = '" & JetSQLFixup(Trim$(pstrRemarkText)) & "' " & _
            "WHERE (((RemarkNum)=" & plngRemarkNum & "));"
    
    End With
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateRemark", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Sub
Sub UpdateRemarkAdviceID(plngAdviceOrderNum As Long, plngRemarkNum As Long, pstrRemarkType As String)
Dim lstrSQL As String

    ShowStatus 55
    On Error GoTo ErrHandler
        
    Select Case pstrRemarkType
    Case "Internal"
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET " & _
            "AdviceRemarkNum = " & plngRemarkNum & _
            " WHERE (((OrderNum)=" & plngAdviceOrderNum & "));"
    Case "Consignment"
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET " & _
            "ConsignRemarkNum = " & plngRemarkNum & _
            " WHERE (((OrderNum)=" & plngAdviceOrderNum & "));"
    End Select
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateRemarkAdviceID", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Sub
Sub UpdateCustomerNotes(plngCustomerNum As Long, pstrNotes As String)
Dim lstrSQL As String

    ShowStatus 56
    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblCustNotes & " SET " & gtblCustNotes & ".Notes = '" & JetSQLFixup(pstrNotes) & "' " & _
        "WHERE (((" & gtblCustNotes & ".CustNum)=" & plngCustomerNum & "));"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateCustomerNotes", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateOldAcIndicator(plngOldAcNumber As Long)
Dim lstrSQL As String
    
    On Error GoTo ErrHandler

    lstrSQL = "UPDATE OldCustAccounts SET OldCustAccounts.DBIndicator = 'X' " & _
        "WHERE (((OldCustAccounts.CustNum)=" & plngOldAcNumber & "));"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOldAcIndicator", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
        
End Sub
Sub ClearCustomerAcount()

    ShowStatus 57
    
    With gstrCustomerAccount
        .lngCustNum = 0
        .strSurname = ""
        .strInitials = ""
        .strSalutation = ""
        .strAdd1 = ""
        .strAdd2 = ""
        .strAdd3 = ""
        .strAdd4 = ""
        .strAdd5 = ""
        .strPostcode = ""
        .strTelephoneNum = ""
        .strDeliverySalutation = ""
        .strDeliveryInitials = ""
        .strDeliverySurname = ""
        .strDeliveryAdd1 = ""
        .strDeliveryAdd2 = ""
        .strDeliveryAdd3 = ""
        .strDeliveryAdd4 = ""
        .strDeliveryAdd5 = ""
        .strDeliveryPostcode = ""
        .strAcctInUseByFlag = ""
        .strAccountType = ""
        .strReceiveMailings = ""
        .strAccountStatus = ""
        .strEmail = ""
        .strEveTelephoneNum = ""
    End With
    
End Sub
Sub ClearAdviceNote()

    ShowStatus 58
    
    With gstrAdviceNoteOrder
        .lngCustNum = 0
        .lngOrderNum = 0
        .strMediaCode = ""
        .datDeliveryDate = 0
        .strCourierCode = ""
        .strPaymentType1 = ""
        .strOrderCode = ""
        .strCardNumber = ""
        .datExpiryDate = 0
        .strDonation = ""
        .strPayment = ""
        .strUnderpayment = ""
        .strReconcilliation = ""
        .strOOSRefund = ""
        .lngAdviceRemarkNum = 0
        .lngConsignRemarkNum = 0
        .strAcctInUseByFlag = ""
        .lngIssueNumber = 0
        .strCardType = ""
        .strPayment2 = 0
        .strPaymentType2 = 0
        .strPostage = 0
        .strTotalIncVat = 0
        .strVAT = 0
        .strOverSeasFlag = ""
        .strSalutation = ""
        .strInitials = ""
        .strSurname = ""
        .strAdd1 = ""
        .strAdd2 = ""
        .strAdd3 = ""
        .strAdd4 = ""
        .strAdd5 = ""
        .strPostcode = ""
        .strTelephoneNum = ""
        .strDeliverySalutation = ""
        .strDeliveryInitials = ""
        .strDeliverySurname = ""
        .strDeliveryAdd1 = ""
        .strDeliveryAdd2 = ""
        .strDeliveryAdd3 = ""
        .strDeliveryAdd4 = ""
        .strDeliveryAdd5 = ""
        .strDeliveryPostcode = ""
        .strOrderStyle = ""
        .datDespatchDate = 0
        .datBankRepPrintDate = 0
        .strCardName = ""
        .strAuthorisationCode = ""
        .datCreationDate = 0
        .intNumOfParcels = 0
        .lngGrossWeight = 0
        .lngStockBatchNum = 0
        .datCardStartDate = 0
        .strDenom = ""
    End With
    
End Sub
Sub OrderTotal(plngCustomerNum As Long)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    ShowStatus 59
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT " & gtblOrderLines & ".CustNum, Sum(" & gtblOrderLines & ".TotalPrice) AS " & _
        "SumOfTotalPrice, sum(" & gtblOrderLines & ".Vat) as SumOfVat " & _
        "From " & gtblOrderLines & " GROUP BY " & gtblOrderLines & ".CustNum " & _
        "HAVING (((" & gtblOrderLines & ".CustNum)=" & plngCustomerNum & "));"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            gstrOrderTotal = .Fields("SumOfTotalPrice")
            gstrVatTotal = .Fields("SumOfVat")
        End If
    End With
    
    lsnaLists.Close

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "OrderTotal", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub ClearGen()
    
    ShowStatus 60
    
    With gstrInternalNote
        .lngRemarkNumber = 0
        .strType = ""
        .strText = ""
    End With
    
    With gstrConsignmentNote
        .lngRemarkNumber = 0
        .strType = ""
        .strText = ""
    End With
    
    gstrOrderTotal = ""
    gstrVatTotal = ""

End Sub
Sub xAddRefund(pstrRefundPrice As String, plngCustomerNum As Long, plngOrderNum As Long)
'Function not in use
Dim lstrSQL As String
    
    ShowStatus 29
    On Error GoTo ErrHandler

    lstrSQL = "INSERT INTO Refunds ( CustNum, OrderNum, Amount ) select " & _
        plngCustomerNum & ", " & plngOrderNum & ", '" & _
        pstrRefundPrice & "';"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddRefund", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateSalesCode()
Dim lstrSQL As String

    ShowStatus 61
    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblOrderLines & ", " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & _
        "" & gtblListDetails & ".ListNum = " & gtblLists & ".ListNum SET " & gtblOrderLines & ".SalesCode " & _
        "= Val([UserDef1]) WHERE (((" & gtblLists & ".ListName)='Product Classes') " & _
        "AND ((" & gtblOrderLines & ".Class)=Val([" & gtblListDetails & "].[ListCode])));"

    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateSalesCode", "Local", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub AppendOldAcToNew(plngNewAcNum As Long, plngOldAcNum As Long)
Dim lstrSQL As String

    ShowStatus 30
    On Error GoTo ErrHandler

    lstrSQL = "UPDATE CustAccounts, OldCustAccounts SET " & _
        "CustAccounts.Salutation = [OldCustAccounts].[Salutation], " & _
        "CustAccounts.Surname = [OldCustAccounts].[Surname], " & _
        "CustAccounts.Initials = [OldCustAccounts].[Initials], " & _
        "CustAccounts.Add1 = [OldCustAccounts].[Add1], " & _
        "CustAccounts.Add2 = [OldCustAccounts].[Add2], " & _
        "CustAccounts.Add3 = [OldCustAccounts].[Add3], " & _
        "CustAccounts.Add4 = [OldCustAccounts].[Add4], " & _
        "CustAccounts.Add5 = [OldCustAccounts].[Add5], " & _
        "CustAccounts.Postcode = [OldCustAccounts].[Postcode], " & _
        "CustAccounts.TelephoneNum = [OldCustAccounts].[TelephoneNum], " & _
        "CustAccounts.DBIndicator = 'N', " & _
        "CustAccounts.CreationDate =[OldCustAccounts].[CreationDate] " & _
        "WHERE (((CustAccounts.CustNum)=" & plngNewAcNum & _
        ") AND ((OldCustAccounts.CustNum)=" & plngOldAcNum & "));"
    
    gdatCentralDatabase.Execute lstrSQL
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AppendOldACToNew", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub

Function GetAdviceNote(plngCustNumber As Long, plngOrderNum As Long) As Boolean
Dim ltabAdviceNotes As Recordset

    ShowStatus 62
    On Error GoTo ErrHandler

    Set ltabAdviceNotes = gdatCentralDatabase.OpenRecordset(gtblAdviceNotes)
    ltabAdviceNotes.Index = "Advice"
    
    With ltabAdviceNotes
        .Seek "=", plngCustNumber, plngOrderNum
        
        If .NoMatch Then
            GetAdviceNote = False
            Exit Function
        End If
        GetAdviceNote = True
        
        gstrAdviceNoteOrder.lngCustNum = Trim$(!CustNum)
        gstrAdviceNoteOrder.lngOrderNum = Trim$(!OrderNum)
        gstrAdviceNoteOrder.strSalutation = Trim$(!CallerSalutation & "")
        gstrAdviceNoteOrder.strSurname = Trim$(!CallerSurname & "")
        gstrAdviceNoteOrder.strInitials = Trim$(!CallerInitials & "")
        gstrAdviceNoteOrder.strAdd1 = Trim$(!AdviceAdd1 & "")
        gstrAdviceNoteOrder.strAdd2 = Trim$(!AdviceAdd2 & "")
        gstrAdviceNoteOrder.strAdd3 = Trim$(!AdviceAdd3 & "")
        gstrAdviceNoteOrder.strAdd4 = Trim$(!AdviceAdd4 & "")
        gstrAdviceNoteOrder.strAdd5 = Trim$(!AdviceAdd5 & "")
        gstrAdviceNoteOrder.strPostcode = Trim$(!AdvicePostcode & "")
        gstrAdviceNoteOrder.strTelephoneNum = Trim$(!TelephoneNum & "")
        gstrAdviceNoteOrder.strDeliverySalutation = Trim$(!DeliverySalutation & "")
        gstrAdviceNoteOrder.strDeliverySurname = Trim$(!DeliverySurname & "")
        gstrAdviceNoteOrder.strDeliveryInitials = Trim$(!DeliveryInitials & "")
        
        gstrAdviceNoteOrder.strDeliveryAdd1 = Trim$(!DeliveryAdd1 & "")
        gstrAdviceNoteOrder.strDeliveryAdd2 = Trim$(!DeliveryAdd2 & "")
        gstrAdviceNoteOrder.strDeliveryAdd3 = Trim$(!DeliveryAdd3 & "")
        gstrAdviceNoteOrder.strDeliveryAdd4 = Trim$(!DeliveryAdd4 & "")
        gstrAdviceNoteOrder.strDeliveryAdd5 = Trim$(!DeliveryAdd5 & "")
        gstrAdviceNoteOrder.strDeliveryPostcode = Trim$(!DeliveryPostcode & "")
        gstrAdviceNoteOrder.strMediaCode = Trim$(!MediaCode & "")
        
        If Not IsNull(!DeliveryDate) Then
            gstrAdviceNoteOrder.datDeliveryDate = Trim(!DeliveryDate & "")
        Else
            gstrAdviceNoteOrder.datDeliveryDate = 0
        End If
        gstrAdviceNoteOrder.strCourierCode = Trim$(!CourierCode & "")
        gstrAdviceNoteOrder.strPaymentType1 = Trim$(!OrderType & "")
        gstrAdviceNoteOrder.strOrderCode = Trim$(!OrderCode & "")
        gstrAdviceNoteOrder.strCardNumber = Trim$(!CardNumber & "")
        
        gstrAdviceNoteOrder.lngIssueNumber = CLng(Val(!CardIssueNumber & ""))
        gstrAdviceNoteOrder.strCardType = Trim$(!CardType & "")
        
        If Not IsNull(!ExpiryDate) Then
            gstrAdviceNoteOrder.datExpiryDate = Trim$(!ExpiryDate)
        Else
            gstrAdviceNoteOrder.datExpiryDate = 0
        End If
        
        If Not IsNull(!CardStartDate) Then
            gstrAdviceNoteOrder.datCardStartDate = Trim$(!CardStartDate)
        Else
            gstrAdviceNoteOrder.datCardStartDate = 0
        End If
        
        gstrAdviceNoteOrder.strDonation = Trim$(!Donation)
        gstrAdviceNoteOrder.strPayment = Trim$(!Payment)
        gstrAdviceNoteOrder.strUnderpayment = Trim$(!UnderPayment)
        gstrAdviceNoteOrder.strReconcilliation = Trim$(!Reconcilliation)
        If Not IsNull(!OOSRefund) Then
            gstrAdviceNoteOrder.strOOSRefund = Trim$(!OOSRefund)
        End If
        gstrAdviceNoteOrder.lngAdviceRemarkNum = Trim$(!AdviceRemarkNum)
        gstrAdviceNoteOrder.lngConsignRemarkNum = Trim$(!ConsignRemarkNum)
        If Not IsNull(!LockingFlag) Then
            gstrAdviceNoteOrder.strAcctInUseByFlag = Trim$(!LockingFlag)
        End If
        
        gstrAdviceNoteOrder.strPayment2 = Trim$(!Payment2)
        gstrAdviceNoteOrder.strPaymentType2 = Trim$(!PaymentType2)
        gstrAdviceNoteOrder.strPostage = Trim$(!Postage)
        gstrAdviceNoteOrder.strTotalIncVat = Trim$(!TotalIncVat)
        gstrAdviceNoteOrder.strVAT = Trim$(!Vat)
        gstrAdviceNoteOrder.strOverSeasFlag = Trim$(!OverSeasFlag)
        gstrAdviceNoteOrder.datCreationDate = Trim(!CreationDate)
        If Not IsNull(!BankRepPrintDate) Then
            gstrAdviceNoteOrder.datBankRepPrintDate = Trim(!BankRepPrintDate)
        Else
            gstrAdviceNoteOrder.datBankRepPrintDate = 0
        End If
        If Not IsNull(!AuthorisationCode) Then
            gstrAdviceNoteOrder.strAuthorisationCode = Trim$(!AuthorisationCode)
        Else
            gstrAdviceNoteOrder.strAuthorisationCode = " "
        End If
        If Not IsNull(!CardName) Then
            gstrAdviceNoteOrder.strCardName = Trim$(!CardName)
        Else
            gstrAdviceNoteOrder.strCardName = " "
        End If
        If Not IsNull(!DespatchDate) Then
            gstrAdviceNoteOrder.datDespatchDate = Trim$(!DespatchDate)
        Else
            gstrAdviceNoteOrder.datDespatchDate = 0
        End If
        If Not IsNull(!OrderStyle) Then
            gstrAdviceNoteOrder.strOrderStyle = Trim$(!OrderStyle)
        Else
            gstrAdviceNoteOrder.strOrderStyle = 0
        End If
    
        gstrAdviceNoteOrder.intNumOfParcels = Val(!NumOfParcels & "")
        gstrAdviceNoteOrder.strDenom = Trim$(!Denom) & ""
        .Close

    End With

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetAdviceNote", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function

Sub UpdateAdviceNote()
Dim lstrSQL As String

    ShowStatus 63
    On Error GoTo ErrHandler
    
    With gstrAdviceNoteOrder

    'This should have not cause an error as Updating old currencies is not allowed,
    'it would have caused an error as you can't update an access field with the non
    'regional currency
    
    lstrSQL = "UPDATE " & gtblAdviceNotes & " SET " & _
        "DeliveryDate = '" & CDate(.datDeliveryDate) & "', " & _
        "CourierCode = '" & OneSpace(.strCourierCode) & "', " & _
        "OrderType = '" & OneSpace(.strPaymentType1) & "', " & _
        "PaymentType2 = '" & OneSpace(.strPaymentType2) & "', " & _
        "OrderCode = '" & OneSpace(.strOrderCode) & "', " & _
        "CardNumber = '" & OneSpace(.strCardNumber) & "', " & _
        "ExpiryDate = '" & CDate(.datExpiryDate) & "', " & _
        "Donation = '" & AdvicePrice(.strDonation) & "', " & _
        "Payment = '" & AdvicePrice(.strPayment) & "', " & _
        "Payment2 = '" & AdvicePrice(.strPayment2) & "', " & _
        "Underpayment = '" & AdvicePrice(.strUnderpayment) & "', " & _
        "Reconcilliation = '" & AdvicePrice(.strReconcilliation) & "', " & _
        "Postage = '" & AdvicePrice(.strPostage) & "', " & _
        "Vat = '" & AdvicePrice(.strVAT) & "', " & _
        "TotalIncVat = '" & AdvicePrice(.strTotalIncVat) & "', " & _
        "OOSRefund = '" & AdvicePrice(.strOOSRefund) & "', " & _
        "AdviceRemarkNum = " & CLng(.lngAdviceRemarkNum) & ", " & _
        "ConsignRemarkNum = " & CLng(.lngConsignRemarkNum) & ", " & _
        "OverSeasFlag = '" & OneSpace(.strOverSeasFlag) & "', " & _
        "ProcessedBy = '" & OneSpace(Trim$(gstrGenSysInfo.strUserName)) & "', "
lstrSQL = lstrSQL & _
        "OrderStyle = '" & OneSpace(.strOrderStyle) & "', " & _
        "CardName = '" & OneSpace(JetSQLFixup(.strCardName)) & "', " & _
        "BankRepPrintDate = '" & CDate(.datBankRepPrintDate) & "', " & _
        "DespatchDate = '" & CDate(.datDespatchDate) & "', " & _
        "AuthorisationCode = '" & OneSpace(.strAuthorisationCode) & "', " & _
        "CardType = '" & OneSpace(.strCardType) & "', " & _
        "CardStartDate = '" & CDate(.datCardStartDate) & "', " & _
        "CardIssueNumber = " & CLng(.lngIssueNumber) & " " & _
        "WHERE (((CustNum)=" & gstrAdviceNoteOrder.lngCustNum & ") AND " & _
        "((OrderNum)=" & gstrAdviceNoteOrder.lngOrderNum & "));"
    End With
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateAdviceNote", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select


End Sub
Sub UpdateOrderStatus(pstrStatus As String, pdatEndDate As Date, pstrParam As String, _
Optional plngOrderNum As Variant, Optional plngEndOrderNum As Variant, Optional pstrFindStatus As Variant)
Dim lstrSQL As String

    ShowStatus 64
    On Error GoTo ErrHandler
        
    lstrSQL = "UPDATE " & gtblAdviceNotes & " SET " & gtblAdviceNotes & ".OrderStatus = '" & pstrStatus & "' "
      
    Select Case pstrParam
    Case "B"
        lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".DespatchDate = '" & Format(Date, "DD/MMM/YYYY") & "' "
        lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderStatus ='B' "
        lstrSQL = lstrSQL & "AND (" & gtblAdviceNotes & ".CreationDate < #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# "
        lstrSQL = lstrSQL & "or " & gtblAdviceNotes & ".CreationDate = #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# );"
    Case "C"
        lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".DespatchDate = '" & Format(Date, "DD/MMM/YYYY") & "' "
        lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderStatus ='C' "
        lstrSQL = lstrSQL & "AND (" & gtblAdviceNotes & ".CreationDate < #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# "
        lstrSQL = lstrSQL & "or " & gtblAdviceNotes & ".CreationDate = #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# );"
    Case "Y"
        lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".StockBatchNum = 'M" & Trim$(glngStockBatchNumber) & "' "
        lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderStatus ='B' "
        lstrSQL = lstrSQL & "AND " & gtblAdviceNotes & ".CreationDate < #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# "
        lstrSQL = lstrSQL & "or AdviceNotes.CreationDate = #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# ;"
    Case "Z"
        lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".StockBatchNum = 'M" & glngStockBatchNumber & "' "
        lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderStatus ='C' "
        lstrSQL = lstrSQL & "AND " & gtblAdviceNotes & ".CreationDate < #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# "
        lstrSQL = lstrSQL & "or " & gtblAdviceNotes & ".CreationDate = #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# ;"
    Case "A"
        lstrSQL = lstrSQL & "WHERE (((" & gtblAdviceNotes & ".OrderStatus)='A') AND " & _
        "((" & gtblAdviceNotes & ".OrderNum)<=" & glngLastOrderPrintedInThisRun & "));"
    Case "R"
        lstrSQL = lstrSQL & "Where (((" & gtblAdviceNotes & ".OrderStatus)='A') AND " & _
            "((" & gtblAdviceNotes & ".OrderNum)>=" & plngOrderNum & _
            " And (" & gtblAdviceNotes & ".OrderNum)<=" & plngEndOrderNum & ")) "
    Case "S" ' Update  for specfic OrderNum
        Select Case pstrStatus
        Case "B", "C"
            lstrSQL = lstrSQL & ", " & gtblAdviceNotes & ".DespatchDate = '" & Format(Date, "DD/MMM/YYYY") & "' "
            lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderNum= " & plngOrderNum & ";"
        Case Else
            lstrSQL = lstrSQL & "WHERE " & gtblAdviceNotes & ".OrderNum= " & plngOrderNum & ";"
        End Select
    Case "F"
        lstrSQL = lstrSQL & "WHERE (((" & gtblAdviceNotes & ".OrderStatus)='" & pstrFindStatus & "') AND " & _
        "((" & gtblAdviceNotes & ".OrderNum)<=" & glngLastOrderPrintedInThisRun & "));"
    
    End Select

    If pstrStatus = "B" And gstrStatic.strVerLogBStatus <> "" Then
        ErrorLogging Now() & " BSTAT-" & "V" & App.Major & "." & App.Minor & "." & App.Revision & " " & gstrGenSysInfo.strUserName & " " & lstrSQL
    End If
    
    gdatCentralDatabase.Execute lstrSQL
    
    ShowStatus 0
    DoEvents

Exit Sub
ErrHandler:

    Select Case GlobalErrorHandler(Err.Number, "UpdateOrderStatus", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        On Error GoTo 0
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function CheckCanConfirm(plngOrderNum As Long, Optional pvarParam As Variant, _
    Optional ByRef pvarOrderStatus As Variant) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pvarParam) Then
        pvarParam = ""
    End If
    
    ShowStatus 65
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT OrderStatus " & _
        "FROM " & gtblAdviceNotes & " WHERE OrderNum = " & plngOrderNum & ";"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
                Select Case .Fields("OrderStatus")
                Case "D", "C", "B", "E"
                    MsgBox "This Order has been flagged as packed (" & .Fields("OrderStatus") & ")", , gconstrTitlPrefix & "Check Order Status"
                    CheckCanConfirm = False
                Case "A", "P"
                    CheckCanConfirm = True
                    pvarOrderStatus = .Fields("OrderStatus")
                Case "X"
                    If pvarParam = "CANCONX" Then
                        CheckCanConfirm = True
                    Else
                        MsgBox "This Order has been Cancelled!", , gconstrTitlPrefix & "Check Order Status"
                        CheckCanConfirm = False
                    End If
                Case "S"
                    MsgBox "This Order has been Cancelled!", , gconstrTitlPrefix & "Check Order Status"
                    CheckCanConfirm = False
                Case "H"
                    MsgBox "This Order is currently on hold, awaiting financial confirmation!", , gconstrTitlPrefix & "Check Order Status"
                    CheckCanConfirm = False
                End Select
            .MoveNext
        Loop

    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckCanConfirm", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub OrderMasterDespTotal(plngOrderNum As Long)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    ShowStatus 66

    On Error GoTo ErrHandler

    lstrSQL = "SELECT Sum(IIf(Trim$([TaxCode])='Z',([DespQty]*[Price]), " & _
        "([DespQty]*[Price])+([DespQty]*[Price])*" & gstrVATRate / 100 & ")) AS TotalPricex, " & _
        "Sum([DespQty]*[Weight]) AS TW, Sum(IIf(Trim$([TaxCode])='Z',0," & _
        "([DespQty]*[Price])*" & gstrVATRate / 100 & ")) AS VAT From " & gtblMasterOrderLines & " " & _
        "WHERE (((" & gtblMasterOrderLines & ".OrderNum)=" & plngOrderNum & "));"
    
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            gstrOrderTotal = .Fields("TotalPricex")
            gstrVatTotal = .Fields("VAT")
        End If
    End With
    
    lsnaLists.Close

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "OrderMasterDespTotal", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select


End Sub
Sub xFillNewDespTotalArray(plngOrderNum As Long)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT IIf(Trim$([TaxCode])='Z',([DespQty]*[Price])," & _
        "([DespQty]*[Price])+([DespQty]*[Price])*" & gstrVATRate / 100 & _
        ") AS TotalPricex, " & _
        "[DespQty]*[Weight] AS TW, IIf(Trim$([TaxCode])='Z',0,([DespQty]*[Price])*" & _
         gstrVATRate / 100 & ") AS VAT From " & gtblMasterOrderLines & _
         " WHERE ((([OrderNum])=" & plngOrderNum & "));"

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        Dim llngRecCount
        Dim pstrIndexArray

        Do Until .EOF
            llngRecCount = llngRecCount + 1

            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillNewDespTotalArray", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
    

End Sub
Sub AddCashBookEntry(pstrChequeNum As String, pcurValue As Currency, _
pstrName As String, plngCustomerNum As Long, plngOrderNum As Long, pstrReason As String)
Dim lstrSQL As String
Dim pdatPrintedDate As Date
Dim pstrPrintedBy As String
Dim lstrInUseByFlag As String
Dim lstrLockingFlag As String
Dim lcurOriginalReconcil As Currency

    'Just in case there needed in the future
    pdatPrintedDate = "00:00:00"
    pstrPrintedBy = ""
    
    ShowStatus 67
    On Error GoTo ErrHandler

    'copy orginal order
    With gstrAdviceNoteOrder
        .datBankRepPrintDate = CDate("0")
        .datCreationDate = CDate(Now())
        .datDeliveryDate = CDate("0")
        .datDespatchDate = CDate("0")
        .intNumOfParcels = 0
        .lngAdviceRemarkNum = 0
        .lngConsignRemarkNum = 0
        .lngGrossWeight = 0
        .strAcctInUseByFlag = ""
        .strAuthorisationCode = ""
        .strDonation = "0"
        .strOOSRefund = "0"
        .strOrderStyle = "1" ' regular
        lcurOriginalReconcil = .strReconcilliation
        .strReconcilliation = "0"
        .strUnderpayment = "0"

        .strOrderCode = "O"  ' Post

        Select Case pstrReason
        Case "OVERPAY"
            .strPayment = CCur("-" & Val(CCur(pcurValue)))
            .strTotalIncVat = CCur("-" & Val(CCur(pcurValue)))
        Case "UNDERPAY"
            .strPayment = CCur(pcurValue)
            .strTotalIncVat = CCur(pcurValue)
        Case "STOCKOUT", "REFUND" 'Refund should never be reached!
            .strPayment = CCur("-" & Val(CCur(pcurValue)))
            .strTotalIncVat = CCur("-" & Val(CCur(pcurValue)))
        Case "OUTOFSTOCK"
            .strPayment = CCur("-" & Val(CCur(pcurValue)))
            .strTotalIncVat = CCur("-" & Val(CCur(pcurValue)))
        End Select
            
        lstrInUseByFlag = LockingPhaseGen(True)
        gstrCustomerAccount.lngCustNum = plngCustomerNum
        GetCustomerAccount plngCustomerNum, False
        
        AddAdviceNote lstrInUseByFlag, "R"
        GetAdviceOrderNum lstrInUseByFlag, .lngCustNum
        
        'Add Consignment note
        lstrLockingFlag = LockingPhaseGen(True)
        AddNewRemark lstrLockingFlag
        GetRemarkNum lstrLockingFlag, gstrConsignmentNote
        gstrAdviceNoteOrder.lngConsignRemarkNum = gstrConsignmentNote.lngRemarkNumber
        UpdateRemarkAdviceID gstrAdviceNoteOrder.lngOrderNum, gstrAdviceNoteOrder.lngConsignRemarkNum, "Consignment"
        ToggleRemarkInUseBy gstrConsignmentNote.lngRemarkNumber, False
        
        gstrConsignmentNote.strType = "Consignment"
        Select Case pstrReason
        Case "OVERPAY"
            gstrConsignmentNote.strText = "There was an over payment on your order."
        Case "UNDERPAY"
            gstrConsignmentNote.strText = "There was an under payment on your order."
        Case "STOCKOUT", "REFUND" 'Refund should never be reached!
            gstrConsignmentNote.strText = "All of the items you ordered were out of stock."
        Case "OUTOFSTOCK"
            If lcurOriginalReconcil = 0 Then
                gstrConsignmentNote.strText = "The items you ordered above were out of stock."
            Else
                gstrConsignmentNote.strText = "Above items out of stock and an over payment."
            End If
        End Select
        
        
        UpdateRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote.strType, gstrConsignmentNote.strText
                
        UpdateAdviceNote
        
        'Add to new refund fields
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET RefundOrignNum = " & plngOrderNum & ", " & _
            "RefundReason = '" & pstrReason & "', ChequeRequestDate = #" & Format(Date, "dd/mmm/yyyy") & "# " & _
            "WHERE CustNum=" & plngCustomerNum & " AND OrderNum=" & .lngOrderNum & ";"
        gdatCentralDatabase.Execute lstrSQL
        UpdateOrderStatus "R", 0, "S", .lngOrderNum
            
        'if out of stock order, append out of stock items.
        lstrSQL = "INSERT INTO OrderLinesMaster ( CustNum, OrderNum, OrderLineNum, " & _
            "CatNum, ItemDescription, BinLocation, Qty, DespQty, Price, Vat, " & _
            "Weight, TaxCode, TotalPrice, TotalWeight, Class, SalesCode, " & _
            "ParcelNumber, Denom ) SELECT CustNum, " & .lngOrderNum & _
            " AS Expr1, OrderLineNum, CatNum, ItemDescription, BinLocation, " & _
            "[Qty]-[DespQty] AS Ex2, [Qty]-[DespQty] AS Ex3, Price, Vat, " & _
            "0, TaxCode, TotalPrice, 0, Class, SalesCode, ParcelNumber, " & _
            "Denom FROM OrderLinesMaster WHERE (((CustNum)=" & plngCustomerNum & _
            ") AND ((DespQty)<>[Qty]) AND ((OrderNum)=" & plngOrderNum & "));"
        gdatCentralDatabase.Execute lstrSQL
        
        If Trim$(UCase$(.strOverSeasFlag)) = "N" Then
            UpdateOrderLinesTotalsDespQty .lngOrderNum, False
        Else
            UpdateOrderLinesTotalsDespQty .lngOrderNum, True
        End If
    End With
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddCashBookEntry", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub CalculateRefund(plngCustomerNum As Long, plngOrderNum As Long, pbooSomethingOutofStock As Boolean)
Dim lintOrderSubTotal As Integer
Dim lbooSomethingOutofStock As Boolean
Dim llngOrderNum As Long
Dim llngCustomerNum As Long
Dim lstrTotalNonCardPayments As String
Dim lstrTotalCardPayments As String
Dim lstrDebugMessage As String
Dim lstrName As String
Dim lbooDeductFromCreditCard As Boolean

    ShowStatus 68
    
    gstrOrderTotal = 0
    gstrVatTotal = 0
    lstrTotalCardPayments = 0
    lstrTotalNonCardPayments = 0
    
    llngOrderNum = plngOrderNum
    
    llngCustomerNum = plngCustomerNum
    
    lbooSomethingOutofStock = pbooSomethingOutofStock
    lbooDeductFromCreditCard = False
    
    On Error GoTo ErrHandler
    Busy True, Screen.ActiveForm
    
    GetAdviceNote llngCustomerNum, llngOrderNum
    With gstrAdviceNoteOrder
        
        lstrName = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))

        If lbooSomethingOutofStock = True Then
            If Trim$(UCase$(.strOverSeasFlag)) = "N" Then
                UpdateOrderLinesTotalsDespQty llngOrderNum, False
            Else
                UpdateOrderLinesTotalsDespQty llngOrderNum, True
            End If
        Else
            'check overpayment refund
            If CCur(.strReconcilliation) > 0 Then
                AddCashBookEntry "", CCur(.strReconcilliation), lstrName, llngCustomerNum, llngOrderNum, "OVERPAY"
                UpdateOrderStatus "C", 0, "S", llngOrderNum
                Busy False, Screen.ActiveForm
                MsgBox "Order status updated to (C)", , gconstrTitlPrefix & "Calculate Refund"
                Exit Sub
            Else
                'order confirmed no refunds, no items OOS
                
                If CCur(.strUnderpayment) > 0 Then
                    AddCashBookEntry "", CCur(.strUnderpayment), lstrName, llngCustomerNum, llngOrderNum, "UNDERPAY"
                End If

                UpdateOrderStatus "C", 0, "S", llngOrderNum
                Busy False, Screen.ActiveForm
                MsgBox "Order status updated to (C)", , gconstrTitlPrefix & "Calculate Refund"
                
                Exit Sub
            End If
        End If
    
        OrderMasterDespTotal llngOrderNum
        gstrVatTotal = SystemPrice(gstrVatTotal)
        gstrOrderTotal = SystemPrice(gstrOrderTotal)
        .strVAT = gstrVatTotal
            
        If SystemPrice(CCur(gstrOrderTotal) + CCur(.strDonation)) <> 0 Then
            gstrOrderTotal = SystemPrice(CCur(gstrOrderTotal) + CCur(.strPostage) + _
                CCur(.strDonation))
        Else
            gstrOrderTotal = 0
            .strTotalIncVat = 0
            .strReconcilliation = 0
            .strUnderpayment = 0
            
            UpdateAdviceNote

            If UCase$(GetListCodeDesc("Payment Method", .strPaymentType1)) = "CHEQUE" Then
                AddCashBookEntry "", CCur(.strPayment), lstrName, llngCustomerNum, llngOrderNum, "STOCKOUT"
            End If
            If UCase$(GetListCodeDesc("Payment Method", .strPaymentType2)) = "CHEQUE" Then
                AddCashBookEntry "", CCur(.strPayment2), lstrName, llngCustomerNum, llngOrderNum, "STOCKOUT"
            End If

            UpdateOrderStatus "R", 0, "S", llngOrderNum

            Busy False, Screen.ActiveForm

            MsgBox "Order status updated to (R), Cancelled / Refunded no items in stock!", , gconstrTitlPrefix & "Calculate Refund"

            Exit Sub
        End If
        
        .strTotalIncVat = gstrOrderTotal
        
        'find out what payments
        'if credit card payment - subtract from that
        If UCase$(GetListCodeDesc("Payment Method", .strPaymentType1)) = "CREDIT CARD" Then
            lstrTotalCardPayments = SystemPrice(CCur(.strPayment))
            lstrTotalNonCardPayments = SystemPrice(CCur(.strPayment2))
        ElseIf UCase$(GetListCodeDesc("Payment Method", .strPaymentType2)) = "CREDIT CARD" Then
            lstrTotalCardPayments = SystemPrice(CCur(.strPayment2))
            lstrTotalNonCardPayments = SystemPrice(CCur(.strPayment))
        Else
            lstrTotalNonCardPayments = SystemPrice(CCur(.strPayment) + CCur(.strPayment2))
        End If
                          
        If CCur(CCur(lstrTotalNonCardPayments) + CCur(lstrTotalCardPayments)) > CCur(gstrOrderTotal) Then
            If CCur(lstrTotalCardPayments) > 0 Then
                If CCur(lstrTotalNonCardPayments) = 0 Then
                    lstrTotalCardPayments = CCur(gstrOrderTotal)
                    If UCase$(GetListCodeDesc("Payment Method", .strPaymentType1)) = "CREDIT CARD" Then
                        .strPayment = SystemPrice(CCur(lstrTotalCardPayments))
                        lbooDeductFromCreditCard = True
                    Else
                        .strPayment2 = SystemPrice(CCur(lstrTotalCardPayments))
                        lbooDeductFromCreditCard = True
                    End If
                Else
                   
                End If
            End If
            .strReconcilliation = SystemPrice(CCur(CCur(lstrTotalNonCardPayments) + CCur(lstrTotalCardPayments)) - CCur(gstrOrderTotal))
            .strUnderpayment = SystemPrice("0")
            .datDespatchDate = Format(Date, "DD/MMM/YYYY")

            UpdateAdviceNote
            If lbooDeductFromCreditCard = False Then
                AddCashBookEntry "", CCur(.strReconcilliation), lstrName, llngCustomerNum, llngOrderNum, "OUTOFSTOCK"
            End If
        Else
            'if payments are greater than order total
            .strUnderpayment = SystemPrice(CCur(gstrOrderTotal) - (CCur(lstrTotalNonCardPayments) + CCur(lstrTotalCardPayments)))
            .strReconcilliation = SystemPrice("0")
            .datDespatchDate = Format(Date, "DD/MMM/YYYY")

            UpdateAdviceNote
            If lbooDeductFromCreditCard = False Then
                AddCashBookEntry "", CCur(.strUnderpayment), lstrName, llngCustomerNum, llngOrderNum, "UNDERPAY"
            End If
        End If

        UpdateOrderStatus "B", 0, "S", llngOrderNum

        Busy False, Screen.ActiveForm
        MsgBox "Order status updated to (B)", , gconstrTitlPrefix & "Calculate Refund"
    End With
Exit Sub
ErrHandler:
    Busy False, Screen.ActiveForm
    
    Select Case GlobalErrorHandler(Err.Number, "CalculateRefund", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Sub XBatchRefundCalculate()
'Ran once only to calculate Refunds for all orders!
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    lstrSQL = "SELECT AdviceNotes.CustNum, AdviceNotes.OrderNum, AdviceNotes.OrderStatus, AdviceNotes.CallerSurname " & _
        "From AdviceNotes WHERE (((AdviceNotes.OrderStatus)='D' Or (AdviceNotes.OrderStatus)='E' Or (AdviceNotes.OrderStatus)='C' Or (AdviceNotes.OrderStatus)='B'));"
     
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        Do Until .EOF
             
            Select Case .Fields("OrderStatus")
            Case "D", "C"
                CalculateRefund .Fields("CustNum"), .Fields("OrderNum"), False
            Case "E", "B"
                CalculateRefund .Fields("CustNum"), .Fields("OrderNum"), True
            End Select
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "XBatchRefundCalculate", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select


End Sub

Function FormatCardNum(pstrCardNum As String) As String
Dim lstrRetCardString As String

    ShowStatus 69
    
    pstrCardNum = Trim$(pstrCardNum)
    
    Select Case Len(Trim$(pstrCardNum))
    Case 18
        lstrRetCardString = Mid$(pstrCardNum, 1, 6) & " " & _
            Mid$(pstrCardNum, 7, 12)
    Case 16
        lstrRetCardString = Mid$(pstrCardNum, 1, 4) & " " & _
            Mid$(pstrCardNum, 5, 4) & " " & _
            Mid$(pstrCardNum, 9, 4) & " " & _
            Mid$(pstrCardNum, 13, 4) & " "
    Case 13
        lstrRetCardString = Mid$(pstrCardNum, 1, 4) & " " & _
            Mid$(pstrCardNum, 5, 3) & " " & _
            Mid$(pstrCardNum, 8, 3) & " " & _
            Mid$(pstrCardNum, 11, 3) & " "
    Case Else
        lstrRetCardString = pstrCardNum
    End Select
    
    FormatCardNum = lstrRetCardString

End Function

Sub UpdateCheqNumAndPrinted()
Dim lstrSQL As String
Dim lintCheqCount As String
Dim lintArrInc As Integer

    ShowStatus 71
    On Error GoTo ErrHandler
    
    For lintArrInc = 0 To UBound(glngChequeOrderNumPrinted)
       
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET ChequePrintedDate = #" & Format$(Now(), "DD/MMM/YYYY") & _
            "#, ChequeNum = " & glngChequeOrderNumPrinted(lintArrInc).lngChequeNum & _
            ", ChequePrintedBy = '" & gstrGenSysInfo.strUserName & "' " & _
            "WHERE Reason <>'UNDERPAY' " & _
            "AND OrderNum =" & glngChequeOrderNumPrinted(lintArrInc).lngOrderNum & ";"
                       
        gdatCentralDatabase.Execute lstrSQL
    Next lintArrInc
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateCheqNumAndPrinted", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub UpdateAdviceBankRepDate()
Dim lstrSQL As String
Dim lintCheqCount As String
Dim lintArrInc As Integer
Dim ldatNow As Date

    ldatNow = Now()
    ShowStatus 72
    On Error GoTo ErrHandler
    
    For lintArrInc = 0 To UBound(glngTrackNUpdate)
        lstrSQL = "UPDATE " & gtblAdviceNotes & " SET " & gtblAdviceNotes & ".BankRepPrintDate = #" & Format$(ldatNow, "DD/MMM/YYYY") & "# " & _
            "WHERE (((" & gtblAdviceNotes & ".OrderNum)=" & glngTrackNUpdate(lintArrInc).lngOrderNum & "));"

        gdatCentralDatabase.Execute lstrSQL
        ErrorLogging Now() & " " & gstrGenSysInfo.strUserName & " " & lstrSQL
    Next lintArrInc
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateAdviceBankRepDate", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Sub GetCurrentStockBatchNumber()
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    ShowStatus 73
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT * FROM " & gtblSystem & ";"

    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            Select Case Trim$(UCase$(.Fields("Item")))
            Case "STOCKBATCHINCR"
                glngStockBatchNumber = CLng(Trim$(.Fields("Value")))
                If IsNull(.Fields("OtherDate")) Then
                    gdatLastStockBatchNumberDate = 0
                Else
                    gdatLastStockBatchNumberDate = CDate(.Fields("OtherDate") & "")
                End If
            End Select
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
       glngStockBatchNumber = 0
    End If
        
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetCurrentStockBatchNumber", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Sub AddFirstStockBatchIncr()
Dim lstrSQL As String

    ShowStatus 74
    On Error GoTo ErrHandler
    
    lstrSQL = "INSERT INTO " & gtblSystem & " ( [Value], Item, OtherDate ) VALUES(10000,'StockBatchIncr',now());"
    
    gdatLocalDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddFirstStockBatchIncr", "LOCAL", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
End Sub
Sub UpdateStockBatchIncr()
Dim lstrSQL As String
    
    ShowStatus 75
    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblSystem & " SET " & gtblSystem & ".[Value] = '" & glngStockBatchNumber & _
    "', OtherDate=now()  WHERE (((" & gtblSystem & ".Item)='StockBatchIncr'));"
    gdatLocalDatabase.Execute lstrSQL


Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateStockBatchIncr", "LOCAL", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        LastChanceCafe
    Case Else
        Resume Next
    End Select
End Sub
Sub RepStockBatch(pstrFilename As String, plngStockBatchNumber As Long)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrPrintString As String
Dim lstrVAT As String

    lstrVAT = gstrVATRate / 100

    On Error GoTo ErrHandler
    lstrSQL = "SELECT Sum(IIf(Trim$([TaxCode])='Z',([DespQty]*[Price]), " & _
        "([DespQty]*[Price])+([DespQty]*[Price])*" & lstrVAT & ")) AS GoodsTotal, " & _
        "Sum(IIf(Trim$([TaxCode])='Z',0,([DespQty]*[Price])*" & lstrVAT & ")) AS VAT, " & _
        "" & gtblAdviceNotes & ".CustNum, " & gtblAdviceNotes & ".OrderNum, " & gtblAdviceNotes & ".CallerInitials, " & _
        "" & gtblAdviceNotes & ".CallerSurname, " & gtblAdviceNotes & ".DespatchDate, " & gtblAdviceNotes & ".StockBatchNum, " & _
        "" & gtblAdviceNotes & ".OrderStatus FROM " & gtblMasterOrderLines & " LEFT JOIN " & gtblAdviceNotes & " ON " & _
        "" & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum GROUP BY " & gtblAdviceNotes & ".CustNum, " & _
        "" & gtblAdviceNotes & ".OrderNum, " & gtblAdviceNotes & ".CallerInitials, " & gtblAdviceNotes & ".CallerSurname, " & _
        "" & gtblAdviceNotes & ".DespatchDate, " & gtblAdviceNotes & ".StockBatchNum, " & gtblAdviceNotes & ".OrderStatus " & _
        "Having (((" & gtblAdviceNotes & ".StockBatchNum) = " & plngStockBatchNumber & _
        ")) ORDER BY " & gtblAdviceNotes & ".OrderNum;"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        
        lstrPrintString = "StockBatchNum" & vbTab & "CustNum,OrderNum" & vbTab & _
            "CallerInitials" & vbTab & "CallerSurname" & vbTab & "" & _
            "VAT" & vbTab & "GoodsTotal" & vbTab & "DespatchDate" & vbTab & "OrderStatus"
        
        Print #lintFileNum, lstrPrintString
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
                
            lstrPrintString = .Fields("StockBatchNum") & vbTab & .Fields("CustNum") & vbTab & _
                .Fields("OrderNum") & vbTab & .Fields("CallerInitials") & vbTab & _
                .Fields("CallerSurname") & vbTab & .Fields("VAT") & vbTab & _
                .Fields("GoodsTotal") & vbTab & .Fields("DespatchDate") & vbTab & _
                .Fields("OrderStatus")
            
            Print #lintFileNum, lstrPrintString
            .MoveNext
        Loop
        
        Close lintFileNum
    End With
    
    If llngRecCount <> 0 Then
        MsgBox "Stock batch report created :- " & vbCrLf & vbTab & pstrFilename & vbCrLf & vbCrLf & "Please inform Finance!", vbInformation, gconstrTitlPrefix & "Stock Batch Report"
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RepStockBatch", "Central")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select
End Sub
Function CheckForSubstitutions(pstrCatNum As String) As String
'Search substituions table field SubCatNum for pstrCatNum
'If TranType = "AUTO" and match found, bass CatNum back
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    CheckForSubstitutions = pstrCatNum
    
    lstrSQL = "SELECT * " & _
        "FROM " & gtblSubstitutions & " WHERE trim$(ucase$(SubCatNum)) = '" & _
        Trim$(UCase$(pstrCatNum)) & "';"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            Select Case UCase$(Trim(.Fields("TranType")))
            Case "AUTO"
                CheckForSubstitutions = .Fields("CatNum")
            Case "STOCKAUTO"
                CheckForSubstitutions = "STOCKAUTO#" & .Fields("CatNum")
            End Select
            .MoveNext
        Loop

    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckForAutoSubstitutions", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Function CheckStockSpecificProduct(pstrCatNum As String) As Long
Dim lsnaLists As Recordset
Dim lstrSQL As String

    On Error GoTo ErrHandler

    lstrSQL = "SELECT * FROM " & gtblProducts & " WHERE trim$(ucase$(CatNum)) = '" & _
        Trim$(UCase$(pstrCatNum)) & "';"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        
        If Not .EOF Then
            CheckStockSpecificProduct = CLng(.Fields("NumInStock"))
        End If
    End With

    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckStockSpecificProduct", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function
Sub GetReferenceInfo(Optional pbooNoStatus As Variant)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrTempVar As String
   
    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 105
    End If
    
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".ListCode, " & _
        "" & gtblListDetails & ".Description FROM " & gtblLists & " INNER JOIN " & _
        "" & gtblListDetails & " ON " & gtblLists & ".ListNum = " & gtblListDetails & ".ListNum " & _
        "WHERE (((" & gtblLists & ".SysUse)=True));"
    
            
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With gstrReferenceInfo
        llngRecCount = 0
        
        Do Until lsnaLists.EOF
            Select Case lsnaLists.Fields("ListName")
            Case "Amount Levels"
                If lsnaLists.Fields("ListCode") = "DENOM" Then .strDenomination = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "POWAIVER" Then .strPostageWaiveratio = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "VAT" Then .strVATRate175 = lsnaLists.Fields("Description")
            Case "Company Address"
                If lsnaLists.Fields("ListCode") = "CONAME" Then .strCompanyName = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COADLI1" Then .strCompanyAddLine1 = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COADLI2" Then .strCompanyAddLine2 = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COADLI3" Then .strCompanyAddLine3 = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COADLI4" Then .strCompanyAddLine4 = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COADLI5" Then .strCompanyAddLine5 = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COCONTA" Then .strCompanyContact = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "COTELEP" Then .strCompanyTelephone = lsnaLists.Fields("Description")
            Case "Card Serv Header"
                If lsnaLists.Fields("ListCode") = "CSHEAD1A" Then .strCreditCardClaimsHead1A = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "CSHEAD1B" Then .strCreditCardClaimsHead1B = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "CSHEAD2A" Then .strCreditCardClaimsHead2A = lsnaLists.Fields("Description")
            Case "System Various"
                If lsnaLists.Fields("ListCode") = "DONAVAIL" Then
                    lstrTempVar = lsnaLists.Fields("Description") & ""
                    If Trim$(lstrTempVar) = "" Then lstrTempVar = False
                    .booDonationAvail = CBool(lstrTempVar)
                End If
                If lsnaLists.Fields("ListCode") = "STCKTHRE" Then
                    lstrTempVar = lsnaLists.Fields("Description") & ""
                    If Trim$(lstrTempVar) = "" Then lstrTempVar = False
                    .booStockThreashold = CBool(lstrTempVar)
                End If
            End Select
            lsnaLists.MoveNext
        Loop

        If Trim$(.strVATRate175) = "" Then
            .strVATRate175 = "17.5"
        End If
        
        If Trim$(.strDenomination) = "" Then
            .strDenomination = "�"
        End If
        
        If Trim$(.strPostageWaiveratio) = "" Then
            .strPostageWaiveratio = .strDenomination & "75"
        Else
            If Left$(.strPostageWaiveratio, 1) = "�" And .strDenomination <> "�" Then
                .strPostageWaiveratio = .strDenomination & Right$(.strPostageWaiveratio, Len(.strPostageWaiveratio) - 1)
            End If
        End If
        
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            .booDonationAvail = True
        End If
        
        If gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
            .booStockThreashold = True
        End If
                
    End With
        
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetReferenceInfo", "Local", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select


End Sub
Function FindOrderDetailFieldError() As Boolean

    With gstrAdviceNoteOrder
        If Trim$(.strPaymentType1) = "" Then
            'If no payment for paymenttype1
            FindOrderDetailFieldError = True
        Else
            FindOrderDetailFieldError = False
        End If
        
        If FindOrderDetailFieldError = True Then
            If Trim(.strPaymentType2) = "" Then
                FindOrderDetailFieldError = True
                Exit Function
            Else
                FindOrderDetailFieldError = False
            End If
        End If

        If Not IsDate(.datDeliveryDate) And .datDeliveryDate <> "00:00:00" Then
            FindOrderDetailFieldError = True
            Exit Function
        End If
        
        If Trim$(.strCardNumber) <> "" Then
            If Trim$(UCase$(.strPaymentType1)) <> "C" And _
                Trim$(UCase$(.strPaymentType2)) <> "C" Then
                FindOrderDetailFieldError = True
                Exit Function
            End If
        End If
        
        If Trim$(UCase$(.strPaymentType1)) = "C" Or _
            Trim$(UCase$(.strPaymentType2)) = "C" Then
            If Trim$(.strCardNumber) = "" Then
                FindOrderDetailFieldError = True
                Exit Function
            End If
        End If
        
        If Trim$(.strOrderStyle) <> "3" Then
            If Trim$(UCase$(.strPaymentType1)) = "" And _
                Trim$(UCase$(.strPaymentType2)) = "" Then
                FindOrderDetailFieldError = True
                Exit Function
            End If
        End If
    End With

End Function
Function GetAdviceOrderStatus(plngCustNumber As Long, plngOrderNum As Long) As String
Dim ltabAdviceNotes As Recordset

    On Error GoTo ErrHandler

    Set ltabAdviceNotes = gdatCentralDatabase.OpenRecordset(gtblAdviceNotes)
    ltabAdviceNotes.Index = "Advice"
    
    With ltabAdviceNotes
        .Seek "=", plngCustNumber, plngOrderNum
        
        If .NoMatch Then
            Exit Function
        End If
        
        GetAdviceOrderStatus = Trim$(!OrderStatus)
        .Close

    End With

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetAdviceOrderStatus", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Function CountCustAccounts() As Integer
Dim lsnaLists As Recordset
Dim lstrSQL As String
    
    On Error GoTo ErrHandler
    
    CountCustAccounts = 0
    
    lstrSQL = "SELECT Count(CustNum) AS Counter FROM " & gtblCustAccounts & ";"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            CountCustAccounts = Val(.Fields("Counter") & "")
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CountCustAccounts", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function

Function NewPADAvailable(Optional pbooNoStatus As Variant) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 27
    End If
    
    On Error GoTo ErrHandler
            
    lstrSQL = "SELECT " & gtblSystem & ".Item, " & gtblMachine & ".Item, " & gtblSystem & ".Value as SVal, " & gtblMachine & ".Value as MVal " & _
        "From " & gtblSystem & ", " & gtblMachine & " WHERE (((" & gtblSystem & ".Item)='PADUploaded') " & _
        "AND ((" & gtblMachine & ".Item)='PADDownloaded'));"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            llngRecCount = llngRecCount + 1
            If CDate(.Fields("SVal")) > CDate(.Fields("MVal")) Then
                NewPADAvailable = True
            Else
                NewPADAvailable = False
            End If
        End If
    End With
        
    lsnaLists.Close
    
    If llngRecCount = 0 Then
        NewPADAvailable = True
        gdatLocalDatabase.Execute "INSERT INTO Machine (Item) VALUES('PADDownloaded');"
    End If
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "NewPADAvailable", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Function AccountBalance(plngCustNum As Long) As String
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim lstrBalance As String

    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT CustNum, Sum([Payment]+[Payment2]) AS SumofPayms, " & _
        "Sum(Underpayment) AS SumOfUnderpayment, Sum(Reconcilliation) AS " & _
        "SumOfReconcilliation, Sum(TotalIncVat) AS SumOfTotalIncVat FROM " & _
        gtblAdviceNotes & " GROUP BY CustNum HAVING (((CustNum)=" & plngCustNum & "));"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            lstrBalance = (.Fields("SumOfTotalIncVat") - .Fields("SumofPayms"))
        End If
    End With
        
    lsnaLists.Close
    
    AccountBalance = SystemPrice(lstrBalance)
        
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AccountBalance", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function
