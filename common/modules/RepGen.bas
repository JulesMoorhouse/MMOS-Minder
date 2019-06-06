Attribute VB_Name = "modReportGen"
Option Explicit
Sub PrintObjAdviceNotesGeneral(pdatStartDate As Date, pdatEndDate As Date, pstrParamater As String, _
    Optional plngOrderNum As Variant, Optional plngEndOrderNum As Variant, _
    Optional pstrSpecificOrderStatus As Variant)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrAdviceNoteType As String
Dim lstrOrderStatus As String
'Converted table names to constants
'Also removed unnecessary table. references

    On Error GoTo ErrHandler

    lstrSQL = "SELECT DeliveryDate, CustNum, OrderNum " & _
        "FROM " & gtblAdviceNotes & " "
    Select Case pdatStartDate
    Case 0
        If pdatEndDate <> 0 Then
            lstrSQL = lstrSQL & "WHERE DeliveryDate " & _
                "<= #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "# "
        Else ' S
            lstrSQL = lstrSQL & "Where "
        End If
    End Select
       
    ReDim gstrOrderLineParcelNumbers(0)
    
    Select Case pstrParamater
    Case "A" 'Awaiting
        lstrSQL = lstrSQL & " and OrderStatus = 'A' "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
        lstrAdviceNoteType = ""
    Case "P" 'Printed
        lstrSQL = lstrSQL & " and OrderStatus = 'P' "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
        lstrAdviceNoteType = ""
    Case "S" 'Specific
        lstrSQL = lstrSQL & " Ordernum = " & plngOrderNum & " "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
        If Not IsMissing(pstrSpecificOrderStatus) Then
            If pstrSpecificOrderStatus = "R" Then
                lstrAdviceNoteType = "REFUND"
            End If
        End If
    Case "R" 'Range print
        'Added 
        lstrSQL = lstrSQL & " (((OrderStatus)='A') AND " & _
            "((OrderNum)>=" & plngOrderNum & _
            " And (OrderNum)<=" & plngEndOrderNum & ")) "
        lstrAdviceNoteType = ""
    Case "O" 'Specific Order Status
        lstrSQL = lstrSQL & " and OrderStatus = '" & pstrSpecificOrderStatus & "' "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
        If pstrSpecificOrderStatus = "R" Then 
            lstrAdviceNoteType = "REFUND"
        End If
    Case Else
       
    End Select
        
    lstrSQL = lstrSQL & "order by OrderNum;" 'CreationDate;"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        llngRecCount = 0
        glngLastOrderPrintedInThisRun = 0
        
        'Open write file put
        Dim lintFreeFile As Integer
        lintFreeFile = FreeFile
        Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)
    
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            If pstrParamater <> "S" Then
                ClearAdviceNote
                ClearCustomerAcount
                ClearGen
            End If
            
            GetAdviceNote .Fields("CustNum"), .Fields("OrderNum")
            
            If pstrParamater = "S" And IsMissing(pstrSpecificOrderStatus) Then
                lstrOrderStatus = GetAdviceOrderStatus(.Fields("CustNum"), .Fields("OrderNum"))
                If lstrOrderStatus = "R" Then
                    lstrAdviceNoteType = "REFUND"
                End If
            End If
            
            If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 Then
                GetRemark gstrAdviceNoteOrder.lngAdviceRemarkNum, gstrInternalNote
            End If
            If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 Then
                GetRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote
            End If
            
            'PrintObjAdviceInvoice "CENTRAL", "", lintFreeFile
            PrintObjAdviceInvoice "CENTRAL", lstrAdviceNoteType, lintFreeFile
            
            glngLastOrderPrintedInThisRun = .Fields("OrderNum")
            
            If glngItemsWouldLikeToPrint > 0 Then
                If glngItemsWouldLikeToPrint = llngRecCount Then
                    lsnaLists.Close
                    Close #lintFreeFile ' Added 19/10/01
                    Exit Sub
                End If
            End If
            
            .MoveNext
        Loop
        
        Close #lintFreeFile
    End With
    
    If llngRecCount = 0 Then
        MsgBox "No records found within the specified criteria.", , gconstrTitlPrefix & "General Advice Note Print"
    End If
    
    lsnaLists.Close
        
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjAdviceNotesGeneral", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select

End Sub

Sub PrintObjAdviceInvoice(pstrMode As String, pstrAdviceType As String, pintFreeFile As Integer)
Dim pstrProductLines() As OrderDetail
Dim llngNumOfProductLines As Long
Dim pstrVatTotal As String: Dim pstrUnitTotal As String: Dim pstrZeroUnitTotal As String
Dim lintArrInc As Integer: Dim lintArrInc2 As Integer: Dim lstrCourier As String
Dim lstrInternalNote As String: Dim lstrAdviceNote As String
Dim lstrPFServiceInd As String
Dim llngMaxproductLines As Long
Dim lintDetailInc As Integer
'Const lconintNumberOflineinHeader = 5: Const lconintNumberOflineinMiddle = 18: Const lconintNumberOflineinFooter = 13
Const lconintNumberOflineinHeader = 5: Const lconintNumberOflineinMiddle = 13: Const lconintNumberOflineinFooter = 11
    
    llngMaxproductLines = lconintNumberOflineinHeader + lconintNumberOflineinMiddle + lconintNumberOflineinFooter
    llngMaxproductLines = gintNumberOfLinesAPage - llngMaxproductLines
   
    If IsMissing(pstrAdviceType) Then pstrAdviceType = ""
    
    gstrAdviceServiceInd.strListName = "PForce Service Indicator"
    
    
    PrintGetProductsOrderlines gstrAdviceNoteOrder.lngCustNum, pstrProductLines(), pstrUnitTotal, pstrVatTotal, gstrAdviceNoteOrder.lngOrderNum, pstrMode, pstrZeroUnitTotal, llngMaxproductLines

    
    If pstrProductLines(0).booLightParcel = False Then
'        UpdateAdviceParcelExtras gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum, UBound(pstrProductLines) + 1, glngGrossWeight
       
        UpdateAdviceParcelExtras gstrAdviceNoteOrder.lngCustNum, _
            gstrAdviceNoteOrder.lngOrderNum, CInt(glngTotalParcel), glngGrossWeight
    End If
    
    For lintArrInc2 = 0 To UBound(pstrProductLines)
        gintCurrentReportPageNum = gintCurrentReportPageNum + 1
        If pstrProductLines(lintArrInc2).lngNumberOfLines = "" Then
            If pstrProductLines(lintArrInc2).strProductLines = "" Then
                If UBound(pstrProductLines) = lintArrInc2 Then
                    'New parcel but no lines  'Shouldn't happen
                End If
            End If
        End If
        
        PrintObjAdviceInvoiceHeader pstrAdviceType, lstrCourier, lstrPFServiceInd, _
            pstrProductLines(), lintArrInc2
        
        PrintObjAdviceInvoiceMiddle lstrCourier, _
            pstrProductLines(lintArrInc2).lngParcelBoxNumber, glngTotalParcel, _
            (lintArrInc2 + 1), UBound(pstrProductLines) + 1

        For lintDetailInc = 0 To UBound(gstrReportData.strDetails, 2)
            If lintDetailInc > UBound(pstrProductLines(lintArrInc2).strLines) Then
                Exit For
            End If
                
            gstrReportData.strDetails(lconPrdCatCode, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strCatCode)
            gstrReportData.strDetails(lconPrdBinLoc, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strBinLoc)
            gstrReportData.strDetails(lconPrdQtyOrd, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strQtyOrd)
            gstrReportData.strDetails(lconPrdQtyDesp, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strQtyDesp)
            gstrReportData.strDetails(lconPrdDesc, lintDetailInc).strValue = (Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strDesc))
            gstrReportData.strDetails(lconPrdUnitPrice, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strUnitPrice)
            gstrReportData.strDetails(lconPrdTaxCode, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strTaxCode)
            gstrReportData.strDetails(lconPrdAmount, lintDetailInc).strValue = Trim$(pstrProductLines(lintArrInc2).strLines(lintDetailInc).strAmount)
        Next lintDetailInc
        PrintObjAdviceInvoiceFooter pstrUnitTotal, pstrZeroUnitTotal, pstrVatTotal, pstrAdviceType
        
        For lintArrInc = 1 To (gintNumberOfLinesAPage - (lconintNumberOflineinHeader + lconintNumberOflineinMiddle + lconintNumberOflineinFooter)) - pstrProductLines(lintArrInc2).lngNumberOfLines
            lstrAdviceNote = lstrAdviceNote & "" & vbCrLf
        Next lintArrInc
    
        'Put file
        Put #pintFreeFile, gintCurrentReportPageNum, gstrReportData
    Next lintArrInc2

End Sub
Function PrintObjAdviceInvoiceHeader(pstrAdviceType As String, ByRef lstrCourier As String, _
    ByRef lstrPFServiceInd As String, ByRef pstrProductLines() As OrderDetail, lintArrInc2 As Integer) As String

   
    gstrReportData.strHeaders(lconCompName).strValue = gstrReferenceInfo.strCompanyName
    gstrReportData.strHeaders(lconCompAdd1).strValue = gstrReferenceInfo.strCompanyAddLine1
    gstrReportData.strHeaders(lconCompAdd2).strValue = gstrReferenceInfo.strCompanyAddLine2
    gstrReportData.strHeaders(lconCompAdd3).strValue = gstrReferenceInfo.strCompanyAddLine3
    gstrReportData.strHeaders(lconCompAdd4).strValue = gstrReferenceInfo.strCompanyAddLine4
    gstrReportData.strHeaders(lconCompAdd5).strValue = gstrReferenceInfo.strCompanyAddLine5
            
    If UCase$(App.ProductName) = "LITE" Then
        If Trim$(gstrReportData.strHeaders(lconCompName).strValue) = "" Or _
            Trim$(gstrReportData.strHeaders(lconCompAdd1).strValue) = "" Then
            MsgBox "You must setup and company name and address, from the main screen" & vbCrLf & _
                "on the Tools menu select Essential Settings!"
        End If
    End If
            
    With gstrAdviceNoteOrder
        If Trim$(pstrAdviceType) = "REFUND" Then
            gstrReportData.strHeaders(lconAdviceTitle).strValue = "REFUND ADVICE NOTE"
        Else
            gstrReportData.strHeaders(lconAdviceTitle).strValue = "ADVICE NOTE"
        End If
        
        If Asc(Left$(gstrPForceServiceInd.strUserDef1, 1)) = 0 Then
            lstrPFServiceInd = ""
        Else
            lstrPFServiceInd = gstrPForceServiceInd.strUserDef1
        End If
        
        If pstrProductLines(lintArrInc2).booLightParcel = True Then
            lstrCourier = "** LETTER POST **"
        Else
            If Trim$(.strCourierCode) = "" Or Trim$(.strCourierCode) = "PF" Then
                gstrAdviceServiceInd.strListCode = "SND"
                GetListVarsAll gstrAdviceServiceInd
                lstrCourier = "** PF" & Trim(lstrPFServiceInd) & " **"
            ElseIf Trim$(.strCourierCode) <> "PF 48" And Trim$(.strCourierCode) <> "" Then
                gstrAdviceServiceInd.strListCode = Trim$(.strCourierCode)
                GetListVarsAll gstrAdviceServiceInd
                lstrCourier = "** PF" & Trim(lstrPFServiceInd) & " **"
            Else
                lstrCourier = "** " & Trim$(.strCourierCode) & " **"
            End If
        End If
    End With
    
     gstrReportData.strHeaders(lconDeliverTitle).strValue = "Deliver to:      "
     gstrReportData.strHeaders(lconDelivServ).strValue = lstrCourier
     
End Function
Function PrintObjAdviceInvoiceMiddle(pstrCourier As String, plngParcelNumber As Long, _
    plngTotalParcels As Long, pintPageNumber As Integer, pintTotalpages As Integer) As String
Dim lstrName As String
Dim lstrDeliveryName As String

    With gstrAdviceNoteOrder
        lstrName = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & Trim$(.strDeliveryInitials) & " " & Trim$(.strDeliverySurname))
        If Trim$(lstrDeliveryName) = "" Then lstrDeliveryName = lstrName
        gstrReportData.strHeaders(lconCustAddName).strValue = lstrName
        gstrReportData.strHeaders(lconDeliverAddName).strValue = lstrDeliveryName
        gstrReportData.strHeaders(lconCustAddLine1).strValue = .strAdd1
        gstrReportData.strHeaders(lconCustAddLine2).strValue = .strAdd2
        gstrReportData.strHeaders(lconCustAddLine3).strValue = .strAdd3
        gstrReportData.strHeaders(lconCustAddLine4).strValue = .strAdd4
        gstrReportData.strHeaders(lconCustAddLine5).strValue = .strAdd5
        gstrReportData.strHeaders(lconCustAddPCode).strValue = .strPostcode
        gstrReportData.strHeaders(lconDeliverAddLine1).strValue = .strDeliveryAdd1
        gstrReportData.strHeaders(lconDeliverAddLine2).strValue = .strDeliveryAdd2
        gstrReportData.strHeaders(lconDeliverAddLine3).strValue = .strDeliveryAdd3
        gstrReportData.strHeaders(lconDeliverAddLine4).strValue = .strDeliveryAdd4
        gstrReportData.strHeaders(lconDeliverAddLine5).strValue = .strDeliveryAdd5
        gstrReportData.strHeaders(lconDeliverAddPCode).strValue = .strDeliveryPostcode
        gstrReportData.strHeaders(lconMediaInfo).strValue = Trim$(.strMediaCode) & " " & GetListCodeDesc("Media Codes", .strMediaCode)
        gstrReportData.strHeaders(lconCustNum).strValue = "M" & .lngCustNum
        gstrReportData.strHeaders(lconOrderNum).strValue = .lngOrderNum
        gstrReportData.strHeaders(lconPageNum).strValue = pintPageNumber & "/" & pintTotalpages
        If .datDeliveryDate = "00:00:00" Then
            gstrReportData.strHeaders(lconOrderDate).strValue = Format(.datCreationDate, "DD/MM/YYYY")
            gstrReportData.strHeaders(lconShipDate).strValue = "Not Specified"
            gstrReportData.strHeaders(lconParcelNum).strValue = plngParcelNumber & "/" & plngTotalParcels
        Else
            gstrReportData.strHeaders(lconOrderDate).strValue = Format(.datCreationDate, "DD/MM/YYYY")
            gstrReportData.strHeaders(lconShipDate).strValue = Format$(.datDeliveryDate, "DD/MM/YYYY")
            gstrReportData.strHeaders(lconParcelNum).strValue = plngParcelNumber & "/" & plngTotalParcels
        
        End If
    End With
    
End Function
Function PrintObjAdviceInvoiceFooter(pstrUnitTotal As String, pstrZeroUnitTotal As String, pstrVatTotal As String, pstrAdviceType As String) As String
Dim lstrConsignmentNoteLine1 As String
Dim lstrConsignmentNoteLine2 As String
Dim lstrTaxSummaryLine1 As String
Dim lstrTaxSummaryLine2 As String
Dim lstrTaxSummaryLine3 As String
Dim lstrPaymentType2 As String
'Dim lstrInternalNote As String

    With gstrAdviceNoteOrder
        If CCur(AdvicePrice(.strDonation)) > 0 Then
            gstrReportData.strFooters(lconPayReceived).strValue = "Payments received. ***  THANK YOU FOR YOUR DONATION  ***"
        Else
            gstrReportData.strFooters(lconPayReceived).strValue = "Payments received."
        End If

        gstrReportData.strFooters(lconGoodsNVatTot).strValue = AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(pstrZeroUnitTotal)) + (PriceVal(pstrVatTotal)))
        
        If .strPayment <> 0 Then
            gstrReportData.strFooters(lconPay1).strValue = GetListCodeDesc("Payment Method", .strPaymentType1) & " " & AdvicePrice(.strPayment)
        End If
        gstrReportData.strFooters(lconPostNPack).strValue = AdvicePrice(.strPostage)
        
        lstrPaymentType2 = GetListCodeDesc("Payment Method", .strPaymentType2) & " " & AdvicePrice(.strPayment2) & vbCrLf '5
        
       
        gstrReportData.strFooters(lconTitleTaxStndPrcnt).strValue = gstrReferenceInfo.strVATRate175 & "%"
        
        gstrReportData.strFooters(lconTitleTaxStndGoodsNPP).strValue = AdvicePrice(PriceVal(pstrUnitTotal))
        gstrReportData.strFooters(lconTitleTaxStndVat).strValue = AdvicePrice(PriceVal(pstrVatTotal))
        
        gstrReportData.strFooters(lconTitleTaxZeroGoodsNPP).strValue = AdvicePrice(PriceVal(pstrZeroUnitTotal))
        gstrReportData.strFooters(lconTitleTaxZeroVat).strValue = AdvicePrice(PriceVal("0"))
        
        If .strPayment2 <> 0 Then
            gstrReportData.strFooters(lconPay2).strValue = lstrPaymentType2
        Else
            gstrReportData.strFooters(lconPay2).strValue = ""
        End If
        
        If gstrReferenceInfo.booDonationAvail = True Then
            gstrReportData.strFooters(lconDonation).strValue = AdvicePrice(.strDonation)
        Else
            gstrReportData.strFooters(lconDonation).strValue = vbTab
        End If
        
        gstrReportData.strFooters(lconTotal).strValue = AdvicePrice(.strTotalIncVat)
        
        If Trim$(pstrAdviceType) = "REFUND" Then
            gstrReportData.strFooters(lconStndNoteLn1).strValue = "Please find enclosed a refund cheque"
            gstrReportData.strFooters(lconStndNoteLn2).strValue = "for the reason shown on this note."
        Else
            If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 And _
            PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) <= 1 Then
                gstrReportData.strFooters(lconStndNoteLn1).strValue = "Please claim refund with next order"
                gstrReportData.strFooters(lconStndNoteLn2).strValue = "quoting order no. and refund value."
            ElseIf PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 1 Then
                gstrReportData.strFooters(lconStndNoteLn1).strValue = "A cheque will follow for the sum and"
                gstrReportData.strFooters(lconStndNoteLn2).strValue = "reason stated, under separate cover."
            End If
        End If
        'If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 And Trim$(gstrInternalNote.strText) <> "" Then
        '    lstrInternalNote = "Note: " & Trim$(gstrInternalNote.strText)
        'Else
        '    lstrInternalNote = ""
        'End If
        'If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 And Trim$(gstrConsignmentNote.strText) <> "" Then
        '    lstrConsignmentNote = "Note: " & Trim$(gstrConsignmentNote.strText)
        'Else
        '    lstrConsignmentNote = ""
        'End If
        If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 And Trim$(gstrConsignmentNote.strText) <> "" Then
            lstrConsignmentNoteLine1 = GrapLine("Note: " & Trim$(gstrConsignmentNote.strText), 0, 60)
            lstrConsignmentNoteLine2 = GrapLine("Note: " & Trim$(gstrConsignmentNote.strText), 1, 60)
        Else
            lstrConsignmentNoteLine1 = vbTab '""'
            lstrConsignmentNoteLine2 = vbTab '""'
        End If
        
        gstrReportData.strFooters(lconConsNoteLn1).strValue = lstrConsignmentNoteLine1
        gstrReportData.strFooters(lconConsNoteLn2).strValue = lstrConsignmentNoteLine2
        
        gstrReportData.strFooters(lconOverPay).strValue = AdvicePrice(.strReconcilliation)
        gstrReportData.strFooters(gconUnderPay).strValue = AdvicePrice(.strUnderpayment)
        
        If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 Then 'Added 20/05/00 by request
            gstrReportData.strFooters(lconTotRefund).strValue = AdvicePrice(PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment))
        Else
            gstrReportData.strFooters(lconTotRefund).strValue = AdvicePrice("0.00")
        End If
    End With
    
End Function
Sub PrintObjChequeRefundAdviceNotes(pdatChequeprintDate As Date)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Dim lintFileNum As Integer
'Converted table names to constants
'Also removed unnecessary table. references
    
    'This function works by specifiy the day the cheques were printed!
    On Error GoTo ErrHandler

    ReDim gstrOrderLineParcelNumbers(0)
    
    'lstrSQL = "SELECT *, " & _
        "Format$([PrintedDate],'dd/mm/yy') AS Expr1 From " & gtblCashBook & " " & _
        "WHERE (((Amount)<>0) AND " & _
        "((Format$([PrintedDate],'dd/mm/yy'))=CDate('" & Format$(pdatChequeprintDate, "dd/mm/yy") & "')) " & _
        "AND ((Reason)<>'UNDERPAY')) ORDER BY OrderNum;"
   
    lstrSQL = "SELECT CustNum, OrderNum, " & _
        "Format$([ChequePrintedDate],'dd/mm/yy') AS Expr1 From " & gtblAdviceNotes & " " & _
        "WHERE (((Underpayment)<>0) AND " & _
        "((Format$([ChequePrintedDate],'dd/mm/yy'))=CDate('" & Format$(pdatChequeprintDate, "dd/mm/yy") & "')) " & _
        "AND ((RefundReason)<>'UNDERPAY')) ORDER BY OrderNum;"
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
    
        llngRecCount = 0
        glngLastOrderPrintedInThisRun = 0
        
        'Open write file put
        Dim lintFreeFile As Integer
        lintFreeFile = FreeFile
        Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)
            
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            ClearAdviceNote
            ClearCustomerAcount
            ClearGen
            
            GetAdviceNote .Fields("CustNum"), .Fields("OrderNum")
            If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 Then
                GetRemark gstrAdviceNoteOrder.lngAdviceRemarkNum, gstrInternalNote
            End If
            If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 Then
                GetRemark gstrAdviceNoteOrder.lngConsignRemarkNum, gstrConsignmentNote
            End If
'            PrintToFileAdviceInvoice pstrFilename, "CENTRAL", "REFUND"
            PrintObjAdviceInvoice "CENTRAL", "REFUND", lintFreeFile
            glngLastOrderPrintedInThisRun = .Fields("OrderNum")
            .MoveNext
        Loop
                
        'close put file
        Close #lintFreeFile
        
    End With
    
    If llngRecCount = 0 Then
        MsgBox "No Refunds to print Advice notes", , gconstrTitlPrefix & "Refund Advice Notes"
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjChequeRefundAdviceNotes", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select
            
End Sub
Sub PrintObjPForceManifestGeneral()
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim lstrLastServiceIDGrouping As String
Dim llngPageNumber As Long
Const lconHeaderLines = 11: Const lconFooterLines = 5
Dim lintNumberOfDetailLinesAllowed As Integer
Dim lintDetailInc As Integer
Dim llngNumConsignmentThisPage As Long
'Converted table names to constants

    ShowStatus 98
    
    lintNumberOfDetailLinesAllowed = (gintNumberOfLinesAPage - lconHeaderLines) - lconFooterLines
    
    llngPageNumber = 1
    lintDetailInc = 0
    llngNumConsignmentThisPage = 0
    Busy True
    On Error GoTo ErrHandler
    lstrSQL = "SELECT OrderNum, CustNum, Status, ServiceID, ConsignNum, ParcelItems From " & _
        "" & gtblPForce & " WHERE (((Status)='P')) ORDER BY ServiceID, ConsignNum;"
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists

        'Open write file put
        Dim lintFreeFile As Integer
        lintFreeFile = FreeFile
        Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)
        
        Do Until .EOF
            'Gain Settings Start
            GetPForceConsignment .Fields("OrderNum"), .Fields("CustNum"), .Fields("ConsignNum")
            
            If lintDetailInc = 0 Then
                PopulateStaticItems
                GetPFContractDetails
                
                gstrPForceServiceInd.strListName = "PForce Service Indicator"
                gstrPForceServiceInd.strListCode = gstrPForceConsignment.strServiceID
                GetListVarsAll gstrPForceServiceInd
                
            End If
            
            
            If lstrLastServiceIDGrouping <> .Fields("ServiceID") And lstrLastServiceIDGrouping <> "" Then
                PrintObjPForceManifestHeader llngPageNumber
                gstrPForceServiceInd.strListName = "PForce Service Indicator"
                gstrPForceServiceInd.strListCode = gstrPForceConsignment.strServiceID
                GetListVarsAll gstrPForceServiceInd
                PrintObjPForceManifestFooter llngNumConsignmentThisPage, lintDetailInc
                Put #lintFreeFile, llngPageNumber, gstrReportData
                ClearReportingDataType "D"
                lintDetailInc = 0
                llngPageNumber = llngPageNumber + 1
                llngNumConsignmentThisPage = 0
            End If
            
                
            PrintObjPForceManifestDetail lintDetailInc
            llngNumConsignmentThisPage = llngNumConsignmentThisPage + gstrPForceConsignment.intParcelItems
            
            If lintDetailInc >= lintNumberOfDetailLinesAllowed Then
                PrintObjPForceManifestHeader llngPageNumber
            
                PrintObjPForceManifestFooter llngNumConsignmentThisPage, lintDetailInc + 1
                Put #lintFreeFile, llngPageNumber, gstrReportData
                ClearReportingDataType "D"
                lintDetailInc = -1
                llngPageNumber = llngPageNumber + 1
                llngNumConsignmentThisPage = 0
            End If
            
            lintDetailInc = lintDetailInc + 1
            lstrLastServiceIDGrouping = .Fields("ServiceID")
            .MoveNext
        Loop
        gintCurrentReportPageNum = llngPageNumber - 1
       
    End With
    
    If lintDetailInc <> 0 Then
    Else
        MsgBox "No Consignment found!", , gconstrTitlPrefix & "General Manifest Production"
    End If
    
    Close #lintFreeFile
    
    lsnaLists.Close
     Busy False
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjPForceManifestGeneral", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Function PrintObjPForceManifestHeader(plngPageNum As Long) As Long
Dim llngLineNum As Long
    
    With gstrReportData
       
        .strHeaders(lconTitlCompAddressName).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyName)) & ", "
        .strHeaders(lconTitlCompAddress1).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyAddLine1 & ","))
        .strHeaders(lconTitlCompAddress2).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyAddLine2 & ", "))
        .strHeaders(lconTitlCompAddress3).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyAddLine3 & ", "))
        .strHeaders(lconTitlCompAddress4).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyAddLine4 & ", "))
        .strHeaders(lconTitlCompAddress5).strValue = UCase$(Trim$(gstrReferenceInfo.strCompanyAddLine5))
                
                
        .strHeaders(lconDateDesp).strValue = Format$(Now(), "DD/MM/YYYY")
        .strHeaders(lconPFPageNum).strValue = plngPageNum
        .strHeaders(lconServDesc).strValue = Trim$(gstrPForceServiceInd.strDescription) & " Manifest"
    
        .strHeaders(lconContractNum).strValue = gstrPFContractDetails.strContractNumber & _
        " Service " & gstrPForceServiceInd.strUserDef1
    'Print #lintFileNum, "--------- ------------------------------ -------- -------------- -- - - -": llngLineNum = llngLineNum + 1
            
    'PrintObjPForceManifestHeader = 11
    End With
    
End Function
Function PrintObjPForceManifestDetail(plngItemsCount As Integer) As Long
Dim lstrDeliveryName As String
    
    With gstrPForceConsignment
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & _
            Trim$(.strDeliveryInitials) & " " & _
            Trim$(.strDeliverySurname))
        
        gstrReportData.strDetails(lconDetsConsNum, plngItemsCount).strValue = Trim$(.strConsignNum)
        gstrReportData.strDetails(lconDetsDeliverName, plngItemsCount).strValue = lstrDeliveryName
        gstrReportData.strDetails(lconDetsPostCode, plngItemsCount).strValue = Trim$(.strDeliveryPostcode)
        gstrReportData.strDetails(lconDetsSendrRef, plngItemsCount).strValue = Trim$("M" & .lngOrderNum & "/" & .lngCustNum)
        gstrReportData.strDetails(lconDetsItems, plngItemsCount).strValue = Trim$(.intParcelItems)
        gstrReportData.strDetails(lconDetsSHS, plngItemsCount).strValue = Trim$(.strSpecialSatDel)
        gstrReportData.strDetails(lconDetsSHB, plngItemsCount).strValue = Trim$(.strSpecialBookIn)
        gstrReportData.strDetails(lconDetsSHP, plngItemsCount).strValue = Trim$(.strSpecialProof)
        
       ' PrintObjPForceManifestDetail = plngItemsCount '.intParcelItems
    End With
    
End Function
Function PrintObjPForceManifestFooter(plngItemCount As Long, plngRecordCount As Integer) As Long
    
    gstrReportData.strFooters(lconNumCons).strValue = plngRecordCount
    gstrReportData.strFooters(lconNumItems).strValue = plngItemCount
    
    PrintObjPForceManifestFooter = 5
    
End Function
Sub PrintObjCheques(plngFirstChequeNumber As Long)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrAddress As String
Dim lcurPounds As Currency
Dim lstrTensOfThousands As String
Dim lstrThousands As String
Dim lstrHundreds As String
Dim lstrTens As String
Dim lstrUnits As String
Dim lstrAmount As String
'Converted table names to constants
'Also removed unnecessary table. references

    ReDim Preserve glngChequeOrderNumPrinted(0)
    On Error GoTo ErrHandler

    'lstrSQL = "SELECT Amount, Name, OrderNum, " & _
        "PrintedDate From " & gtblCashBook & " " & _
        "WHERE (((Amount)<>0) AND " & _
        "((PrintedDate)=0) AND " & _
        "((Reason)<>'UNDERPAY')) OR " & _
        "(((Amount)<>0) AND " & _
        "((PrintedDate) Is Null) AND " & _
        "((Reason)<>'UNDERPAY')) " & _
        "ORDER BY OrderNum;"
   
    'lstrSQL = "SELECT TotalIncVat AS Amount, CardName AS Name, OrderNum, " & _
        "ChequePrintedDate AS PrintedDate FROM " & gtblAdviceNotes & " WHERE (Underpayment<>0 AND" & _
        "ChequePrintedDate=0 AND RefundReason<>'UNDERPAY' AND OrderType='Q') OR " & _
        "(Underpayment<>0 AND ChequePrintedDate Is Null AND RefundReason<>'UNDERPAY' " & _
        "AND OrderType='Q') ORDER BY OrderNum;"
   
    'lstrSQL = "SELECT TotalIncVat AS Amount, CardName AS Name, OrderNum, " & _
        "ChequePrintedDate AS PrintedDate FROM " & gtblAdviceNotes & " WHERE (Underpayment<>0 AND " & _
        "ChequePrintedDate=0 AND RefundReason<>'UNDERPAY' AND OrderType='Q') OR " & _
        "(Underpayment<>0 AND ChequePrintedDate Is Null AND RefundReason<>'UNDERPAY' " & _
        "AND OrderType='Q') ORDER BY OrderNum;"
   
    lstrSQL = "SELECT TotalIncVat AS Amount, CardName AS Name, OrderNum, " & _
        "ChequePrintedDate AS PrintedDate FROM " & gtblAdviceNotes & " WHERE ( " & _
        "ChequePrintedDate=0 AND RefundReason<>'UNDERPAY' AND OrderType='Q') OR " & _
        "(ChequePrintedDate Is Null AND RefundReason<>'UNDERPAY' " & _
        "AND OrderType='Q') ORDER BY OrderNum;"
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lstrTensOfThousands = ""
        lstrThousands = ""
        lstrHundreds = ""
        lstrTens = ""
        lstrUnits = ""
        'lintFileNum = FreeFile
        'Open pstrFilename For Output As lintFileNum
        
        'Open write file put
        Dim lintFreeFile As Integer
        lintFreeFile = FreeFile
        Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, Spacer("", 57) & Format$(Now(), "DD/MMM/YYYY")
            gstrReportData.strHeaders(lconChqDate).strValue = Format$(Now(), "DD/MMM/YYYY")
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, " " & Spacer(.Fields("OrderNum"), 8) & Spacer(Trim$(.Fields("Name")), 50) & Format(.Fields("Amount"), "0.00")
            gstrReportData.strHeaders(lconChqOrderNum).strValue = .Fields("OrderNum")
            gstrReportData.strHeaders(lconChqName).strValue = Trim$(.Fields("Name"))
            lstrAmount = .Fields("Amount") & ""
            If Left$(lstrAmount, 1) = "-" Then
                lstrAmount = Right$(lstrAmount, Len(lstrAmount & "") - 1)
            End If
            'gstrReportData.strHeaders(lconChqAmount).strValue = Format(.Fields("Amount"), "0.00")
            gstrReportData.strHeaders(lconChqAmount).strValue = Format(lstrAmount, "0.00")
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            'Print #lintFileNum, ""
            
            
            If InStr(1, "" & .Fields("Amount") & "", ".") > 0 Then
                lcurPounds = CCur(Left$("" & .Fields("Amount") & "", InStr(1, "" & .Fields("Amount") & "", ".")))
            Else
                lcurPounds = CCur(.Fields("Amount"))
            End If
            If Left$(lcurPounds, 1) = "-" Then
                lcurPounds = Right$(lcurPounds, Len(lcurPounds & "") - 1)
            End If
            If lcurPounds = 0 Then
                lstrTensOfThousands = "**"
                lstrThousands = "**"
                lstrHundreds = "**"
                lstrTens = "**"
                lstrUnits = "**"
             Else
                Select Case Len(Format$(lcurPounds, "0"))
                Case 1
                    lstrTensOfThousands = "**"
                    lstrThousands = "**"
                    lstrHundreds = "**"
                    lstrTens = "**"
                    lstrUnits = NumToWord(Mid$(lcurPounds, 1, 1))
                Case 2
                    lstrTensOfThousands = "**"
                    lstrThousands = "**"
                    lstrHundreds = "**"
                    lstrTens = NumToWord(Mid$(lcurPounds, 1, 1))
                    lstrUnits = NumToWord(Mid$(lcurPounds, 2, 1))
                Case 3
                    lstrTensOfThousands = "**"
                    lstrThousands = "**"
                    lstrHundreds = NumToWord(Mid$(lcurPounds, 1, 1))
                    lstrTens = NumToWord(Mid$(lcurPounds, 2, 1))
                    lstrUnits = NumToWord(Mid$(lcurPounds, 3, 1))
                Case 4
                    lstrTensOfThousands = "**"
                    lstrThousands = NumToWord(Mid$(lcurPounds, 1, 1))
                    lstrHundreds = NumToWord(Mid$(lcurPounds, 2, 1))
                    lstrTens = NumToWord(Mid$(lcurPounds, 3, 1))
                    lstrUnits = NumToWord(Mid$(lcurPounds, 4, 1))
                Case 5
                    lstrTensOfThousands = NumToWord(Mid$(lcurPounds, 1, 1))
                    lstrThousands = NumToWord(Mid$(lcurPounds, 2, 1))
                    lstrHundreds = NumToWord(Mid$(lcurPounds, 3, 1))
                    lstrTens = NumToWord(Mid$(lcurPounds, 4, 1))
                    lstrUnits = NumToWord(Mid$(lcurPounds, 5, 1))
                End Select
            End If
            
            'Print #lintFileNum, " " & Spacer(lstrTensOfThousands, 5) & " " & _
                Spacer(lstrThousands, 5) & " " & _
                Spacer(lstrHundreds, 5) & " " & _
                Spacer(lstrTens, 5) & " " & _
                Spacer(lstrUnits, 5)
            gstrReportData.strHeaders(lconChqTensOfThousands).strValue = lstrTensOfThousands
            gstrReportData.strHeaders(lconChqThousands).strValue = lstrThousands
            gstrReportData.strHeaders(lconChqHundreds).strValue = lstrHundreds
            gstrReportData.strHeaders(lconChqTens).strValue = lstrTens
            gstrReportData.strHeaders(lconChqUnits).strValue = lstrUnits
                
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""
                'Print #lintFileNum, ""

            If llngRecCount = 1 Then
                'ReDim Preserve glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted) + 1)
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngOrderNum = .Fields("OrderNum")
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngChequeNum = plngFirstChequeNumber
                
            Else
                ReDim Preserve glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted) + 1)
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngOrderNum = .Fields("OrderNum")
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngChequeNum = plngFirstChequeNumber + (llngRecCount - 1)
            
            End If
                'Put file
                Put #lintFreeFile, llngRecCount, gstrReportData
            .MoveNext
        Loop
        gintCurrentReportPageNum = llngRecCount
        Close #lintFreeFile
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjCheques", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select



End Sub
Function PrintObjBatchPickings() As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Dim lintFileNum As Integer
Dim pstrOrderNumBatch() As BatchLines
Dim lintArrInc As Integer
Dim lintLineNum As Integer
Dim lintArrInc2 As Integer
Dim lintMorePages As Integer

Dim llngPageRecordNum As Long

Const lconCatSpacer As Integer = 11
Const lconBinSpacer As Integer = 15
Const lconQtySpacer As Integer = 9
Const lconProdSpacer As Integer = 37
'Converted table names to constants

    PrintObjBatchPickings = True
 
       
    On Error GoTo ErrHandler
    
    'Open write file put
    Dim lintFreeFile As Integer
    lintFreeFile = FreeFile
    Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)
    
    
    RepBatchPickCountOrderWeights pstrOrderNumBatch()
    For lintArrInc = 0 To UBound(pstrOrderNumBatch)
    
        If pstrOrderNumBatch(lintArrInc).strSQLWhereClause = "" Then
            MsgBox "No Orders found!", , gconstrTitlPrefix & "Batch Pickings"
            PrintObjBatchPickings = False
            Close #lintFreeFile
            Exit Function
        End If
        lstrSQL = "SELECT Sum(Val([Qty])) AS Quantity, Sum(Val([TotalWeight])) AS " & _
            "TW, " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & "." & _
            "BinLocation, " & gtblMasterOrderLines & ".ItemDescription " & _
            "From " & gtblMasterOrderLines & _
            " INNER JOIN " & gtblAdviceNotes & " ON " & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum "
        
        lstrSQL = lstrSQL & " where " & pstrOrderNumBatch(lintArrInc).strSQLWhereClause
                

        lstrSQL = lstrSQL & " GROUP BY " & gtblMasterOrderLines & "." & _
            "CatNum, " & gtblMasterOrderLines & ".BinLocation, " & _
            gtblMasterOrderLines & ".ItemDescription "
        lstrSQL = lstrSQL & "ORDER BY " & gtblMasterOrderLines & ".BinLocation;"
        llngPageRecordNum = llngPageRecordNum + 1
        Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
            With lsnaLists
            
                llngRecCount = 0
                gstrReportData.strHeaders(lconBPDate).strValue = Format$(Now(), "DD/MMM/YYYY")
                gstrReportData.strHeaders(lconBPOrderNums).strValue = pstrOrderNumBatch(lintArrInc).strOrderNumberHeader
                lintLineNum = lintLineNum + 8
                Do Until .EOF
                    gstrReportData.strDetails(lconBPDetsCatCode, llngRecCount).strValue = .Fields("CatNum")
                    gstrReportData.strDetails(lconBPDetsBin, llngRecCount).strValue = .Fields("BinLocation")
                    gstrReportData.strDetails(lconBPDetsQty, llngRecCount).strValue = .Fields("Quantity")
                    gstrReportData.strDetails(lconBPDetsProd, llngRecCount).strValue = .Fields("ItemDescription")
                    gstrReportData.strDetails(lconBPDetsWeight, llngRecCount).strValue = Format$(Val(.Fields("TW")) / 1000, "0.000")
                    llngRecCount = llngRecCount + 1
                     lintLineNum = lintLineNum + 1
                    If (lintLineNum) > 50 Then
                        llngRecCount = 0
                        lintLineNum = 0
                        Put #lintFreeFile, llngPageRecordNum, gstrReportData
                        llngPageRecordNum = llngPageRecordNum + 1
                        ClearReportingDataType "D"
                    End If
                    .MoveNext
                
                Loop
            End With
        If llngRecCount = 0 Then
        End If
            
        lsnaLists.Close
        gstrReportData.strFooters(lconBPTotalWeight).strValue = Format$(Val(pstrOrderNumBatch(lintArrInc).strBatchTotal) / 1000, "0.000") & "kg"
         lintLineNum = lintLineNum + 2
         
        If lintLineNum > gintNumberOfLinesAPage Then
            lintMorePages = CInt(lintLineNum / gintNumberOfLinesAPage)
            lintLineNum = gintNumberOfLinesAPage - ((gintNumberOfLinesAPage * lintMorePages) - lintLineNum)
        Else
            lintMorePages = 0
        End If
        
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        lintLineNum = 0
        'Put file
        'llngPageRecordNum = llngPageRecordNum + 1
        Put #lintFreeFile, llngPageRecordNum, gstrReportData
        ClearReportingDataType "D"
        
    Next lintArrInc
    
    gintCurrentReportPageNum = llngPageRecordNum '- 1
    
    Close #lintFreeFile
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjBatchPickings", "Central")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select


End Function
Function PrintObjCreditCardClaim(Optional pbooRefundClaims As Variant) As String
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
'Dim lstrAddress As String
Dim lcurTotal As Currency
Dim lcurPageTotal As Currency
Dim lintArrInc2 As Integer
Dim lintLineNum As Integer
Dim lintPageNum As Integer
Dim lcurTotalCCPayment As Currency
Dim lstrCardNum As String
'Converted table names to constants

    If IsMissing(pbooRefundClaims) Then
        pbooRefundClaims = False
    End If
    
   
    gstrReportData.strHeaders(lconCCCHead1A).strValue = Trim$(gstrReferenceInfo.strCreditCardClaimsHead1A)
    gstrReportData.strHeaders(lconCCCHead1B).strValue = Trim$(gstrReferenceInfo.strCreditCardClaimsHead1B)
    gstrReportData.strHeaders(lconCCCHead2A).strValue = Trim$(gstrReferenceInfo.strCreditCardClaimsHead2A)

    ReDim Preserve glngTrackNUpdate(0)
    On Error GoTo ErrHandler

    'Open write file put
    Dim lintFreeFile As Integer
    lintFreeFile = FreeFile
    Open gstrReport.strDelimDetailsFile For Random As #lintFreeFile Len = Len(gstrReportData)

    'lstrSQL = "SELECT AdviceNotes.CardNumber, AdviceNotes.TotalIncVat, " & _
        "AdviceNotes.AuthorisationCode, AdviceNotes.OrderNum, " & _
        "AdviceNotes.CreationDate, AdviceNotes.DespatchDate, " & _
        "AdviceNotes.OrderType, AdviceNotes.Payment, " & _
        "AdviceNotes.PaymentType2, AdviceNotes.Payment2, " & _
        "AdviceNotes.CallerSalutation, AdviceNotes.CallerInitials, " & _
        "AdviceNotes.CallerSurname, AdviceNotes.AdviceAdd1, " & _
        "AdviceNotes.AdviceAdd2, AdviceNotes.AdviceAdd3, " & _
        "AdviceNotes.AdviceAdd4, AdviceNotes.AdviceAdd5, " & _
        "AdviceNotes.AdvicePostcode , AdviceNotes.BankRepPrintDate " & _
        "From AdviceNotes " & _
        "WHERE (((AdviceNotes.BankRepPrintDate)=0 Or (AdviceNotes.BankRepPrintDate) " & _
        "Is Null) AND ((AdviceNotes.OrderStatus)='C' Or (AdviceNotes.OrderStatus)='B' " & _
        "Or (AdviceNotes.OrderStatus)='D' Or (AdviceNotes.OrderStatus)='E') AND " & _
        "((Trim$([AdviceNotes].[OrderType]))='C')) OR (((AdviceNotes.BankRepPrintDate)=0 " & _
        "Or (AdviceNotes.BankRepPrintDate) Is Null) AND ((AdviceNotes.OrderStatus)='C' Or " & _
        "(AdviceNotes.OrderStatus)='B' Or (AdviceNotes.OrderStatus)='D' Or " & _
        "(AdviceNotes.OrderStatus)='E') AND ((Trim$([AdviceNotes].[PaymentType2]))='C')) " & _
        "ORDER BY AdviceNotes.OrderNum;"

    If pbooRefundClaims = False Then
       
        lstrSQL = "SELECT CardNumber, TotalIncVat, AuthorisationCode, OrderNum, CreationDate, DespatchDate, " & _
            "OrderType, Payment, PaymentType2, Payment2, CallerSalutation, CallerInitials, CallerSurname, " & _
            "AdviceAdd1, AdviceAdd2, AdviceAdd3, AdviceAdd4, AdviceAdd5, AdvicePostcode, BankRepPrintDate, " & _
            "CardType From " & gtblAdviceNotes & " WHERE (((BankRepPrintDate)=0 Or (BankRepPrintDate) Is Null) AND " & _
            "((OrderStatus)='C' Or (OrderStatus)='B' Or (OrderStatus)='D' Or (OrderStatus)='E') AND " & _
            "((Trim$([OrderType]))='C') AND ((CardType)<>'SWITCH')) OR (((BankRepPrintDate)=0 Or " & _
            "(BankRepPrintDate) Is Null) AND ((OrderStatus)='C' Or (OrderStatus)='B' Or (OrderStatus)='D' " & _
            "Or (OrderStatus)='E') AND ((Trim$([PaymentType2]))='C') AND ((CardType)<>'SWITCH')) " & _
            "ORDER BY OrderNum;"
    Else
       
        lstrSQL = "SELECT CardNumber, TotalIncVat, AuthorisationCode, OrderNum, CreationDate, " & _
            "DespatchDate, OrderType, Payment, PaymentType2, Payment2, CallerSalutation, " & _
            "CallerInitials, CallerSurname, AdviceAdd1, AdviceAdd2, AdviceAdd3, AdviceAdd4, " & _
            "AdviceAdd5, AdvicePostcode, BankRepPrintDate, CardType, Reconcilliation, " & _
            "RefundReason FROM " & gtblAdviceNotes & " WHERE (((BankRepPrintDate)=0 Or " & _
            "(BankRepPrintDate) Is Null) AND ((CardType)<>'SWITCH') AND ((OrderStatus)='C' " & _
            "Or (OrderStatus)='B' Or (OrderStatus)='D' Or (OrderStatus)='E') AND " & _
            "((Trim$([OrderType]))='C') AND ((Reconcilliation)<>0) AND " & _
            "((RefundReason)='OVERPAY')) OR (((BankRepPrintDate)=0 Or " & _
            "(BankRepPrintDate) Is Null) AND ((CardType)<>'SWITCH') AND " & _
            "((OrderStatus)='C' Or (OrderStatus)='B' Or (OrderStatus)='D' Or " & _
            "(OrderStatus)='E') AND ((Trim$([PaymentType2]))='C') AND ((Reconcilliation)<>0) " & _
            "AND ((RefundReason)='OVERPAY')) ORDER BY OrderNum;"
    End If
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lcurTotal = 0
        lcurPageTotal = 0
        lintPageNum = 1
        lintFileNum = FreeFile
        
        gstrReportData.strHeaders(lconCCCAsAt).strValue = Now()
        lintLineNum = lintLineNum + 4

        Do Until .EOF
            llngRecCount = llngRecCount + 1
                                
            If UCase$(Trim$(.Fields("OrderType"))) = "C" Then
                lcurTotalCCPayment = CCur(.Fields("Payment"))
            End If
            
            If UCase$(Trim$(.Fields("PaymentType2"))) = "C" Then
                lcurTotalCCPayment = CCur(.Fields("Payment2"))
            End If
            
            If pbooRefundClaims = True Then
                lcurTotalCCPayment = CCur(.Fields("Reconcilliation"))
            End If
            
            lintLineNum = lintLineNum + 1
            
            lstrCardNum = FormatCardNum(.Fields("CardNumber") & "")
            If lstrCardNum = "" Then
                lstrCardNum = "0"
            End If
            gstrReportData.strDetails(lconCCCDetsCardNum, llngRecCount - 1).strValue = lstrCardNum
            lstrCardNum = ""
            gstrReportData.strDetails(lconCCCDetsAmount, llngRecCount - 1).strValue = Trim$(Format$(lcurTotalCCPayment, "0.00") & "")
            gstrReportData.strDetails(lconCCCDetsAuthCode, llngRecCount - 1).strValue = Trim$(.Fields("AuthorisationCode") & "")
            gstrReportData.strDetails(lconCCCDetsOrderNum, llngRecCount - 1).strValue = Trim$(.Fields("OrderNum"))
            gstrReportData.strDetails(lconCCCDetsDespDate, llngRecCount - 1).strValue = Format$(.Fields("DespatchDate") & "", "DD/MM/YY")
            gstrReportData.strDetails(lconCCCDetsCustomerName, llngRecCount - 1).strValue = Trim$(Trim$(.Fields("CallerSalutation") & "") & " " & Trim$(.Fields("CallerInitials") & "") & " " & Trim$(.Fields("CallerSurname") & ""))
            
            If Trim$(.Fields("AdviceAdd1")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerLine1, llngRecCount - 1).strValue = Trim$(.Fields("AdviceAdd1")) & ", "
            End If
            If Trim$(.Fields("AdviceAdd2")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerLine2, llngRecCount - 1).strValue = Trim$(.Fields("AdviceAdd2")) & ", "
            End If
            If Trim$(.Fields("AdviceAdd3")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerLine3, llngRecCount - 1).strValue = Trim$(.Fields("AdviceAdd3")) & ", "
            End If
            If Trim$(.Fields("AdviceAdd4")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerLine4, llngRecCount - 1).strValue = Trim$(.Fields("AdviceAdd4")) & ", "
            End If
            If Trim$(.Fields("AdviceAdd5")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerLine5, llngRecCount - 1).strValue = Trim$(.Fields("AdviceAdd5")) & ", "
            End If
            If Trim$(.Fields("AdvicePostcode")) <> "" Then
                gstrReportData.strDetails(lconCCCDetsCustomerPostCode, llngRecCount - 1).strValue = Trim$(.Fields("AdvicePostcode"))
            End If
            
            lintLineNum = lintLineNum + 1
            lcurPageTotal = lcurPageTotal + CCur(lcurTotalCCPayment)
            
           
            lcurTotal = lcurTotal + CCur(lcurTotalCCPayment)
             
            If lintLineNum = ((lintPageNum * gintNumberOfLinesAPage) - 8) Then
                gstrReportData.strFooters(lconCCCPageTotal).strValue = Format$(lcurPageTotal, "0.00")
                lcurPageTotal = 0
                lintLineNum = lintLineNum + 11
                gstrReportData.strFooters(lconCCCPageNum).strValue = lintPageNum
                gstrReportData.strFooters(lconCCCGrandTotal).strValue = Format$(lcurTotal, "0.00"):
                
                'Put file
                Put #lintFreeFile, lintPageNum, gstrReportData
                ClearReportingDataType "D"
                lintPageNum = lintPageNum + 1
                'lintPageNum = lintPageNum + 1
                
            End If
                                
           
            'lcurTotal = lcurTotal + CCur(lcurTotalCCPayment)
            
            If llngRecCount = 1 Then
                glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = .Fields("OrderNum")
            Else
                ReDim Preserve glngTrackNUpdate(UBound(glngTrackNUpdate) + 1)
                glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = .Fields("OrderNum")
            End If
            
            .MoveNext
        Loop
                
        If lcurPageTotal > 0 Then
            lintLineNum = lintLineNum + 4
            gstrReportData.strFooters(lconCCCPageTotal).strValue = Format$(lcurPageTotal, "0.00")
        End If
        
        lintLineNum = lintLineNum + 3
        
       
        'lcurTotal = lcurTotal + CCur(lcurTotalCCPayment)
        
        gstrReportData.strFooters(lconCCCGrandTotal).strValue = Format$(lcurTotal, "0.00"):
                
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        'Put file
        gstrReportData.strFooters(lconCCCPageNum).strValue = lintPageNum
        Put #lintFreeFile, lintPageNum, gstrReportData
        'lintPageNum = lintPageNum + 1
        ClearReportingDataType "D"
        
        'Close lintFileNum
        Close #lintFreeFile
        gintCurrentReportPageNum = lintPageNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintObjCreditCardClaim", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select

End Function
