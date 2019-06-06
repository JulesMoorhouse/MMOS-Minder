Attribute VB_Name = "modOutRep"
Option Explicit
Sub AddparcelNumberArray(plngOrderNumber As Long, pstrCatNum As String, plngParcelNumber As Long)

    ReDim Preserve gstrOrderLineParcelNumbers(UBound(gstrOrderLineParcelNumbers) + 1)
                
    With gstrOrderLineParcelNumbers(UBound(gstrOrderLineParcelNumbers))
        .lngOrderNum = plngOrderNumber
        .strCatNum = pstrCatNum
        .lngParcelBoxNumber = plngParcelNumber
    End With

End Sub

Sub PrintBankingReport(pstrFilename As String, pdatDate As Date)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrName As String
Dim lcurTotal As Currency
Dim lcurVoucher As Currency
Dim lcurCheque As Currency
Dim lcurCash As Currency
Dim lcurCreditCard As Currency

Dim lcurTotalVoucher As Currency
Dim lcurTotalCheque As Currency
Dim lcurTotalCash As Currency
Dim lcurTotalCreditCard As Currency
Dim lcurTotalBank As Currency
Dim lintArrInc2 As Integer
Dim lintLineNum As Integer
Dim lintPageNum As Integer

    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT CreationDate, CallerSalutation, CallerInitials, CallerSurname, " & _
        "CustNum, TotalIncVat, OrderType, Payment, PaymentType2, Payment2 " & _
        "From " & gtblAdviceNotes & " WHERE CreationDate = #" & pdatDate & "# And OrderStatus='D' " & _
        "Or OrderStatus='E' Or OrderStatus='C' " & _
        "Or OrderStatus='B' ORDER BY CreationDate, CallerSurname;"

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lcurTotal = 0
        lcurTotalVoucher = 0
        lcurTotalCheque = 0
        lcurTotalCash = 0
        lcurTotalCreditCard = 0
        lcurTotalBank = 0
        lintPageNum = 1
        
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        Print #lintFileNum, "BANKING REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
        Print #lintFileNum, " ": lintLineNum = lintLineNum + 1

        Print #lintFileNum, Spacer("Date", 11) & _
             Spacer("Customer Name", 27) & Spacer("Cust No.", 9) & _
             Spacer("Total Value", 11, "L") & _
             Spacer("Voucher", 11, "L") & _
             Spacer("Cheque", 11, "L") & _
             Spacer("Cash", 11, "L") & _
             Spacer("Total Bank", 11, "L") & _
             Spacer("C/Card", 11, "L"): lintLineNum = lintLineNum + 1
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            lcurVoucher = 0
            lcurCheque = 0
            lcurCash = 0
            lcurCreditCard = 0
                        
            lstrName = Trim$(Trim$(.Fields("CallerSalutation")) & " " & _
                Trim$(.Fields("CallerInitials")) & " " & Trim$(.Fields("CallerSurname")))
                                    
            Select Case Trim$(.Fields("OrderType"))
            Case "C": lcurCreditCard = CCur(.Fields("Payment"))
            Case "Q": lcurCheque = CCur(.Fields("Payment"))
            Case "V": lcurVoucher = CCur(.Fields("Payment"))
            End Select
            
            Select Case Trim$(.Fields("PaymentType2"))
            Case "C": lcurCreditCard = CCur(.Fields("Payment2"))
            Case "Q": lcurCheque = CCur(.Fields("Payment2"))
            Case "V": lcurVoucher = CCur(.Fields("Payment2"))
            End Select
            
            Print #lintFileNum, Spacer(Format$(.Fields("CreationDate"), "DD/MM/YY"), 11) & _
                 Spacer(lstrName, 27) & Spacer(.Fields("CustNum"), 9) & _
                 Spacer(Format$(.Fields("TotalIncVat"), "0.00"), 11, "L") & _
                 Spacer(Format$(lcurVoucher, "0.00"), 11, "L") & _
                 Spacer(Format$(lcurCheque, "0.00"), 11, "L") & _
                 Spacer(Format$(lcurCash, "0.00"), 11, "L") & _
                 Spacer(Format$((lcurVoucher + lcurCheque + lcurCash), "0.00"), 11, "L") & _
                 Spacer(Format$(lcurCreditCard, "0.00"), 11, "L"): lintLineNum = lintLineNum + 1
                                
            lcurTotalBank = lcurTotalBank + CCur(lcurVoucher + lcurCheque + lcurCash)
            lcurTotal = lcurTotal + CCur(.Fields("TotalIncvat"))
            
            lcurTotalVoucher = lcurTotalVoucher + CCur(lcurVoucher)
            lcurTotalCheque = lcurTotalCheque + CCur(lcurCheque)
            lcurTotalCash = lcurTotalCash + CCur(lcurCash)
            lcurTotalCreditCard = lcurTotalCreditCard + CCur(lcurCreditCard)
            
            If lintLineNum = lintPageNum * (gintNumberOfLinesAPage - 2) Then
                Print #lintFileNum, Spacer("", 70) & "Page No. " & lintPageNum: lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                lintPageNum = lintPageNum + 1
                Print #lintFileNum, "BANKING REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
                Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
    
                Print #lintFileNum, Spacer("Date", 11) & _
                     Spacer("Customer Name", 27) & Spacer("Cust No.", 9) & _
                     Spacer("Total Value", 11, "L") & _
                     Spacer("Voucher", 11, "L") & _
                     Spacer("Cheque", 11, "L") & _
                     Spacer("Cash", 11, "L") & _
                     Spacer("Total Bank", 11, "L") & _
                     Spacer("C/Card", 11, "L"): lintLineNum = lintLineNum + 1
            End If
            .MoveNext
        Loop
        
        Print #lintFileNum, Spacer("", 11) & _
             Spacer("", 27) & Spacer("", 9) & _
             Spacer("==========", 11, "L") & _
             Spacer("==========", 11, "L") & _
             Spacer("==========", 11, "L") & _
             Spacer("==========", 11, "L") & _
             Spacer("==========", 11, "L") & _
             Spacer("==========", 11, "L"): lintLineNum = lintLineNum + 1
                
        Print #lintFileNum, Spacer("", 11) & _
             Spacer("", 27) & Spacer("", 9) & _
             Spacer(Format$(lcurTotal, "0.00"), 11, "L") & _
             Spacer(Format$(lcurTotalVoucher, "0.00"), 11, "L") & _
             Spacer(Format$(lcurTotalCheque, "0.00"), 11, "L") & _
             Spacer(Format$(lcurTotalCash, "0.00"), 11, "L") & _
             Spacer(Format$(lcurTotalBank, "0.00"), 11, "L") & _
             Spacer(Format$(lcurTotalCreditCard, "0.00"), 11, "L"): lintLineNum = lintLineNum + 1
                
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            Print #lintFileNum, ""
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintBankingReport", "Central")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select
End Sub


Function PrintToFileAdviceInvoiceFooter(pstrUnitTotal As String, pstrZeroUnitTotal As String, pstrVatTotal As String, pstrAdviceType As String) As String
Dim lstrAdviceNote As String
Dim lstrConsignmentNoteLine1 As String
Dim lstrConsignmentNoteLine2 As String
Dim lstrTaxSummaryLine1 As String
Dim lstrTaxSummaryLine2 As String
Dim lstrTaxSummaryLine3 As String
Dim lstrPaymentType2 As String

    With gstrAdviceNoteOrder
        lstrAdviceNote = lstrAdviceNote & "" & vbCrLf '1
        lstrAdviceNote = lstrAdviceNote & Space(57) & Line(21) & vbCrLf '2
        If CCur(AdvicePrice(.strDonation)) > 0 Then
            lstrAdviceNote = lstrAdviceNote & Spacer("Payments received. ***  THANK YOU FOR YOUR DONATION  ***", 57) & "Goods & VAT " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(pstrZeroUnitTotal)) + (PriceVal(pstrVatTotal))), 10, "L") & vbCrLf '3
        Else
            lstrAdviceNote = lstrAdviceNote & Spacer("Payments received.", 57) & "Goods & VAT " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(pstrZeroUnitTotal)) + (PriceVal(pstrVatTotal))), 10, "L") & vbCrLf '3
        End If
        If .strPayment <> 0 Then
            lstrAdviceNote = lstrAdviceNote & Spacer(GetListCodeDesc("Payment Method", .strPaymentType1) & " " & AdvicePrice(.strPayment), 57) & "P&P         " & Spacer(AdvicePrice(.strPostage), 10, "L") & vbCrLf '4
        Else
            lstrAdviceNote = lstrAdviceNote & Spacer("", 57) & "P&P         " & Spacer(AdvicePrice(.strPostage), 10, "L") & vbCrLf '4
        End If
        lstrPaymentType2 = GetListCodeDesc("Payment Method", .strPaymentType2) & " " & AdvicePrice(.strPayment2) & vbCrLf '5
        lstrTaxSummaryLine1 = "Tax Summary    Goods+P&P    Vat"
        lstrTaxSummaryLine2 = "Standard " & gstrReferenceInfo.strVATRate175 & "% " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal)), 9, "L") & Spacer(AdvicePrice(PriceVal(pstrVatTotal)), 7, "L")
        lstrTaxSummaryLine3 = "Zero     0%    " & Spacer(AdvicePrice(PriceVal(pstrZeroUnitTotal)), 9, "L") & Spacer(AdvicePrice(PriceVal("0")), 7, "L")
        If .strPayment2 <> 0 Then
            lstrAdviceNote = lstrAdviceNote & Spacer(Spacer(lstrPaymentType2, 25) & lstrTaxSummaryLine1, 57) & vbCrLf '6
        Else
            lstrAdviceNote = lstrAdviceNote & Space(25) & Spacer(lstrTaxSummaryLine1, 32) & vbCrLf '6
        End If
        lstrAdviceNote = lstrAdviceNote & Space(25) & Spacer(lstrTaxSummaryLine2, 32) & "Donation    " & Spacer(AdvicePrice(.strDonation), 10, "L") & vbCrLf '7
        lstrAdviceNote = lstrAdviceNote & Space(25) & Spacer(lstrTaxSummaryLine3, 32) & vbCrLf '8
        If Trim$(pstrAdviceType) = "REFUND" Then
            lstrAdviceNote = lstrAdviceNote & Spacer("Please find enclosed a refund cheque", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L") & vbCrLf '9
            lstrAdviceNote = lstrAdviceNote & Spacer("for the reason shown on this note.", 57) & Line(21) & vbCrLf '10
        Else
            If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 And _
            PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) <= 1 Then
                lstrAdviceNote = lstrAdviceNote & Spacer("Please claim refund with next order", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L") & vbCrLf '9
                lstrAdviceNote = lstrAdviceNote & Spacer("quoting order no. and refund value.", 57) & Line(21) & vbCrLf '10
            ElseIf PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 1 Then
                lstrAdviceNote = lstrAdviceNote & Spacer("A cheque will follow for the sum and", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L") & vbCrLf '9
                lstrAdviceNote = lstrAdviceNote & Spacer("reason stated, under separate cover.", 57) & Line(21) & vbCrLf '10
            Else
                lstrAdviceNote = lstrAdviceNote & Spacer("", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L") & vbCrLf '9
                lstrAdviceNote = lstrAdviceNote & Space(57) & Line(21) & vbCrLf '10
            End If
        End If

        If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 And Trim$(gstrConsignmentNote.strText) <> "" Then
            lstrConsignmentNoteLine1 = GrapLine("Note: " & Trim$(gstrConsignmentNote.strText), 0, 55)
            lstrConsignmentNoteLine2 = GrapLine("Note: " & Trim$(gstrConsignmentNote.strText), 1, 55)
        Else
            lstrConsignmentNoteLine1 = ""
            lstrConsignmentNoteLine2 = ""
        End If
        lstrAdviceNote = lstrAdviceNote & Spacer(" ", 57) & "OverPayment " & Spacer(AdvicePrice(.strReconcilliation), 10, "L") & vbCrLf '11
        lstrAdviceNote = lstrAdviceNote & Spacer(lstrConsignmentNoteLine1, 57) & "UnderPaym't " & Spacer(AdvicePrice(.strUnderpayment), 10, "L") & vbCrLf '12
        If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 Then
            lstrAdviceNote = lstrAdviceNote & Spacer(lstrConsignmentNoteLine2, 57) & "Total Refund" & Spacer(AdvicePrice(PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment)), 10, "L") & vbCrLf '13
        Else
            lstrAdviceNote = lstrAdviceNote & Spacer(lstrConsignmentNoteLine2, 57) & "Total Refund" & Spacer(AdvicePrice("0.00"), 10, "L") & vbCrLf '13
        End If
    End With
    
    PrintToFileAdviceInvoiceFooter = lstrAdviceNote
    
End Function
Function PrintToFileAdviceInvoiceHeader(pstrAdviceType As String, ByRef lstrCourier As String, _
    ByRef lstrPFServiceInd As String, ByRef pstrProductLines() As OrderDetail, lintArrInc2 As Integer) As String
Dim lstrAdviceNote As String

    With gstrAdviceNoteOrder
        If Trim$(pstrAdviceType) = "REFUND" Then
            lstrAdviceNote = "REFUND ADVICE NOTE" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf '5
        Else
            lstrAdviceNote = "ADVICE NOTE" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf '5
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
    
    PrintToFileAdviceInvoiceHeader = lstrAdviceNote
    
End Function


Function PrintToFileAdviceInvoiceMiddle(pstrCourier As String, plngParcelNumber As Long, _
    plngTotalParcels As Long, pintPageNumber As Integer, pintTotalpages As Integer) As String
    
Dim lstrAdviceNote As String
Dim lstrName As String
Dim lstrDeliveryName As String

    With gstrAdviceNoteOrder
        lstrAdviceNote = lstrAdviceNote & Spacer("", 42) & "Deliver to:      " & pstrCourier & vbCrLf '1
        lstrName = Trim$(Trim$(.strSalutation) & " " & Trim$(.strInitials) & " " & Trim$(.strSurname))
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & Trim$(.strDeliveryInitials) & " " & Trim$(.strDeliverySurname))
        If Trim$(lstrDeliveryName) = "" Then lstrDeliveryName = lstrName
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(lstrName, 38) & lstrDeliveryName & vbCrLf '2
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strAdd1, 38) & .strDeliveryAdd1 & vbCrLf '3
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strAdd2, 38) & .strDeliveryAdd2 & vbCrLf '4
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strAdd3, 38) & .strDeliveryAdd3 & vbCrLf '5
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strAdd4, 38) & .strDeliveryAdd4 & vbCrLf '6
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strAdd5, 38) & .strDeliveryAdd5 & vbCrLf '7
        lstrAdviceNote = lstrAdviceNote & Space(8) & Spacer(.strPostcode, 38) & .strDeliveryPostcode & vbCrLf ' 8
        lstrAdviceNote = lstrAdviceNote & "" & vbCrLf '9
        lstrAdviceNote = lstrAdviceNote & vbTab & vbTab & Trim$(.strMediaCode) & " " & GetListCodeDesc("Media Codes", .strMediaCode) & vbCrLf '10
        lstrAdviceNote = lstrAdviceNote & vbTab & vbTab & vbTab & "Page " & pintPageNumber & "/" & pintTotalpages & vbCrLf '10
        lstrAdviceNote = lstrAdviceNote & "Customer" & vbTab & "Order" & vbTab & vbTab & "  Order" & vbTab & vbTab & "  Ship" & vbTab & vbTab & " Parcel" & vbCrLf '12
        lstrAdviceNote = lstrAdviceNote & "   No." & vbTab & vbTab & " No." & vbTab & vbTab & "  Date" & vbTab & vbTab & "  Date" & vbTab & vbTab & "  No." & vbCrLf '13
        If .datDeliveryDate = "00:00:00" Then
            lstrAdviceNote = lstrAdviceNote & Spacer("M" & .lngCustNum, 8, "L") & vbTab & Spacer("" & .lngOrderNum, 8, "L") & vbTab & _
                Format(.datCreationDate, "DD/MM/YYYY") & vbTab & "Not Specified" & vbTab & "  " & plngParcelNumber & "/" & plngTotalParcels & vbCrLf  '14
        Else
            lstrAdviceNote = lstrAdviceNote & Spacer("M" & .lngCustNum, 8, "L") & vbTab & Spacer("" & .lngOrderNum, 8, "L") & vbTab & _
                Format(.datCreationDate, "DD/MM/YYYY") & vbTab & Format$(.datDeliveryDate, "DD/MM/YYYY") & vbTab & "  " & plngParcelNumber & "/" & plngTotalParcels & vbCrLf '14
        End If
        lstrAdviceNote = lstrAdviceNote & Line(78) & vbCrLf '15
        lstrAdviceNote = lstrAdviceNote & "Cat             Qty   Qty" & Space(34) & "Unit   Vat    Amount" & vbCrLf '16
        lstrAdviceNote = lstrAdviceNote & "Code     Bin    Ord   Desp" & Space(33) & "Price  Code" & vbCrLf '17
        lstrAdviceNote = lstrAdviceNote & Line(78) & vbCrLf '18
    End With

    PrintToFileAdviceInvoiceMiddle = lstrAdviceNote
    
End Function

Sub PrintToFileAdviceInvoiceSAFE(pstrFilename As String, pstrMode As String, Optional pstrAdviceType As String)
Dim lintFileNum As Integer
Dim lstrName As String
Dim lstrDeliveryName As String
Dim pstrProductLines() As OrderDetail
Dim llngNumOfProductLines As Long
Dim pstrVatTotal As String
Dim pstrUnitTotal As String
Dim pstrZeroUnitTotal As String
Dim lintLineNum As Integer
Dim lintArrInc As Integer
Dim lintArrInc2 As Integer
Dim lstrCourier As String
Dim lstrInternalNote As String
Dim lstrConsignmentNote As String
Dim lstrTaxSummaryLine1 As String
Dim lstrTaxSummaryLine2 As String
Dim lstrTaxSummaryLine3 As String
Dim lstrPaymentType2 As String

    lintLineNum = 0
    lintFileNum = FreeFile
    
    If IsMissing(pstrAdviceType) Then
        pstrAdviceType = ""
    End If
    
    gstrAdviceServiceInd.strListName = "PForce Service Indicator"

    If pstrProductLines(0).booLightParcel = False Then
        UpdateAdviceParcelExtras gstrAdviceNoteOrder.lngCustNum, gstrAdviceNoteOrder.lngOrderNum, UBound(pstrProductLines) + 1, glngGrossWeight
    End If
    
    For lintArrInc2 = 0 To UBound(pstrProductLines)
    
    If pstrProductLines(lintArrInc2).lngNumberOfLines = "" Then
        If pstrProductLines(lintArrInc2).strProductLines = "" Then
            If UBound(pstrProductLines) = lintArrInc2 Then
                'New parcel but no lines
                'Shouldn't happen
                
            End If
        End If
    End If
    
    With gstrAdviceNoteOrder
        lintLineNum = 0
        
        Open pstrFilename For Append As lintFileNum
        If Trim$(pstrAdviceType) = "REFUND" Then
            Print #lintFileNum, "REFUND ADVICE NOTE": lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, "ADVICE NOTE": lintLineNum = lintLineNum + 1
        End If
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        
        Dim lstrPFServiceInd As String
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
        
        Print #lintFileNum, Spacer("", 42) & "Deliver to:      " & lstrCourier: lintLineNum = lintLineNum + 1
                    
        lstrName = Trim$(Trim$(.strSalutation) & " " & _
            Trim$(.strInitials) & " " & Trim$(.strSurname))
        
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & _
            Trim$(.strDeliveryInitials) & " " & Trim$(.strDeliverySurname))
            
        If Trim$(lstrDeliveryName) = "" Then
            lstrDeliveryName = lstrName
        End If
        
        Print #lintFileNum, Space(8) & Spacer(lstrName, 38) & lstrDeliveryName: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strAdd1, 38) & .strDeliveryAdd1: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strAdd2, 38) & .strDeliveryAdd2: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strAdd3, 38) & .strDeliveryAdd3: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strAdd4, 38) & .strDeliveryAdd4: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strAdd5, 38) & .strDeliveryAdd5: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(8) & Spacer(.strPostcode, 38) & .strDeliveryPostcode: lintLineNum = lintLineNum + 1
    
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Print #lintFileNum, vbTab & vbTab & Trim$(.strMediaCode) & " " & GetListCodeDesc("Media Codes", .strMediaCode): lintLineNum = lintLineNum + 1
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1

        Print #lintFileNum, "Customer" & vbTab & "Order" & vbTab & vbTab & "  Order" & vbTab & vbTab & "  Ship" & vbTab & vbTab & " Parcel": lintLineNum = lintLineNum + 1
        Print #lintFileNum, "   No." & vbTab & vbTab & " No." & vbTab & vbTab & "  Date" & vbTab & vbTab & "  Date" & vbTab & vbTab & "  No.": lintLineNum = lintLineNum + 1
        
        If .datDeliveryDate = "00:00:00" Then
        
            Print #lintFileNum, Spacer("M" & .lngCustNum, 8, "L") & vbTab & Spacer("" & .lngOrderNum, 8, "L") & vbTab & _
                Format(.datCreationDate, "DD/MM/YYYY") & vbTab & "Not Specified" & vbTab & (lintArrInc2 + 1) & "/" & UBound(pstrProductLines) + 1: lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, Spacer("M" & .lngCustNum, 8, "L") & vbTab & Spacer("" & .lngOrderNum, 8, "L") & vbTab & _
                Format(.datCreationDate, "DD/MM/YYYY") & vbTab & Format$(.datDeliveryDate, "DD/MM/YYYY") & vbTab & (lintArrInc2 + 1): lintLineNum = lintLineNum + 1
        End If
            
        Print #lintFileNum, Line(78): lintLineNum = lintLineNum + 1
        Print #lintFileNum, "Cat             Qty   Qty" & Space(34) & "Unit   Vat    Amount": lintLineNum = lintLineNum + 1
        Print #lintFileNum, "Code     Bin    Ord   Desp" & Space(33) & "Price  Code": lintLineNum = lintLineNum + 1
        Print #lintFileNum, Line(78): lintLineNum = lintLineNum + 1
        
        'Products
        Print #lintFileNum, pstrProductLines(lintArrInc2).strProductLines
        
        Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(57) & Line(21): lintLineNum = lintLineNum + 1
        
        If CCur(AdvicePrice(.strDonation)) > 0 Then
            Print #lintFileNum, Spacer("Payments received. ***  THANK YOU FOR YOUR DONATION  ***", 57) & "Goods & VAT " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(pstrZeroUnitTotal)) + (PriceVal(pstrVatTotal))), 10, "L"): lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, Spacer("Payments received.", 57) & "Goods & VAT " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(pstrZeroUnitTotal)) + (PriceVal(pstrVatTotal))), 10, "L"): lintLineNum = lintLineNum + 1
        End If
        
        If .strPayment <> 0 Then
            Print #lintFileNum, Spacer(GetListCodeDesc("Payment Method", .strPaymentType1) & " " & AdvicePrice(.strPayment), 57) & "P&P         " & Spacer(AdvicePrice(.strPostage), 10, "L"): lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, Spacer("", 57) & "P&P         " & Spacer(AdvicePrice(.strPostage), 10, "L"): lintLineNum = lintLineNum + 1
        End If
        
        lstrPaymentType2 = GetListCodeDesc("Payment Method", .strPaymentType2) & " " & AdvicePrice(.strPayment2)
        
        lstrTaxSummaryLine1 = "Tax Summary    Goods+P&P    Vat"
        lstrTaxSummaryLine2 = "Standard " & gstrReferenceInfo.strVATRate175 & "% " & Spacer(AdvicePrice(PriceVal(pstrUnitTotal)), 9, "L") & Spacer(AdvicePrice(PriceVal(pstrVatTotal)), 7, "L")
        lstrTaxSummaryLine3 = "Zero     0%    " & Spacer(AdvicePrice(PriceVal(pstrZeroUnitTotal)), 9, "L") & Spacer(AdvicePrice(PriceVal("0")), 7, "L")
        
        If .strPayment2 <> 0 Then
            Print #lintFileNum, Spacer(Spacer(lstrPaymentType2, 25) & lstrTaxSummaryLine1, 57): lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, Space(25) & Spacer(lstrTaxSummaryLine1, 32): lintLineNum = lintLineNum + 1
        End If
        
        
        Print #lintFileNum, Space(25) & Spacer(lstrTaxSummaryLine2, 32) & "Donation    " & Spacer(AdvicePrice(.strDonation), 10, "L"): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Space(25) & Spacer(lstrTaxSummaryLine3, 32): lintLineNum = lintLineNum + 1
        
        If Trim$(pstrAdviceType) = "REFUND" Then
            Print #lintFileNum, Spacer("Please find enclosed a refund cheque", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L"): lintLineNum = lintLineNum + 1
            Print #lintFileNum, Spacer("for the reason shown on this note.", 57) & Line(21): lintLineNum = lintLineNum + 1
        Else

            If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 And _
            PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) <= 1 Then
                Print #lintFileNum, Spacer("Please claim refund with next order", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L"): lintLineNum = lintLineNum + 1
                Print #lintFileNum, Spacer("quoting order no. and refund value.", 57) & Line(21): lintLineNum = lintLineNum + 1
            ElseIf PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 1 Then
                Print #lintFileNum, Spacer("A cheque will follow for the sum and", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L"): lintLineNum = lintLineNum + 1
                Print #lintFileNum, Spacer("reason stated, under separate cover.", 57) & Line(21): lintLineNum = lintLineNum + 1
            Else
                Print #lintFileNum, Spacer("", 57) & "Total       " & Spacer(AdvicePrice(.strTotalIncVat), 10, "L"): lintLineNum = lintLineNum + 1
                Print #lintFileNum, Space(57) & Line(21): lintLineNum = lintLineNum + 1
            End If
        
        End If
                
        If gstrAdviceNoteOrder.lngAdviceRemarkNum <> 0 And Trim$(gstrInternalNote.strText) <> "" Then
            lstrInternalNote = "Note: " & Trim$(gstrInternalNote.strText)
        Else
            lstrInternalNote = ""
        End If
        If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 And Trim$(gstrConsignmentNote.strText) <> "" Then
            lstrConsignmentNote = "Note: " & Trim$(gstrConsignmentNote.strText)
        Else
            lstrConsignmentNote = ""
        End If
        
        Print #lintFileNum, Spacer(lstrInternalNote, 57) & "OverPayment " & Spacer(AdvicePrice(.strReconcilliation), 10, "L"): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer(lstrConsignmentNote, 57) & "UnderPaym't " & Spacer(AdvicePrice(.strUnderpayment), 10, "L"): lintLineNum = lintLineNum + 1
        
        If PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment) > 0 Then
            Print #lintFileNum, Spacer("", 57) & "Total Refund" & Spacer(AdvicePrice(PriceVal(.strReconcilliation) - PriceVal(.strUnderpayment)), 10, "L"): lintLineNum = lintLineNum + 1
        Else
            Print #lintFileNum, Spacer("", 57) & "Total Refund" & Spacer(AdvicePrice("0.00"), 10, "L"): lintLineNum = lintLineNum + 1
        End If
        
        For lintArrInc = 1 To (gintNumberOfLinesAPage - lintLineNum) - pstrProductLines(lintArrInc2).lngNumberOfLines
            Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        Next lintArrInc
        Debug.Print .lngOrderNum & " " & lintLineNum + pstrProductLines(lintArrInc2).lngNumberOfLines
        lintLineNum = 0
    End With
    Close lintFileNum
    
    Next lintArrInc2
        

    'Shell "notepad " & lstrFileName

End Sub
Sub PrintToFileAdviceInvoice(pstrFilename As String, pstrMode As String, Optional pstrAdviceType As String)
Dim lintFileNum As Integer
Dim pstrProductLines() As OrderDetail
Dim llngNumOfProductLines As Long
Dim pstrVatTotal As String: Dim pstrUnitTotal As String: Dim pstrZeroUnitTotal As String
Dim lintArrInc As Integer: Dim lintArrInc2 As Integer: Dim lstrCourier As String
Dim lstrInternalNote As String: Dim lstrAdviceNote As String
Dim lstrPFServiceInd As String
Dim llngMaxproductLines As Long
Const lconintNumberOflineinHeader = 5: Const lconintNumberOflineinMiddle = 18: Const lconintNumberOflineinFooter = 13
    
    llngMaxproductLines = lconintNumberOflineinHeader + lconintNumberOflineinMiddle + lconintNumberOflineinFooter
    llngMaxproductLines = gintNumberOfLinesAPage - llngMaxproductLines
    
    If IsMissing(pstrAdviceType) Then pstrAdviceType = ""
    
    gstrAdviceServiceInd.strListName = "PForce Service Indicator"
    
    PrintGetProductsOrderlines gstrAdviceNoteOrder.lngCustNum, pstrProductLines(), pstrUnitTotal, pstrVatTotal, gstrAdviceNoteOrder.lngOrderNum, pstrMode, pstrZeroUnitTotal, llngMaxproductLines
    
    If pstrProductLines(0).booLightParcel = False Then
        UpdateAdviceParcelExtras gstrAdviceNoteOrder.lngCustNum, _
            gstrAdviceNoteOrder.lngOrderNum, CInt(glngTotalParcel), glngGrossWeight
    End If
    
    For lintArrInc2 = 0 To UBound(pstrProductLines)
        If pstrProductLines(lintArrInc2).lngNumberOfLines = "" Then
            If pstrProductLines(lintArrInc2).strProductLines = "" Then
                If UBound(pstrProductLines) = lintArrInc2 Then
                    'New parcel but no lines  'Shouldn't happen
                End If
            End If
        End If
            
            lstrAdviceNote = PrintToFileAdviceInvoiceHeader(pstrAdviceType, lstrCourier, lstrPFServiceInd, _
                pstrProductLines(), lintArrInc2)
            
            lstrAdviceNote = lstrAdviceNote & PrintToFileAdviceInvoiceMiddle(lstrCourier, _
                pstrProductLines(lintArrInc2).lngParcelBoxNumber, glngTotalParcel, _
                (lintArrInc2 + 1), UBound(pstrProductLines) + 1)

            lstrAdviceNote = lstrAdviceNote & pstrProductLines(lintArrInc2).strProductLines & vbCrLf
            
            lstrAdviceNote = lstrAdviceNote & PrintToFileAdviceInvoiceFooter(pstrUnitTotal, pstrZeroUnitTotal, pstrVatTotal, pstrAdviceType)
            
            For lintArrInc = 1 To (gintNumberOfLinesAPage - (lconintNumberOflineinHeader + lconintNumberOflineinMiddle + lconintNumberOflineinFooter)) - pstrProductLines(lintArrInc2).lngNumberOfLines
                lstrAdviceNote = lstrAdviceNote & "" & vbCrLf
            Next lintArrInc
    
            lintFileNum = FreeFile
            Open pstrFilename For Append As lintFileNum
            Print #lintFileNum, lstrAdviceNote
            Close lintFileNum
    Next lintArrInc2


End Sub
Sub PrintGetProductsOrderlinesSAFEOld(plngCustomerNum As Long, ByRef pstrProductLines() As OrderDetail, pstrUnitTotal As String, pstrVatTotal As String, plngOrderNumber As Long, pstrMode As String, pstrZeroUnitTotal As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrStepError As String
Dim llngTotalWeight As Long
Dim lintParcelNumber As Integer
Dim lbooOutOfStockCaption As Boolean

    On Error GoTo ErrHandler
    If pstrMode = "LOCAL" Then
        lstrSQL = "SELECT * From " & gtblOrderLines & " Where " & gtblOrderLines & ".CustNum = " & _
        plngCustomerNum & " ORDER BY " & gtblOrderLines & ".BinLocation;"
    ElseIf pstrMode = "CENTRAL" Then
        lstrSQL = "SELECT * From " & gtblMasterOrderLines & " "
        lstrSQL = lstrSQL & "Where CustNum = " & _
        plngCustomerNum & " and OrderNum = " & plngOrderNumber & " ORDER BY BinLocation;"
    End If
    lintParcelNumber = 0
    
    ReDim pstrProductLines(lintParcelNumber)
    pstrProductLines(lintParcelNumber).booLightParcel = False
    pstrProductLines(lintParcelNumber).lngNumberOfLines = 0
    
    lstrStepError = "Step One"
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        glngGrossWeight = 0
        llngRecCount = 0
        lstrStepError = "Step Two"
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
                If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 18000 Then
                    If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 800 Then
                        pstrProductLines(lintParcelNumber).booLightParcel = True
                    Else
                        pstrProductLines(lintParcelNumber).booLightParcel = False
                    End If
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(lintParcelNumber).lngNumberOfLines = llngRecCount
                ElseIf (llngTotalWeight + Val(.Fields("TotalWeight"))) > 18000 And llngTotalWeight = 0 Then
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(lintParcelNumber).lngNumberOfLines = llngRecCount
                    pstrProductLines(lintParcelNumber).booLightParcel = False
                Else
                    'New Parcel
                    pstrProductLines(lintParcelNumber).booLightParcel = False
                    llngTotalWeight = 0
                    llngRecCount = 1
                    lintParcelNumber = lintParcelNumber + 1
                    ReDim Preserve pstrProductLines(lintParcelNumber)
                    pstrProductLines(lintParcelNumber).booLightParcel = False
                    pstrProductLines(lintParcelNumber).lngNumberOfLines = 0
                End If
            
                If Trim$(pstrProductLines(lintParcelNumber).strProductLines) <> "" Then
                    pstrProductLines(lintParcelNumber).strProductLines = pstrProductLines(lintParcelNumber).strProductLines & vbCrLf
                    lstrStepError = "Step three"
                End If
                
                If Val(.Fields("Qty")) = 0 Or Val(.Fields("DespQty")) = 0 Then
                    lbooOutOfStockCaption = True 'show out of stock caption
                Else
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                If Trim$(.Fields("CatNum")) = "REFUND" Then
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                If lbooOutOfStockCaption = True Then
                    pstrProductLines(lintParcelNumber).strProductLines = pstrProductLines(lintParcelNumber).strProductLines & Spacer(.Fields("CatNum"), 6) & _
                        Spacer(.Fields("BinLocation"), 9) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer("***  Out of Stock  ***", 22, "L")
                Else
                    pstrProductLines(lintParcelNumber).strProductLines = pstrProductLines(lintParcelNumber).strProductLines & Spacer(.Fields("CatNum"), 8) & _
                        Spacer(.Fields("BinLocation"), 7) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer(AdvicePrice(.Fields("Price")), 8, "L") & " " & Spacer(.Fields("TaxCode"), 4, "L") & _
                        Spacer(AdvicePrice(.Fields("TotalPrice")), 10, "L")
                End If
                                        
                lstrStepError = "Step Four"
                If .Fields("TaxCode") = "Z" Then
                    '0
                    pstrZeroUnitTotal = AdvicePrice(PriceVal(pstrZeroUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                Else
                    lstrStepError = "Step Five"
                    pstrVatTotal = AdvicePrice(PriceVal(pstrVatTotal) + (Val(.Fields("DespQty")) * PriceVal(.Fields("Price")) * Val(gstrVATRate) / 100))
                    pstrUnitTotal = AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                End If
                            
                lstrStepError = "Step Six"
                pstrProductLines(lintParcelNumber).lngNumberOfLines = llngRecCount
                glngGrossWeight = glngGrossWeight + llngTotalWeight
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    lstrStepError = "Step Seven"
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
        
    Select Case GlobalErrorHandler(Err.Number, "PrintGetProductsOrderLines " & lstrStepError, "Local")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select
  
End Sub
Sub PrintGetProductsOrderlines(plngCustomerNum As Long, ByRef pstrProductLines() As OrderDetail, _
    pstrUnitTotal As String, pstrVatTotal As String, plngOrderNumber As Long, _
    pstrMode As String, pstrZeroUnitTotal As String, plngMaxproductLines As Long)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrStepError As String
Dim llngTotalWeight As Long
Dim lintParcelNumber As Integer
Dim lintPageNumber As Integer
Dim lbooOutOfStockCaption As Boolean
Dim lintLinesNumber As Integer 

    On Error GoTo ErrHandler
    If pstrMode = "LOCAL" Then
        lstrSQL = "SELECT * From " & gtblOrderLines & " Where " & gtblOrderLines & ".CustNum = " & _
        plngCustomerNum & " ORDER BY " & gtblOrderLines & ".BinLocation;"
    ElseIf pstrMode = "CENTRAL" Then
        lstrSQL = "SELECT * From " & gtblMasterOrderLines & " "
        lstrSQL = lstrSQL & "Where CustNum = " & _
        plngCustomerNum & " and OrderNum = " & plngOrderNumber & " ORDER BY BinLocation;"
    End If
    lintParcelNumber = 1
    glngTotalParcel = 1
    
    ReDim pstrProductLines(0)
    pstrProductLines(0).booLightParcel = False
    pstrProductLines(0).lngNumberOfLines = 0
    
    lstrStepError = "Step One"
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        glngGrossWeight = 0
        llngRecCount = 0
        lintPageNumber = 1
        lstrStepError = "Step Two"
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
                If llngRecCount > plngMaxproductLines Then
                    'New Page
                    llngRecCount = 1
                    lintPageNumber = lintPageNumber + 1
                    ReDim Preserve pstrProductLines(UBound(pstrProductLines) + 1)
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = 0
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                    lintLinesNumber = 0
                End If
                If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 18000 Then
                    If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 800 Then
                        pstrProductLines(UBound(pstrProductLines)).booLightParcel = True
                    Else
                        pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    End If
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                ElseIf (llngTotalWeight + Val(.Fields("TotalWeight"))) > 18000 And llngTotalWeight = 0 Then
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                Else
                    'New Parcel
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    llngTotalWeight = 0
                    llngRecCount = 1
                    lintParcelNumber = lintParcelNumber + 1
                    ReDim Preserve pstrProductLines(UBound(pstrProductLines) + 1)
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = 0
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                    lintLinesNumber = 0
                End If
            
                If Trim$(pstrProductLines(UBound(pstrProductLines)).strProductLines) <> "" Then
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & vbCrLf
                    lstrStepError = "Step three"
                End If
                
                If Val(.Fields("Qty")) = 0 Or Val(.Fields("DespQty")) = 0 Then
                    lbooOutOfStockCaption = True 'show out of stock caption
                Else
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                If Trim$(.Fields("CatNum")) = "REFUND" Then
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                AddparcelNumberArray plngOrderNumber, .Fields("CatNum"), _
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber
                
                'Block added 
                pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strCatCode = .Fields("CatNum")
                pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strBinLoc = .Fields("BinLocation") & "" 
                pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strQtyOrd = .Fields("Qty")
                pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strQtyDesp = .Fields("DespQty")
                pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strDesc = .Fields("ItemDescription")
                
                If lbooOutOfStockCaption = True Then
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & Spacer(.Fields("CatNum"), 6) & _
                        Spacer(.Fields("BinLocation"), 9) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer("***  Out of Stock  ***", 22, "L")
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strUnitPrice = "***  Out of Stock  ***"
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strTaxCode = ""
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strAmount = ""
                        
                Else
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & Spacer(.Fields("CatNum"), 8) & _
                        Spacer(.Fields("BinLocation") & "", 7) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer(AdvicePrice(.Fields("Price")), 8, "L") & " " & Spacer(.Fields("TaxCode"), 4, "L") & _
                        Spacer(AdvicePrice(.Fields("TotalPrice")), 10, "L")
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strUnitPrice = AdvicePrice(.Fields("Price"))
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strTaxCode = .Fields("TaxCode")
                    pstrProductLines(UBound(pstrProductLines)).strLines(lintLinesNumber).strAmount = AdvicePrice(.Fields("TotalPrice"))
                End If
                                                                            
                
                lintLinesNumber = lintLinesNumber + 1
                                                        
                lstrStepError = "Step Four"
                If .Fields("TaxCode") = "Z" Then
                    pstrZeroUnitTotal = AdvicePrice(PriceVal(pstrZeroUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                Else
                    lstrStepError = "Step Five"
                    pstrVatTotal = AdvicePrice(PriceVal(pstrVatTotal) + (Val(.Fields("DespQty")) * PriceVal(.Fields("Price")) * Val(gstrVATRate) / 100))
                    pstrUnitTotal = AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                End If
                            
                lstrStepError = "Step Six"
                pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                glngGrossWeight = glngGrossWeight + llngTotalWeight
            .MoveNext
        Loop
    End With
    
    glngTotalParcel = lintParcelNumber
    If llngRecCount = 0 Then
    End If
    lstrStepError = "Step Seven"
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
        
    Select Case GlobalErrorHandler(Err.Number, "PrintGetProductsOrderLines " & lstrStepError, "Local")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select
  
End Sub
Sub PrintGetProductsOrderlinesSAFE(plngCustomerNum As Long, ByRef pstrProductLines() As OrderDetail, _
    pstrUnitTotal As String, pstrVatTotal As String, plngOrderNumber As Long, _
    pstrMode As String, pstrZeroUnitTotal As String, plngMaxproductLines As Long)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrStepError As String
Dim llngTotalWeight As Long
Dim lintParcelNumber As Integer
Dim lintPageNumber As Integer
Dim lbooOutOfStockCaption As Boolean
    
    On Error GoTo ErrHandler
    If pstrMode = "LOCAL" Then
        lstrSQL = "SELECT * From " & gtblOrderLines & " Where " & gtblOrderLines & ".CustNum = " & _
        plngCustomerNum & " ORDER BY " & gtblOrderLines & ".BinLocation;"
    ElseIf pstrMode = "CENTRAL" Then
        lstrSQL = "SELECT * From " & gtblMasterOrderLines & " "
        lstrSQL = lstrSQL & "Where CustNum = " & _
        plngCustomerNum & " and OrderNum = " & plngOrderNumber & " ORDER BY BinLocation;"
    End If
    lintParcelNumber = 1
    glngTotalParcel = 1
    
    ReDim pstrProductLines(0)
    pstrProductLines(0).booLightParcel = False
    pstrProductLines(0).lngNumberOfLines = 0
    
    lstrStepError = "Step One"
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        glngGrossWeight = 0
        llngRecCount = 0
        lintPageNumber = 1
        lstrStepError = "Step Two"
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
                If llngRecCount > plngMaxproductLines Then
                    'New Page
                    llngRecCount = 1
                    lintPageNumber = lintPageNumber + 1
                    ReDim Preserve pstrProductLines(UBound(pstrProductLines) + 1)
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = 0
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                End If
                If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 18000 Then
                    If (llngTotalWeight + Val(.Fields("TotalWeight"))) < 800 Then
                        pstrProductLines(UBound(pstrProductLines)).booLightParcel = True
                    Else
                        pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    End If
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                ElseIf (llngTotalWeight + Val(.Fields("TotalWeight"))) > 18000 And llngTotalWeight = 0 Then
                    llngTotalWeight = llngTotalWeight + Val(.Fields("TotalWeight"))
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                Else
                    'New Parcel
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    llngTotalWeight = 0
                    llngRecCount = 1
                    lintParcelNumber = lintParcelNumber + 1
                    ReDim Preserve pstrProductLines(UBound(pstrProductLines) + 1)
                    pstrProductLines(UBound(pstrProductLines)).booLightParcel = False
                    pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = 0
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber = lintParcelNumber
                End If
            
                If Trim$(pstrProductLines(UBound(pstrProductLines)).strProductLines) <> "" Then
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & vbCrLf
                    lstrStepError = "Step three"
                End If
                
                If Val(.Fields("Qty")) = 0 Or Val(.Fields("DespQty")) = 0 Then
                    lbooOutOfStockCaption = True 'show out of stock caption
                Else
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                If Trim$(.Fields("CatNum")) = "REFUND" Then
                    lbooOutOfStockCaption = False ' Don't show out of stock caption
                End If
                
                AddparcelNumberArray plngOrderNumber, .Fields("CatNum"), _
                    pstrProductLines(UBound(pstrProductLines)).lngParcelBoxNumber
                
                If lbooOutOfStockCaption = True Then
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & Spacer(.Fields("CatNum"), 6) & _
                        Spacer(.Fields("BinLocation"), 9) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer("***  Out of Stock  ***", 22, "L")
                Else
                    pstrProductLines(UBound(pstrProductLines)).strProductLines = pstrProductLines(UBound(pstrProductLines)).strProductLines & Spacer(.Fields("CatNum"), 8) & _
                        Spacer(.Fields("BinLocation"), 7) & _
                        Spacer(.Fields("Qty"), 5, "L") & " " & Spacer(.Fields("DespQty"), 5, "L") & " " & _
                        Spacer(.Fields("ItemDescription"), 29) & _
                        Spacer(AdvicePrice(.Fields("Price")), 8, "L") & " " & Spacer(.Fields("TaxCode"), 4, "L") & _
                        Spacer(AdvicePrice(.Fields("TotalPrice")), 10, "L")
                End If
                                                        
                lstrStepError = "Step Four"
                If .Fields("TaxCode") = "Z" Then
                    pstrZeroUnitTotal = AdvicePrice(PriceVal(pstrZeroUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                Else
                    lstrStepError = "Step Five"
                    pstrVatTotal = AdvicePrice(PriceVal(pstrVatTotal) + (Val(.Fields("DespQty")) * PriceVal(.Fields("Price")) * Val(gstrVATRate) / 100))
                    pstrUnitTotal = AdvicePrice(PriceVal(pstrUnitTotal) + (PriceVal(.Fields("Price")) * .Fields("DespQty")))
                End If
                            
                lstrStepError = "Step Six"
                pstrProductLines(UBound(pstrProductLines)).lngNumberOfLines = llngRecCount
                glngGrossWeight = glngGrossWeight + llngTotalWeight
            .MoveNext
        Loop
    End With
    
    glngTotalParcel = lintParcelNumber
    If llngRecCount = 0 Then
    End If
    lstrStepError = "Step Seven"
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
        
    Select Case GlobalErrorHandler(Err.Number, "PrintGetProductsOrderLines " & lstrStepError, "Local")
    Case gconIntErrHandRetry
        Resume
    Case Else
        Resume Next
    End Select
  
End Sub

Sub RepBatchPickCountOrderWeights(ByRef pstrOrderNumBatch() As BatchLines)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim llngBatchTally As Long
Dim lintArrInc As Integer
Dim lstrMessage As String
Dim gconIntMaxBatchPickingsWeight As Long
    
    gconIntMaxBatchPickingsWeight = 18000
    On Error GoTo ErrHandler
   
    lstrSQL = "SELECT " & gtblMasterOrderLines & _
        ".OrderNum, Sum(Val([TotalWeight])) AS TW " & _
        "FROM " & gtblMasterOrderLines & " INNER JOIN " & gtblAdviceNotes & " ON " & _
        gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum " & _
        "Where (((" & gtblAdviceNotes & ".OrderStatus) = 'P') and (" & gtblAdviceNotes & ".PickPrinted = False)) " & _
        "GROUP BY " & gtblMasterOrderLines & ".OrderNum " & _
        "Having (((" & gtblMasterOrderLines & ".OrderNum) <> 0)) " & _
        "ORDER BY " & gtblMasterOrderLines & ".OrderNum;"
        
Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        ReDim Preserve pstrOrderNumBatch(0)
        llngRecCount = 0
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            If ((llngBatchTally + .Fields("TW")) < gconIntMaxBatchPickingsWeight) Or _
            (llngBatchTally = 0 And (llngBatchTally + .Fields("TW")) > gconIntMaxBatchPickingsWeight) Then
                llngBatchTally = llngBatchTally + .Fields("TW")
                
                If Trim$(pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause) = "" Then
                    pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause = _
                        pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause & _
                            " " & gtblMasterOrderLines & ".OrderNum = " & .Fields("OrderNum")
                Else
                    pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause = _
                        pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause & _
                        " OR " & gtblMasterOrderLines & ".OrderNum = " & .Fields("OrderNum")
                End If
                    
                pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strOrderNumberHeader = _
                    pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strOrderNumberHeader & _
                    .Fields("OrderNum") & ", "
                                            
                pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strBatchTotal = llngBatchTally
            Else
                ReDim Preserve pstrOrderNumBatch(UBound(pstrOrderNumBatch) + 1)
                pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause = _
                    pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strSQLWhereClause & _
                    " " & gtblMasterOrderLines & ".OrderNum = " & .Fields("OrderNum")
                llngBatchTally = .Fields("TW")
                pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strOrderNumberHeader = _
                    pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strOrderNumberHeader & _
                    .Fields("OrderNum") & ", "

                pstrOrderNumBatch(UBound(pstrOrderNumBatch)).strBatchTotal = llngBatchTally
            End If

            .MoveNext
        Loop

    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
     
    Select Case GlobalErrorHandler(Err.Number, "RepBatchPickCountorderWeights", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select


End Sub

Sub RepStockOrderSummary(pstrFilename As String, pdatStartDate As Date, pdatEndDate As Date, pstrParam As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrPrintString As String
'Converted table names to constants 

    On Error GoTo ErrHandler
    Select Case pstrParam
    Case "BC"
    lstrSQL = "SELECT " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & ".ItemDescription, " & _
        "" & gtblMasterOrderLines & ".BinLocation, " & gtblMasterOrderLines & ".SalesCode, " & _
        "Sum(" & gtblMasterOrderLines & ".Qty) AS SumOfQty, Sum(" & gtblMasterOrderLines & ".DespQty) " & _
        "AS SumOfDespQty FROM " & gtblMasterOrderLines & " INNER JOIN " & gtblAdviceNotes & " ON " & _
        "" & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum WHERE (((" & _
        "" & gtblAdviceNotes & ".CreationDate) <= #" & Format$(pdatEndDate, "DD/MMM/YYYY") & _
        "#) AND ((" & gtblAdviceNotes & ".OrderStatus)='C' Or (" & gtblAdviceNotes & ".OrderStatus)='B')) GROUP BY " & gtblMasterOrderLines & ".CatNum, " & _
        "" & gtblMasterOrderLines & ".ItemDescription, " & gtblMasterOrderLines & ".BinLocation, " & gtblMasterOrderLines & ".SalesCode " & _
        "ORDER BY " & gtblMasterOrderLines & ".CatNum;"
    Case "AP"
    lstrSQL = "SELECT " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & ".ItemDescription, " & _
        "" & gtblMasterOrderLines & ".BinLocation, " & gtblMasterOrderLines & ".SalesCode, " & _
        "Sum(" & gtblMasterOrderLines & ".Qty) AS SumOfQty FROM " & gtblMasterOrderLines & " INNER JOIN " & _
        "" & gtblAdviceNotes & " ON " & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum " & _
        "WHERE (((" & gtblAdviceNotes & ".CreationDate) " & _
        "<= #" & Format$(pdatEndDate, "DD/MMM/YYYY") & "#) AND ((" & gtblAdviceNotes & ".OrderStatus)='A' Or (" & gtblAdviceNotes & ".OrderStatus)='P')) " & _
        "GROUP BY " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & ".ItemDescription, " & _
        "" & gtblMasterOrderLines & ".BinLocation , " & gtblMasterOrderLines & ".SalesCode ORDER BY " & _
        "" & gtblMasterOrderLines & ".CatNum;"
        
        
    End Select
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            'Print #lintFileNum, .Fields("CatNum") & "," & _
                .Fields("ItemDescription") & "," & _
                .Fields("SumOfQty") & "," & _
                .Fields("SumOfDespQty") & "," & _
                .Fields("BinLocation") & "," & _
                .Fields("SalesCode") '& "," & .Fields("CreationDate")
                
            lstrPrintString = .Fields("CatNum") & vbTab & _
                .Fields("ItemDescription") & vbTab & _
                .Fields("SumOfQty") & vbTab
            
            Select Case pstrParam
            Case "BC"
                lstrPrintString = lstrPrintString & .Fields("SumOfDespQty") & vbTab
            End Select
            
            lstrPrintString = lstrPrintString & .Fields("BinLocation") & vbTab & _
                .Fields("SalesCode") '& vbtab & .Fields("CreationDate")
            
            Print #lintFileNum, lstrPrintString
            .MoveNext
        Loop
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RepStockOrderSummary", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select
End Sub
Sub RepSalesSummaryExport(pstrFilename As String, pdatStartDate As Date, pdatEndDate As Date)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
'Converted table names to constants 

    On Error GoTo ErrHandler
        
'    lstrSQL = "SELECT OrderLinesMaster.SalesCode, Sum([Price]*[Qty]) " & _
        "AS Total, Format([CreationDate],'mmmm') AS Month, " & _
        "Year([CreationDate]) AS Year FROM OrderLinesMaster INNER JOIN " & _
        "AdviceNotes ON OrderLinesMaster.OrderNum = AdviceNotes.OrderNum " & _
        "Where (((AdviceNotes.OrderStatus) = 'D' Or (AdviceNotes.OrderStatus) " & _
        "= 'E' Or (AdviceNotes.OrderStatus) = 'C' Or " & _
        "(AdviceNotes.OrderStatus) = 'B')) GROUP BY OrderLinesMaster" & _
        ".SalesCode, Format([CreationDate],'mmmm'), Year([CreationDate]) " & _
        "Having (((OrderLinesMaster.SalesCode) <> 0)) " & _
        "ORDER BY Year([CreationDate]), Format([CreationDate],'mmmm');"
        
    lstrSQL = "SELECT " & gtblMasterOrderLines & ".SalesCode, Sum([Price]*[DespQty]) " & _
        "AS Total, Format([CreationDate],'mmmm') AS Month, " & _
        "Year([CreationDate]) AS Year FROM " & gtblMasterOrderLines & " INNER JOIN " & _
        "" & gtblAdviceNotes & " ON " & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum " & _
        "Where (" & gtblAdviceNotes & ".CreationDate >= #" & pdatStartDate & _
        "# And " & gtblAdviceNotes & ".CreationDate <= #" & pdatEndDate & "#) AND " & _
        "(((" & gtblAdviceNotes & ".OrderStatus) = 'D' Or (" & gtblAdviceNotes & ".OrderStatus) " & _
        "= 'E' Or (" & gtblAdviceNotes & ".OrderStatus) = 'C' Or " & _
        "(" & gtblAdviceNotes & ".OrderStatus) = 'B')) GROUP BY " & gtblMasterOrderLines & "" & _
        ".SalesCode, Format([CreationDate],'mmmm'), Year([CreationDate]) " & _
        "Having (((" & gtblMasterOrderLines & ".SalesCode) <> 0)) " & _
        "ORDER BY Year([CreationDate]), Format([CreationDate],'mmmm');"
        
        
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        Print #lintFileNum, "Sales Code,Month,Year,Total"

        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            Print #lintFileNum, .Fields("SalesCode") & "," & .Fields("Month") & "," & .Fields("Year") & "," & AdvicePrice(.Fields("Total"))

            .MoveNext
        Loop
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RepSalesSummaryExport", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select

End Sub
Function RepBatchPickings(pstrFilename As String) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
'Dim lstrOrderLinesMaster As String
Dim pstrOrderNumBatch() As BatchLines
Dim lintArrInc As Integer
Dim lintLineNum As Integer
Dim lintArrInc2 As Integer
Dim lintMorePages As Integer

Const lconCatSpacer As Integer = 11
Const lconBinSpacer As Integer = 15
Const lconQtySpacer As Integer = 9
Const lconProdSpacer As Integer = 37
'Converted table names to constants 

    RepBatchPickings = True
    'Select Case gstrUserMode
    'Case gconstrLiveMode
        ''lstrOrderLinesMaster = "OrderLinesMaster"
    'Case gconstrTestingMode
    '    lstrOrderLinesMaster = "OrderLinesMasterTraining"
    'End Select
       
    On Error GoTo ErrHandler
    
    RepBatchPickCountOrderWeights pstrOrderNumBatch()
    For lintArrInc = 0 To UBound(pstrOrderNumBatch)
    
        If pstrOrderNumBatch(lintArrInc).strSQLWhereClause = "" Then
            MsgBox "No Orders found!", , gconstrTitlPrefix & "Batch Pickings"
            RepBatchPickings = False
            Exit Function
        End If
        lstrSQL = "SELECT Sum(Val([Qty])) AS Quantity, Sum(Val([TotalWeight])) AS " & _
            "TW, " & gtblMasterOrderLines & ".CatNum, " & gtblMasterOrderLines & "." & _
            "BinLocation, " & gtblMasterOrderLines & ".ItemDescription " & _
            "From " & gtblMasterOrderLines & _
            " INNER JOIN " & gtblAdviceNotes & " ON " & gtblMasterOrderLines & ".OrderNum = " & gtblAdviceNotes & ".OrderNum "
        
'        lstrSQL = lstrSQL & " where ((AdviceNotes.OrderStatus) = 'P') AND " & _
        pstrOrderNumBatch(lintArrInc).strSQLWhereClause
        
        'added 
        lstrSQL = lstrSQL & " where " & pstrOrderNumBatch(lintArrInc).strSQLWhereClause
                

        lstrSQL = lstrSQL & " GROUP BY " & gtblMasterOrderLines & "." & _
            "CatNum, " & gtblMasterOrderLines & ".BinLocation, " & _
            gtblMasterOrderLines & ".ItemDescription "
        'added 
        lstrSQL = lstrSQL & "ORDER BY " & gtblMasterOrderLines & ".BinLocation;"



        Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
            With lsnaLists
            
                llngRecCount = 0
                lintFileNum = FreeFile
                
                Open pstrFilename For Append As lintFileNum
                Print #lintFileNum, "Batch Picking Summary                           " & Format$(Now(), "DD/MMM/YYYY"): lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, "Order Nos. " & pstrOrderNumBatch(lintArrInc).strOrderNumberHeader: lintLineNum = lintLineNum + 1
                
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, Spacer("Cat. Code", lconCatSpacer) & _
                    Spacer("Face/Bin", lconBinSpacer) & _
                    Spacer("Quantity", lconQtySpacer, "L") & " " & _
                    Spacer("Product", lconProdSpacer) & _
                    "Weight": lintLineNum = lintLineNum + 1
                
                Print #lintFileNum, Spacer("---------", lconCatSpacer) & _
                    Spacer("--------", lconBinSpacer) & _
                    Spacer("--------", lconQtySpacer, "L") & " " & _
                    Spacer("-------", lconProdSpacer) & _
                    "------": lintLineNum = lintLineNum + 1
    
                Do Until .EOF
                    llngRecCount = llngRecCount + 1
                    
                    Print #lintFileNum, Spacer(.Fields("CatNum"), lconCatSpacer) & _
                    Spacer(.Fields("BinLocation"), lconBinSpacer) & _
                    Spacer(.Fields("Quantity"), lconQtySpacer, "L") & " " & _
                    Spacer(.Fields("ItemDescription"), lconProdSpacer) & _
                    Format$(Val(.Fields("TW")) / 1000, "0.000"): lintLineNum = lintLineNum + 1
        
                    .MoveNext
                Loop
            End With
        If llngRecCount = 0 Then
        End If
            
        lsnaLists.Close
        Print #lintFileNum, Spacer("", lconCatSpacer + lconBinSpacer + lconQtySpacer + lconProdSpacer) & "------": lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", lconCatSpacer + lconBinSpacer + lconQtySpacer + lconProdSpacer) & Format$(Val(pstrOrderNumBatch(lintArrInc).strBatchTotal) / 1000, "0.000") & "kg": lintLineNum = lintLineNum + 1
        '    Dim x As Integer
            
        If lintLineNum > gintNumberOfLinesAPage Then
            lintMorePages = CInt(lintLineNum / gintNumberOfLinesAPage)
            lintLineNum = gintNumberOfLinesAPage - ((gintNumberOfLinesAPage * lintMorePages) - lintLineNum)
        Else
            lintMorePages = 0
        End If
        
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            Print #lintFileNum, ""
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        'Debug.Print lintLineNum
        lintLineNum = 0
        
        Close #lintFileNum
    Next lintArrInc
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RepBatchPickings", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select


End Function
Sub PrintAdviceNotesGeneral(pdatStartDate As Date, pdatEndDate As Date, pstrParamater As String, _
    pstrFilename As String, Optional plngOrderNum As Variant, Optional plngEndOrderNum As Variant)
    
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Dim lintFileNum As Integer
'Converted table names to constants 
'Also removed unnecessary References

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
    Case "A"
        lstrSQL = lstrSQL & " and OrderStatus = 'A' "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
    Case "P"
        lstrSQL = lstrSQL & " and OrderStatus = 'P' "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
    Case "S"
        lstrSQL = lstrSQL & " Ordernum = " & plngOrderNum & " "
        lstrSQL = lstrSQL & "and OrderNum <> CLng(0) and OrderNum <> null "
    Case "R"
        'Added 
        lstrSQL = lstrSQL & " (((OrderStatus)='A') AND " & _
            "((OrderNum)>=" & plngOrderNum & _
            " And (OrderNum)<=" & plngEndOrderNum & ")) "
    Case Else
       
    End Select
        
    lstrSQL = lstrSQL & "order by OrderNum;" 'CreationDate;"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
    
        llngRecCount = 0
        glngLastOrderPrintedInThisRun = 0
        
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
            
            PrintToFileAdviceInvoice pstrFilename, "CENTRAL"
            glngLastOrderPrintedInThisRun = .Fields("OrderNum")
            
            If glngItemsWouldLikeToPrint > 0 Then 
                If glngItemsWouldLikeToPrint = llngRecCount Then
                    lsnaLists.Close
                    Exit Sub
                End If
            End If
            
            .MoveNext
        Loop
                
    End With
    
    If llngRecCount = 0 Then
        MsgBox "No records found within the specified criteria.", , gconstrTitlPrefix & "General Advice Note Print"
    End If
    
    lsnaLists.Close
        
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintAdviceNotesGeneral", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select
            


End Sub
Function PrintCreditCardClaim(pstrFilename As String) As String
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrAddress As String
Dim lcurTotal As Currency
Dim lcurPageTotal As Currency
Dim lintArrInc2 As Integer
Dim lintLineNum As Integer
Dim lintPageNum As Integer
Dim lcurTotalCCPayment As Currency
'Converted table names to constants 

    ReDim Preserve glngTrackNUpdate(0)
    On Error GoTo ErrHandler


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
        
    
    lstrSQL = "SELECT CardNumber, TotalIncVat, AuthorisationCode, OrderNum, CreationDate, DespatchDate, " & _
        "OrderType, Payment, PaymentType2, Payment2, CallerSalutation, CallerInitials, CallerSurname, " & _
        "AdviceAdd1, AdviceAdd2, AdviceAdd3, AdviceAdd4, AdviceAdd5, AdvicePostcode, BankRepPrintDate, " & _
        "CardType From " & gtblAdviceNotes & " WHERE (((BankRepPrintDate)=0 Or (BankRepPrintDate) Is Null) AND " & _
        "((OrderStatus)='C' Or (OrderStatus)='B' Or (OrderStatus)='D' Or (OrderStatus)='E') AND " & _
        "((Trim$([OrderType]))='C') AND ((CardType)<>'SWITCH')) OR (((BankRepPrintDate)=0 Or " & _
        "(BankRepPrintDate) Is Null) AND ((OrderStatus)='C' Or (OrderStatus)='B' Or (OrderStatus)='D' " & _
        "Or (OrderStatus)='E') AND ((Trim$([PaymentType2]))='C') AND ((CardType)<>'SWITCH')) " & _
        "ORDER BY OrderNum;"

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lcurTotal = 0
        lcurPageTotal = 0
        lintPageNum = 1
        lintFileNum = FreeFile
        
        Open pstrFilename For Output As lintFileNum
        
        'Print #lintFileNum, "MIDLAND CARD SERV CHARGE-CARD CLAIMS My Company Plc. - Merchant No. 1234    as at " & Now(): lintLineNum = lintLineNum + 1
        'Print #lintFileNum, "CONTACT - Joe Bloggs    Telephone 0123 4567890": lintLineNum = lintLineNum + 1
       
        Print #lintFileNum, Trim$(gstrReferenceInfo.strCreditCardClaimsHead1A) & Trim$(gstrReferenceInfo.strCreditCardClaimsHead1B) & "    as at " & Now(): lintLineNum = lintLineNum + 1
        Print #lintFileNum, gstrReferenceInfo.strCreditCardClaimsHead2A: lintLineNum = lintLineNum + 1
        Print #lintFileNum, " ": lintLineNum = lintLineNum + 1

        Print #lintFileNum, Spacer("Credit Card Number", 20) & _
            Spacer("Amount", 6) & " " & _
            Spacer("Auth.Code", 10) & _
            Spacer("Order No.", 10) & _
            Spacer("Desp Date", 10) & _
            "Customer": lintLineNum = lintLineNum + 1

        Do Until .EOF
            llngRecCount = llngRecCount + 1
                                
            If UCase$(Trim$(.Fields("OrderType"))) = "C" Then
                lcurTotalCCPayment = CCur(.Fields("Payment"))
            End If
            
            If UCase$(Trim$(.Fields("PaymentType2"))) = "C" Then
                lcurTotalCCPayment = CCur(.Fields("Payment2"))
            End If
            
            lstrAddress = ""
            If Trim$(.Fields("AdviceAdd1")) <> "" Then lstrAddress = Trim$(.Fields("AdviceAdd1")) & ", "
            If Trim$(.Fields("AdviceAdd2")) <> "" Then lstrAddress = lstrAddress & Trim$(.Fields("AdviceAdd2")) & ", "
            If Trim$(.Fields("AdviceAdd3")) <> "" Then lstrAddress = lstrAddress & Trim$(.Fields("AdviceAdd3")) & ", "
            If Trim$(.Fields("AdviceAdd4")) <> "" Then lstrAddress = lstrAddress & Trim$(.Fields("AdviceAdd4")) & ", "
            If Trim$(.Fields("AdviceAdd5")) <> "" Then lstrAddress = lstrAddress & Trim$(.Fields("AdviceAdd5")) & ", "
            If Trim$(.Fields("AdvicePostcode")) <> "" Then lstrAddress = lstrAddress & Trim$(.Fields("AdvicePostcode"))
            
            Print #lintFileNum, Spacer(FormatCardNum(.Fields("CardNumber") & ""), 20) & _
                Spacer(Trim$(Format$(lcurTotalCCPayment, "0.00") & ""), 6, "L") & " " & _
                Spacer(Trim$(.Fields("AuthorisationCode") & ""), 10) & _
                Spacer(Trim$(.Fields("OrderNum")), 10) & _
                Spacer(Format$(.Fields("DespatchDate") & "", "DD/MM/YY"), 10) & _
                Spacer(Trim$(Trim$(.Fields("CallerSalutation") & "") & " " & _
                Trim$(.Fields("CallerInitials") & "") & " " & _
                Trim$(.Fields("CallerSurname") & "")) & ", " & lstrAddress, 76): lintLineNum = lintLineNum + 1
                                
            Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
            
            lcurPageTotal = lcurPageTotal + CCur(lcurTotalCCPayment)

            
            If lintLineNum = ((lintPageNum * gintNumberOfLinesAPage) - 8) Then
                        
                Print #lintFileNum, Spacer("", 21) & Spacer("---------", 7): lintLineNum = lintLineNum + 1
                Print #lintFileNum, Spacer("Page Total", 21) & Format$(lcurPageTotal, "0.00"): lintLineNum = lintLineNum + 1
                lcurPageTotal = 0
                Print #lintFileNum, Spacer("", 21) & Spacer("---------", 7): lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1

                Print #lintFileNum, Spacer("", 60) & "Page No. " & lintPageNum: lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                lintPageNum = lintPageNum + 1
                
                'Print #lintFileNum, "MIDLAND CARD SERV CHARGE-CARD CLAIMS My Company Plc. - Merchant No. 1234    as at " & Now(): lintLineNum = lintLineNum + 1
                'Print #lintFileNum, "CONTACT - Joe Bloggs    Telephone 0123 4567890": lintLineNum = lintLineNum + 1
                Print #lintFileNum, Trim$(gstrReferenceInfo.strCreditCardClaimsHead1A) & Trim$(gstrReferenceInfo.strCreditCardClaimsHead1B) & "    as at " & Now(): lintLineNum = lintLineNum + 1
                Print #lintFileNum, gstrReferenceInfo.strCreditCardClaimsHead2A: lintLineNum = lintLineNum + 1
        
                Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
    
                Print #lintFileNum, Spacer("Credit Card Number", 20) & _
                    Spacer("Amount", 6) & " " & _
                    Spacer("Auth.Code", 10) & _
                    Spacer("Order No.", 10) & _
                    Spacer("Desp Date", 10) & _
                    "Customer": lintLineNum = lintLineNum + 1
            End If
                                
            lcurTotal = lcurTotal + CCur(lcurTotalCCPayment)
            
            
            If llngRecCount = 1 Then
                glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = .Fields("OrderNum")
            Else
                ReDim Preserve glngTrackNUpdate(UBound(glngTrackNUpdate) + 1)
                glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = .Fields("OrderNum")
            End If
            
            .MoveNext
        Loop
                
        If lcurPageTotal > 0 Then
            Print #lintFileNum, Spacer("", 21) & Spacer("---------", 7): lintLineNum = lintLineNum + 1
            Print #lintFileNum, Spacer("Page Total", 21) & Format$(lcurPageTotal, "0.00"): lintLineNum = lintLineNum + 1
            Print #lintFileNum, Spacer("", 21) & Spacer("---------", 7): lintLineNum = lintLineNum + 1
            Print #lintFileNum, "": lintLineNum = lintLineNum + 1
        End If
        
        Print #lintFileNum, Spacer("", 21) & Spacer("----------", 8): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("Grand Total", 21) & Format$(lcurTotal, "0.00"): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", 21) & Spacer("----------", 8): lintLineNum = lintLineNum + 1
                
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            Print #lintFileNum, ""
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintCreditCardClaim", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub PrintChequeAllocation(pstrFilename As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrAddress As String
Dim lcurTotal As Currency
Dim lintArrInc2 As Integer
Dim lintLineNum As Integer
Dim lintPageNum As Integer
'Converted table names to constants 
'Also removed unnecessary references

    On Error GoTo ErrHandler

    'lstrSQL = "SELECT ChequeNum, PrintedDate, Amount, Name, CustNum, OrderNum, Reason " & _
        "From " & gtblCashBook & " Where (((PrintedDate) <> 0)) ORDER BY ChequeNum;"
    
    lstrSQL = "SELECT ChequeNum, ChequePrintedDate as PrintedDate, " & _
        "Reconcilliation + Underpayment as Amount, CardName as Name, CustNum, OrderNum, RefundReason as Reason " & _
        "From " & gtblAdviceNotes & " Where (((ChequePrintedDate) <> 0)) ORDER BY ChequeNum;"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lcurTotal = 0
        lintPageNum = 1
        
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        Print #lintFileNum, "CHEQUE ALLOCATION REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
        Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
        
        Print #lintFileNum, Spacer("ChequeNum", 10) & _
            Spacer("Date", 10) & _
            Spacer("Value", 8) & _
            Spacer("Name", 50) & _
            Spacer("CustNum", 9) & _
            Spacer("OrderNum", 9) & _
            Spacer("Reason", 25): lintLineNum = lintLineNum + 1
                
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            
            Print #lintFileNum, Spacer(.Fields("ChequeNum") & "", 10) & _
                Spacer(Format$(.Fields("PrintedDate") & "", "DD/MM/YY"), 10) & _
                Spacer(Trim$(Format$(.Fields("Amount"), "0.00") & ""), 7, "L") & " " & _
                Spacer(Trim$(.Fields("Name")), 50) & _
                Spacer(Trim$(.Fields("CustNum")), 9) & _
                Spacer(Trim$(.Fields("OrderNum")), 9) & _
                Spacer(Trim$(.Fields("Reason")), 25): lintLineNum = lintLineNum + 1
            
            
            If lintLineNum = lintPageNum * (gintNumberOfLinesAPage - 2) Then
                Print #lintFileNum, Spacer("", 70) & "Page No. " & lintPageNum: lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                lintPageNum = lintPageNum + 1
                Print #lintFileNum, "CHEQUE ALLOCATION REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
                Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
    
                Print #lintFileNum, Spacer("ChequeNum", 10) & _
                    Spacer("Date", 10) & _
                    Spacer("Value", 8) & _
                    Spacer("Name", 50) & _
                    Spacer("CustNum", 9) & _
                    Spacer("OrderNum", 9) & _
                    Spacer("Reason", 25): lintLineNum = lintLineNum + 1
            End If
            lcurTotal = lcurTotal + CCur(.Fields("Amount"))
            
            .MoveNext
        Loop
                        
        Print #lintFileNum, Spacer("", 20) & Spacer("--------", 8): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", 20) & lcurTotal: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", 20) & Spacer("--------", 8): lintLineNum = lintLineNum + 1
                        
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            Print #lintFileNum, ""
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintChequeAllocation", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select



End Sub
Sub PrintCheques(pstrFilename As String, plngFirstChequeNumber As Long)
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
'Converted table names to constants 
'Also remove uneccessary references

    ReDim Preserve glngChequeOrderNumPrinted(0)
    On Error GoTo ErrHandler

    lstrSQL = "SELECT Amount, Name, OrderNum, PrintedDate From " & gtblCashBook & " " & _
        "WHERE (((Amount)<>0) AND ((PrintedDate)=0) AND ((Reason)<>'UNDERPAY')) OR " & _
        "(((Amount)<>0) AND ((PrintedDate) Is Null) AND ((Reason)<>'UNDERPAY')) " & _
        "ORDER BY OrderNum;"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lstrTensOfThousands = ""
        lstrThousands = ""
        lstrHundreds = ""
        lstrTens = ""
        lstrUnits = ""
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
                        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, Spacer("", 57) & Format$(Now(), "DD/MMM/YYYY")
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, " " & Spacer(.Fields("OrderNum"), 8) & Spacer(Trim$(.Fields("Name")), 50) & Format(.Fields("Amount"), "0.00")
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            Print #lintFileNum, ""
            'If .Fields("OrderNum") = 136 Then MsgBox "Wait"
            
            If InStr(1, "" & .Fields("Amount") & "", ".") > 0 Then
                lcurPounds = CCur(Left$("" & .Fields("Amount") & "", InStr(1, "" & .Fields("Amount") & "", ".")))
            Else
                lcurPounds = CCur(.Fields("Amount"))
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
            
            Print #lintFileNum, " " & Spacer(lstrTensOfThousands, 5) & " " & _
                Spacer(lstrThousands, 5) & " " & _
                Spacer(lstrHundreds, 5) & " " & _
                Spacer(lstrTens, 5) & " " & _
                Spacer(lstrUnits, 5)

                'Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""
                Print #lintFileNum, ""

            If llngRecCount = 1 Then
                'ReDim Preserve glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted) + 1)
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngOrderNum = .Fields("OrderNum")
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngChequeNum = plngFirstChequeNumber
                
            Else
                ReDim Preserve glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted) + 1)
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngOrderNum = .Fields("OrderNum")
                glngChequeOrderNumPrinted(UBound(glngChequeOrderNumPrinted)).lngChequeNum = plngFirstChequeNumber + (llngRecCount - 1)
            
            End If
            .MoveNext
        Loop
                        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintCheques", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select



End Sub
Sub PrintMiscAdjustments(pstrFilename As String, pbooPaidFlag As Boolean)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintFileNum As Integer
Dim lstrAddress As String
Dim lcurTotal As Currency
Dim lintArrInc2 As Integer
Dim lintLineNum As Integer
Dim lintPageNum As Integer
'Converted table names to constants 
'Also removed uneccessary references

    On Error GoTo ErrHandler

    'lstrSQL = "SELECT OrderNum, CustNum, Name, Amount, ClearedDate From " & gtblCashBook & " "
    
    lstrSQL = "SELECT OrderNum, CustNum, CardName as Name, UnderPayment as Amount," & _
        "ChequeClearedDate as ClearedDate From " & gtblAdviceNotes & " "
    If pbooPaidFlag = True Then
        'lstrSQL = lstrSQL & "WHERE (((Reason)='UNDERPAY') AND " & _
            "((ClearedDate)<> null));"
        
        lstrSQL = lstrSQL & "WHERE (((RefundReason)='UNDERPAY') AND " & _
            "((ChequeClearedDate)<> null));"
    ElseIf pbooPaidFlag = False Then
        'lstrSQL = lstrSQL & "WHERE (((Reason)='UNDERPAY') AND " & _
            "((ClearedDate) is null));"
        
        lstrSQL = lstrSQL & "WHERE (((RefundReason)='UNDERPAY') AND " & _
            "((ChequeClearedDate) is null));"
    End If
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        lcurTotal = 0
        lintPageNum = 1
        lintFileNum = FreeFile
        Open pstrFilename For Output As lintFileNum
        
        Print #lintFileNum, "MISC ADJUSTMENTS REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
        Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
            
        Print #lintFileNum, Spacer("CustNum", 9) & _
            Spacer("OrderNum", 9) & _
            Spacer("Customer Name", 50) & _
            Spacer("Amount", 8) & _
            Spacer("Cleared Date", 10) & _
            "": lintLineNum = lintLineNum + 1
                
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            
            Print #lintFileNum, Spacer(Trim$(.Fields("CustNum")), 9) & _
                Spacer(Trim$(.Fields("OrderNum")), 9) & _
                Spacer(Trim$(.Fields("Name")), 50) & _
                Spacer(Trim$(.Fields("Amount") & ""), 8) & _
                Spacer(Format$(.Fields("ClearedDate") & "", "DD/MM/YY"), 10) & _
                "": lintLineNum = lintLineNum + 1
                                            
            If lintLineNum = lintPageNum * (gintNumberOfLinesAPage - 2) Then
                Print #lintFileNum, Spacer("", 70) & "Page No. " & lintPageNum: lintLineNum = lintLineNum + 1
                Print #lintFileNum, "": lintLineNum = lintLineNum + 1
                lintPageNum = lintPageNum + 1
                Print #lintFileNum, "MISC ADJUSTMENTS REPORT             Printed:" & Now() & "  By: " & gstrGenSysInfo.strUserName: lintLineNum = lintLineNum + 1
                Print #lintFileNum, " ": lintLineNum = lintLineNum + 1
    
                Print #lintFileNum, Spacer("CustNum", 9) & _
                    Spacer("OrderNum", 9) & _
                    Spacer("Customer Name", 50) & _
                    Spacer("Amount", 8) & _
                    Spacer("Cleared Date", 10) & _
                    "": lintLineNum = lintLineNum + 1
            End If
                                            
            lcurTotal = lcurTotal + CCur(.Fields("Amount"))
            .MoveNext
        Loop
                        
                        
        Print #lintFileNum, Spacer("", 68) & Spacer("--------", 8): lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", 68) & lcurTotal: lintLineNum = lintLineNum + 1
        Print #lintFileNum, Spacer("", 68) & Spacer("--------", 8): lintLineNum = lintLineNum + 1
                        
        For lintArrInc2 = 1 To gintNumberOfLinesAPage - lintLineNum
            Print #lintFileNum, ""
            lintLineNum = lintLineNum + 1
        Next lintArrInc2
        
        Close lintFileNum
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintMiscAdjustments", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select

End Sub
Sub PrintChequeRefundAdviceNotes(pstrFilename As String, pdatChequeprintDate As Date)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Dim lintFileNum As Integer
'Converted table names to constants 
'Also removed unneccassary references

    On Error GoTo ErrHandler

    
    ReDim gstrOrderLineParcelNumbers(0)
       
    lstrSQL = "SELECT *, " & _
        "Format$([PrintedDate],'dd/mm/yy') AS Expr1 From " & gtblCashBook & " " & _
        "WHERE (((Amount)<>0) AND " & _
        "((Format$([PrintedDate],'dd/mm/yy'))=CDate('" & Format$(pdatChequeprintDate, "dd/mm/yy") & "')) " & _
        "AND ((Reason)<>'UNDERPAY')) ORDER BY OrderNum;"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
    
        llngRecCount = 0
        glngLastOrderPrintedInThisRun = 0
        
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
            
            PrintToFileAdviceInvoice pstrFilename, "CENTRAL", "REFUND"
            glngLastOrderPrintedInThisRun = .Fields("OrderNum")
            .MoveNext
        Loop
                
    End With
    
    If llngRecCount = 0 Then
        MsgBox "No Refunds to print Advice notes", , gconstrTitlPrefix & "Refund Advice Notes"
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PrintChequeRefundAdvicveNotes", "Central")
    Case gconIntErrHandRetry
        Resume
    'Case gconIntErrHandExitFunction
    '    Exit Function
    Case Else
        Resume Next
    End Select
            
End Sub
Sub UpdateBatchPrintedFlag()
Dim lstrSQL As String
'Converted table names to constants 
'Also removed uneccassasry references

    lstrSQL = "UPDATE " & gtblAdviceNotes & " SET PickPrinted = True " & _
        "WHERE (((OrderStatus)='P') AND ((PickPrinted)=False));"

    gdatCentralDatabase.Execute lstrSQL
    
End Sub

Sub UpdateOrderLineParcelNumber()
Dim lintArrInc As Integer
Dim lstrSQL As String
'Converted table names to constants 
'Also removed uneccassasry references

    On Error GoTo ErrHandler

    For lintArrInc = 0 To UBound(gstrOrderLineParcelNumbers)
        With gstrOrderLineParcelNumbers(lintArrInc)
            If .lngOrderNum <> 0 Then
                lstrSQL = "UPDATE " & gtblMasterOrderLines & " SET ParcelNumber = " & .lngParcelBoxNumber & _
                    " WHERE (((OrderNum)=" & .lngOrderNum & _
                    ") AND ((TRIM(CatNum))='" & Trim(.strCatNum) & "'));"
                gdatCentralDatabase.Execute lstrSQL
            End If
        End With
    Next lintArrInc
            
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateOrderLineParcelNumber", "Central", True)
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

Sub GetPrinterInfo(pstrPortNumber As String, pintNumOfLinesAPage As Integer, pobjForm As Form)

'    Load frmPrinter
    pobjForm.LinesPerPage = pintNumOfLinesAPage
    pobjForm.LPTPort = pstrPortNumber
    pobjForm.Show vbModal
    
    gintNumberOfLinesAPage = pobjForm.LinesPerPage
    gstrLPTPortNumber = pobjForm.LPTPort
    
End Sub
