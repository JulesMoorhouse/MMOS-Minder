Attribute VB_Name = "modPForce"
Option Explicit

Global gbooTestingVarRanAlready As Boolean
Global gintTestingVarRanAlready As Integer

Global gstrPForceServiceInd As ListVars

Type ParcelNumber
    strAlphaPrefix              As String * 2
    strNumber6Dig               As String * 6
    strChkDigit                 As String * 1
End Type

Global gstrParcelNumber     As ParcelNumber
Global gstrConsignRangeStart As String * 6
Global gstrConsignRangeEnd  As String * 6
Global gstrBatchIncrNum     As String * 4

Type PFContractDetails
    strContractNumber       As String * 7
    strPFAcctNumber         As String 'Not used in Program
    strHalconCollectionID   As String * 6
    strHalconCustumerCode   As String * 8
End Type
Global gstrPFContractDetails As PFContractDetails

Type PForce
    lngCustNum              As Long
    lngOrderNum             As Long
    strStatus               As String * 1
    strConsignNum           As String * 9
    strServiceID            As String * 10
    strBatchNumber          As String * 4
    datDespatchDate         As Date
    strDeliverySalutation   As String * 15
    strDeliverySurname      As String * 25
    strDeliveryInitials     As String * 20
    strDeliveryAdd1         As String * 30
    strDeliveryAdd2         As String * 30
    strDeliveryAdd3         As String * 30
    strDeliveryAdd4         As String * 30
    strDeliveryAdd5         As String * 30
    strDeliveryPostcode     As String * 9
    intParcelItems          As Integer
    lngGrossWeight          As Long
    strWeekendHandCode      As String * 10
    strPrepaidInd           As String * 1
    strNotificationCode     As String * 10
    strConsignRemark        As String * 100 
    strSpecialSatDel        As String * 1
    strSpecialBookIn        As String * 1
    strSpecialProof         As String * 1
End Type
Global gstrPForceConsignment As PForce

Type ThermalLabel
    lngCustNum          As Long
    lngOrderNum         As Long
    strSalutation       As String * 15
    strSurname          As String * 25
    strInitials         As String * 20
    strAdd1             As String * 30
    strAdd2             As String * 30
    strAdd3             As String * 30
    strAdd4             As String * 30
    strAdd5             As String * 30
    strPostcode         As String * 9
    intNumberOfParcels  As Integer
End Type
Global gstrThermalLabel As ThermalLabel

Type FileHeader
    strRecordTypeInd    As String * 1
    strFileVersion      As String * 2
    strFileType         As String * 4
End Type
Global gstrFileHeader As FileHeader

Type SenderRecord
    strRecordTypeInd    As String * 1
    strFileVersion      As String * 2
    strSenderName       As String * 18
    strSenderAdd1       As String * 24
    strSenderAdd2       As String * 24
    strSenderAdd3       As String * 24
    strSenderPostCode   As String * 8
End Type
Global gstrSenderRecord As SenderRecord

Type DetailRecord
    strRecordTypeInd    As String * 1
    strFileVersion      As String * 2
    strWeekendHand      As String * 4
    strPostCodeKeyed    As String * 8 'Unused
    strPostTownDerived  As String * 20 'Unused
End Type
Global gstrDetailRecord As DetailRecord

Type DetailSupplementRecord
    strRecordTypeInd    As String * 1
    strFileVersion      As String * 2
End Type
Global gstrDetailSupplementRecord As DetailSupplementRecord

Type TrailerRecord
    strRecordTypeInd    As String * 1
    strFileVersion      As String * 2
End Type
Global gstrTrailerRecord As TrailerRecord

Function CalcChkDigitConsign(pstrConsignmentRealNumber As String) As String
Dim lintDigits(5) As Integer
Dim lintArrInc As Integer
Dim lintsubTotal As Integer
Dim lintRemainder As Integer
Dim lintSubtractor As Integer

    For lintArrInc = 0 To 5
        lintDigits(lintArrInc) = Mid$(pstrConsignmentRealNumber, lintArrInc + 1, 1)
    Next lintArrInc
    
    lintDigits(0) = lintDigits(0) * 4
    lintDigits(1) = lintDigits(1) * 2
    lintDigits(2) = lintDigits(2) * 3
    lintDigits(3) = lintDigits(3) * 5
    lintDigits(4) = lintDigits(4) * 9
    lintDigits(5) = lintDigits(5) * 7
    
    For lintArrInc = 0 To 5
        lintsubTotal = lintsubTotal + lintDigits(lintArrInc)
    Next lintArrInc
    
    lintRemainder = lintsubTotal Mod 11
    lintSubtractor = 11 - lintRemainder
    
    Select Case lintSubtractor
    Case 10
        CalcChkDigitConsign = 0
    Case 11
        CalcChkDigitConsign = 5
    Case Else
        CalcChkDigitConsign = lintSubtractor
    End Select
'Thus, given a 6 digit number of 162738 the check digit calculation is as follows:
'5)  Check digit = 3

End Function

Sub GetConsignBatchIncrs()
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Converted table names to constants 
'Also removed unnecessary table. references

    ShowStatus 85
    
    On Error GoTo ErrHandler
    lstrSQL = "SELECT Item, Value, DateCreated FROM " & gtblSystem & ";"

    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            Select Case Trim$(UCase$(.Fields("Item")))
            Case "LASTPFCONSIGNNUMINCR"
                gstrParcelNumber.strNumber6Dig = Trim$(.Fields("Value"))
            Case "BATCHINCR"
                gstrBatchIncrNum = Val(.Fields("Value"))
            End Select
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetConsignBatchIncr", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub

Function GetConsignmentRange()
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    ShowStatus 86
    On Error GoTo ErrHandler
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".* " & _
        "FROM " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & gtblListDetails & ".ListNum = " & _
        "" & gtblLists & ".ListNum WHERE " & gtblLists & ".ListName='PForce Consignment Range' and " & _
        "" & gtblListDetails & ".InUse = True;"

    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1

            Select Case Trim$(UCase$(.Fields("ListCode")))
            Case "ALPHAPREF"
                gstrParcelNumber.strAlphaPrefix = Trim$(UCase$(.Fields("Description")))
            Case "START"
                gstrConsignRangeStart = Left$(.Fields("Description"), 6)
            Case "END"
                gstrConsignRangeEnd = Left$(.Fields("Description"), 6)
            End Select
            
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CreateNextConsignmentNum", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function

Function GetListVarsAll(pstrListVars As ListVars)
Dim lsnaLists As Recordset
Dim lstrSQL As String
        
    On Error GoTo ErrHandler
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".ListCode, " & gtblListDetails & ".Description, " & _
        "" & gtblListDetails & ".UserDef1, " & gtblListDetails & ".UserDef2 FROM " & gtblLists & " INNER JOIN " & _
        "" & gtblListDetails & " ON " & gtblLists & ".ListNum = " & gtblListDetails & ".ListNum WHERE (((trim(" & gtblLists & ".ListName))='" & _
        Trim(pstrListVars.strListName) & "') AND ((trim(" & gtblListDetails & ".ListCode))='" & _
        Trim(pstrListVars.strListCode) & "'));"
                
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        If Not .NoMatch Then
            pstrListVars.strDescription = .Fields("Description") & ""
            pstrListVars.strUserDef1 = .Fields("UserDef1") & ""
            pstrListVars.strUserDef2 = .Fields("UserDef2") & ""
        End If
    End With
    
    lsnaLists.Close

    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetListVarsAll", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function

Sub IncrementConsignBatchIncrs()

    ShowStatus 87

    gstrParcelNumber.strNumber6Dig = Val(gstrParcelNumber.strNumber6Dig) + 1
    
    If Val(gstrParcelNumber.strNumber6Dig) >= Val(gstrConsignRangeEnd) Then
        MsgBox "WARNING: You have reached the end of your Parcel Force " & vbCrLf & _
            "Consignment number range!  Please check with your " & vbCrLf & _
            "Parcelforce representative, check that you may restart at " & vbCrLf & _
            "the begin of your range again." & vbCrLf & vbCrLf & _
            "This warning has been added as a requirement from " & vbCrLf & _
            "Parcelforce. (Press OK to continue!)", vbExclamation, gconstrTitlPrefix & "PF Consignment Range"
            
        gstrParcelNumber.strNumber6Dig = Val(gstrConsignRangeStart) + 1
    End If
    
    'Add Check digit
    gstrParcelNumber.strChkDigit = CalcChkDigitConsign(Format$(gstrParcelNumber.strNumber6Dig, "000000"))
    'All leading zeros
    gstrParcelNumber.strNumber6Dig = Format$(gstrParcelNumber.strNumber6Dig, "000000")
    
    'Batch Number
    gstrBatchIncrNum = Val(gstrBatchIncrNum) + 1
    
    If Val(gstrBatchIncrNum) = 9999 Then
        gstrBatchIncrNum = 1
    End If
    
    'All leading zeros
    gstrBatchIncrNum = Format$(gstrBatchIncrNum, "0000")
    
                    
End Sub

Sub PFPCLFile(pstrFilename As String)
Dim lintFileNum As Integer
Dim lintArrInc As Integer
Dim lstrDeliveryName As String
Dim lstrConsignmentNoteLine1 As String
Dim lstrConsignmentNoteLine2 As String

    ShowStatus 88
    
    With gstrPForceConsignment
    lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & _
        Trim$(.strDeliveryInitials) & " " & _
        Trim$(.strDeliverySurname))

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    
    For lintArrInc = 1 To .intParcelItems
    
        Print #lintFileNum, ":Printer Settings..."                                      '1.
        Print #lintFileNum, "WN"                                                        '2.
        Print #lintFileNum, "N"                                                         '3.
        Print #lintFileNum, "R0,0"                                                      '4.
        Print #lintFileNum, "N"                                                         '5.
        Print #lintFileNum, "I8,0,061"                                                  '6.
        Print #lintFileNum, "S2"                                                        '7.
        Print #lintFileNum, "D4"                                                        '8.
        Print #lintFileNum, "ZT"                                                        '9.
        Print #lintFileNum, "Q1242,020"                                                 '10.
        
        Print #lintFileNum, ":Carrier Logo..."                                          '11.
        Print #lintFileNum, "GG50,1090," & Chr(34) & "BLOB400" & Chr(34)                '12.
        
        Print #lintFileNum, ":Service Indicator..."                                     '13.
        
        Print #lintFileNum, "A790,1195,2,A,1,1,N," & Chr(34) & _
            Trim$(gstrPForceServiceInd.strUserDef1) & Chr(34)                                  '14.
        Print #lintFileNum, "LO530,975,290,4"                                           '15.
        Print #lintFileNum, "LO530,975,4,232"                                           '16.
        
        Print #lintFileNum, ":739 Bullet(type 1)..."                                    '17.
        Print #lintFileNum, "A125,1005,1,1,1,1,N," & Chr(34) & "739" & Chr(34)          '18.
        Print #lintFileNum, "A460,1030,2,4,1,1,N," & Chr(34) & _
            "PB " & .strConsignNum & " " & Format$(lintArrInc, "000") & Chr(34)        '19.
        
        
        Print #lintFileNum, ":Despatch Date..."                                         '20.
        Print #lintFileNum, "A270,935,2,3,1,1,N," & Chr(34) & _
            "Despatch Date/Day" & Chr(34)                                               '21.
        Print #lintFileNum, "A270,915,2,3,1,1,N," & Chr(34) & _
            Format$(.datDespatchDate, "DD/MM/YY,DDD") & Chr(34)                         '22.
        
        Print #lintFileNum, ":Address Info..."                                          '23.
        Print #lintFileNum, "A795,945,2,4,1,1,N," & Chr(34) & Trim$(UCase$(lstrDeliveryName)) & Chr(34) '24.
        Print #lintFileNum, "A795,915,2,4,1,1,N," & Chr(34) & Trim$(UCase$(.strDeliveryAdd1)) & Chr(34)  '25.
        Print #lintFileNum, "A795,885,2,4,1,1,N," & Chr(34) & Trim$(UCase$(.strDeliveryAdd2)) & Chr(34)   '26.
        Print #lintFileNum, "A795,855,2,4,1,1,N," & Chr(34) & Trim$(UCase$(.strDeliveryAdd3)) & Chr(34)    '26.
        Print #lintFileNum, "A795,825,2,4,1,1,N," & Chr(34) & Trim$(UCase$(.strDeliveryAdd4)) & Chr(34)    '26.
        Print #lintFileNum, "A795,795,2,3,2,2,N," & Chr(34) & Trim$(UCase$(.strDeliveryAdd5)) & Chr(34)      '27.
        Print #lintFileNum, "A795,745,2,5,1,1,N," & Chr(34) & Trim$(UCase$(.strDeliveryPostcode)) & Chr(34)    '28.
        
        Print #lintFileNum, ":Parcel Number..."                                         '29.
        Print #lintFileNum, "A210,745,2,4,1,2,N," & Chr(34) & lintArrInc & Chr(34)       '30.
        Print #lintFileNum, "A85,745,2,4,1,2,N," & Chr(34) & .intParcelItems & Chr(34)   '31.
        Print #lintFileNum, "A140,735,2,4,1,1,N," & Chr(34) & "OF" & Chr(34)            '32.
        
        Print #lintFileNum, ":Special Instr..."                                         '33.
        Print #lintFileNum, "LO20,685,800,4"                                            '34.
        Print #lintFileNum, "LO20,625,800,4"                                            '35.
        
        If gstrAdviceNoteOrder.lngConsignRemarkNum <> 0 And Trim$(gstrConsignmentNote.strText) <> "" Then
            lstrConsignmentNoteLine1 = GrapLine(Trim$(.strConsignRemark), 0, 60)
            lstrConsignmentNoteLine2 = GrapLine(Trim$(.strConsignRemark), 1, 60)
            Print #lintFileNum, "A795,680,2,3,1,1,N," & Chr(34) & _
                lstrConsignmentNoteLine1 & Chr(34)                                          '36.
            Print #lintFileNum, "A795,655,2,3,1,1,N," & Chr(34) & _
                lstrConsignmentNoteLine2 & Chr(34)                                          '36.
        End If
        
        Print #lintFileNum, ":POD Bullet..."                                             '38.
        Print #lintFileNum, "A120,480,1,3,1,1,N," & Chr(34) & "LIFT" & Chr(34)          '39.
        Print #lintFileNum, "A180,536,2,3,1,1,R," & Chr(34) & _
        Trim$(gstrPForceServiceInd.strUserDef1) & Chr(34)                                      '40.
        
        
        Print #lintFileNum, "A650,536,2,4,1,1,N," & Chr(34) & _
            "PB " & .strConsignNum & " " & Format$(lintArrInc, "000") & Chr(34)           '41.
        Print #lintFileNum, "A725,606,2,2,1,1,N," & Chr(34) & _
            Trim$(UCase$(lstrDeliveryName)) & Chr(34)                                     '42.
        Print #lintFileNum, "A725,586,2,2,1,1,N," & Chr(34) & _
            Trim$(UCase$(.strDeliveryAdd1)) & Chr(34)                                      '43.
        Print #lintFileNum, "A725,566,2,2,1,1,N," & Chr(34) & _
            Trim$(Trim$(UCase$(.strDeliveryAdd5)) & " " & Trim$(UCase$(.strDeliveryPostcode))) & Chr(34)  '44.
            
        Print #lintFileNum, ":Main Barcode..."                                          '45.
        Print #lintFileNum, "B640,496,2,1,3,10,306,N," & Chr(34) & _
            "PB" & .strConsignNum & Format$(lintArrInc, "000") & Chr(34)                '46.
        Print #lintFileNum, "A630,170,2,3,1,1,N," & Chr(34) & "PB" & Chr(34)            '47.
        Print #lintFileNum, "A610,180,2,3,2,2,N," & Chr(34) & " " & .strConsignNum & Chr(34)    '48.
        Print #lintFileNum, "A255,170,2,3,1,1,N," & Chr(34) & Format$(lintArrInc, "000") & Chr(34)  '49.
        
        With gstrReferenceInfo
            Print #lintFileNum, ":Sender Address..."                                         '50.
            Print #lintFileNum, "A800,130,1,2,1,1,N," & Chr(34) & "From " & Chr(34)         '51.
            Print #lintFileNum, "A800,200,1,2,1,1,N," & Chr(34) & _
                UCase$(.strCompanyName) & ", " & UCase$(.strCompanyAddLine1) & Chr(34) '52.
            Print #lintFileNum, "A780,200,1,2,1,1,N," & Chr(34) & _
                UCase$(.strCompanyAddLine2) & ", " & UCase$(.strCompanyAddLine3) & Chr(34) '53.
            Print #lintFileNum, "A760,200,1,2,1,1,N," & Chr(34) & _
                UCase$(.strCompanyAddLine4) & ", " & UCase$(.strCompanyAddLine5) & Chr(34) '54.
        End With
        
        Print #lintFileNum, ":Sender Ref..."                                            '55.
        Print #lintFileNum, "A60,150,1,2,1,1,N," & Chr(34) & "Sender's:" & Chr(34)      '56.
        Print #lintFileNum, "A40,150,1,2,1,1,N," & Chr(34) & "Ref :" & Chr(34)          '57.
        Print #lintFileNum, "A40,260,1,2,1,1,N," & Chr(34) & _
        Trim$("M" & .lngOrderNum & "/" & .lngCustNum) & Chr(34)                         '58.
        
        Print #lintFileNum, ":Cust Pack Peel Off..."                                     '59.
        Print #lintFileNum, "B640,105,2,1,3,4,60,B," & Chr(34) & _
            "PB" & .strConsignNum & Format$(lintArrInc, "000") & Chr(34)                '60.
        Print #lintFileNum, "A800,70,2,2,1,1,N," & Chr(34) & "Customer" & Chr(34)       '61.
        Print #lintFileNum, "A800,50,2,2,1,1,N," & Chr(34) & "Use Only" & Chr(34)       '62.
        Print #lintFileNum, "A120,70,2,3,2,2,N," & Chr(34) & _
        Trim$(gstrPForceServiceInd.strUserDef1) & Chr(34)                                      '63.
        
        Print #lintFileNum, "P1"                                                        '64.
        Print #lintFileNum, ""
        
    Next lintArrInc
    End With
    
    Close #lintFileNum
End Sub
Sub xPFMain()

    If gbooTestingVarRanAlready = False Then
        gintTestingVarRanAlready = 20
        PopulateStaticItems ' Clear & populate static
        GetPFContractDetails
        GetConsignmentRange
        GetConsignBatchIncrs
        gbooTestingVarRanAlready = True
    End If
    IncrementConsignBatchIncrs
    UpdateIncrs

    With gstrAdviceNoteOrder
        .lngCustNum = 2197
        .lngOrderNum = CLng(gintTestingVarRanAlready)
        .strCourierCode = "PF 48"
        .datDespatchDate = "16/06/1973 12:30.20"
        .strDeliverySalutation = "Mr"
        .strDeliverySurname = "Jones"
        .strDeliveryInitials = "F"
        .strDeliveryAdd1 = "30 Any Street"
        .strDeliveryAdd2 = "The town"
        .strDeliveryAdd3 = ""
        .strDeliveryAdd4 = ""
        .strDeliveryAdd5 = "France"
        .strDeliveryPostcode = "NL32 5KU"
        .intNumOfParcels = 2
        .lngGrossWeight = 300
    End With
    
    PFPCLFile "c:\Thermal.txt"
    
    gintTestingVarRanAlready = gintTestingVarRanAlready + 1
    
End Sub
Sub GetAdviceForPForce(Optional pstrParam As Variant, Optional plngStartOrderNum As Variant, _
    Optional plngEndOrderNum As Variant)
    
Dim lsnaLists As Recordset
Dim lstrSQL As String

    ShowStatus 89
    If IsMissing(pstrParam) Then
        pstrParam = ""
    End If
    
    PopulateStaticItems ' Clear & populate static
    GetPFContractDetails
    GetConsignmentRange
    GetConsignBatchIncrs
        
    On Error GoTo ErrHandler
    
    Select Case pstrParam
    Case ""
        lstrSQL = "SELECT CustNum, OrderNum From " & gtblAdviceNotes & " WHERE " & _
            "(((OrderStatus)='A') AND ((OrderNum)<=" & _
            glngLastOrderPrintedInThisRun & ")  and (val(NumOfParcels & '') > 0));"
                        
    Case gconstrAdviceReportTypeRange
        lstrSQL = "SELECT DeliveryDate, CustNum, OrderNum " & _
            "FROM " & gtblAdviceNotes & " Where "
        lstrSQL = lstrSQL & " (((" & gtblAdviceNotes & ".OrderNum)>=" & plngStartOrderNum & _
            " And (" & gtblAdviceNotes & ".OrderNum)<=" & plngEndOrderNum & _
            ") AND ((" & gtblAdviceNotes & ".OrderStatus)='A') AND ((" & gtblAdviceNotes & ".NumOfParcels)>0)) "
        lstrSQL = lstrSQL & "order by " & gtblAdviceNotes & ".OrderNum;"
    End Select
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        Do Until .EOF
            GetAdviceNote .Fields("CustNum"), .Fields("OrderNum")
            IncrementConsignBatchIncrs
            AddParcelForce .Fields("OrderNum"), .Fields("CustNum")
            .MoveNext
        Loop
    End With
    
    lsnaLists.Close
    UpdateIncrs
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetAdviceForPForce", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Sub GetPForceAwaitings(pstrFilename As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String

    ShowStatus 90
    gstrPForceServiceInd.strListName = "PForce Service Indicator"
    Busy True
    
    On Error GoTo ErrHandler
    lstrSQL = "SELECT OrderNum, CustNum, ConsignNum From " & gtblPForce & " " & _
        "Where (((Status) = 'A')) ORDER BY OrderNum;"
                
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        Do Until .EOF
            GetPForceConsignment .Fields("OrderNum"), .Fields("CustNum"), .Fields("ConsignNum")
            gstrPForceServiceInd.strListCode = gstrPForceConsignment.strServiceID
            GetListVarsAll gstrPForceServiceInd
            PFPCLFile pstrFilename

            .MoveNext
        Loop
    End With
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetPForceAwaitings", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub

Sub PForceElectronicFileHeader(pstrFilename As String)
Dim lintFileNum As Integer
Const FC = "+"

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    
    With gstrFileHeader
        Print #lintFileNum, .strRecordTypeInd & FC & .strFileVersion & FC & _
            .strFileType & FC & _
            Trim$(gstrPFContractDetails.strHalconCustumerCode) & FC & _
            Trim$(gstrPFContractDetails.strContractNumber) & FC & _
            Trim$(gstrPForceConsignment.strBatchNumber) & FC & _
            Format$(gstrPForceConsignment.datDespatchDate, "YYYYMMDD") & FC & _
            Format$(gstrPForceConsignment.datDespatchDate, "HHMMSS") & FC
            'Here Despatch time/date taken from first parcel record, as the file will be sent
            'Daily and there will only be of PF pickup a day.
    End With

    With gstrSenderRecord
        Print #lintFileNum, .strRecordTypeInd & FC & .strFileVersion & FC & _
        Trim$(.strSenderName) & FC & _
        Trim$(.strSenderAdd1) & FC & _
        Trim$(.strSenderAdd2) & FC & _
        Trim$(.strSenderAdd3) & FC & FC & FC & FC & _
        Trim$(.strSenderPostCode) & FC
    End With
        
    Close #lintFileNum
    
End Sub
Function PForceManifestHeader(pstrFilename As String, plngPageNum As Long) As Long
Dim lintFileNum As Integer
Dim llngLineNum As Long

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    
    Print #lintFileNum, "Date Despatched : " & Format$(Now(), "DD/MM/YYYY") & Spacer("", 34) & "Page  " & plngPageNum: llngLineNum = llngLineNum + 1
    Print #lintFileNum, Spacer("", 28) & "PARCELFORCE MANIFEST": llngLineNum = llngLineNum + 1
    With gstrReferenceInfo 
        Print #lintFileNum, Spacer("", 7) & UCase$(Trim$(.strCompanyName) & ", " & .strCompanyAddLine1 & "," & .strCompanyAddLine2 & ", " & .strCompanyAddLine3 & ", " & .strCompanyAddLine4 & ", " & .strCompanyAddLine5): llngLineNum = llngLineNum + 1
    End With
    
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, Spacer("", 21) & Trim$(gstrPForceServiceInd.strDescription) & " Manifest": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "Contract No. " & gstrPFContractDetails.strContractNumber & _
        " Service " & gstrPForceServiceInd.strUserDef1: llngLineNum = llngLineNum + 1
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "Consignment                                       Senders         SpeHand": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "Number    Delivery Name                  Postcode Reference    Itms S B P": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "--------- ------------------------------ -------- -------------- -- - - -": llngLineNum = llngLineNum + 1
        
    Close #lintFileNum
    PForceManifestHeader = llngLineNum
    
End Function
Function PForceElectronicFileDetail(pstrFilename As String) As Long
Dim lintFileNum As Integer
Const FC = "+"
Dim lstrDeliveryName As String

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    
    PForceElectronicFileDetail = 1
    
    lstrDeliveryName = Trim$(Trim$(gstrPForceConsignment.strDeliverySalutation) & " " & _
        Trim$(gstrPForceConsignment.strDeliveryInitials) & " " & _
        Trim$(gstrPForceConsignment.strDeliverySurname))
        
    With gstrDetailRecord
        Print #lintFileNum, .strRecordTypeInd & FC & .strFileVersion & FC & _
        Trim$(gstrPForceConsignment.strConsignNum) & FC & _
        Trim$(gstrPForceConsignment.strServiceID) & FC & _
        Trim$(.strWeekendHand) & FC & _
        Trim$(.strPostCodeKeyed) & FC & _
        Trim$(.strPostTownDerived) & FC & _
        Trim$("M" & gstrPForceConsignment.lngOrderNum & "/" & gstrPForceConsignment.lngCustNum) & FC & _
        Trim$(gstrPFContractDetails.strHalconCollectionID) & FC & _
        Trim$(gstrPFContractDetails.strContractNumber) & FC & _
        Trim$(gstrPForceConsignment.lngGrossWeight) & FC & _
        Trim$(gstrPForceConsignment.intParcelItems) & FC & _
        gstrPForceConsignment.strPrepaidInd & FC & _
        Trim$(lstrDeliveryName) & FC & _
        Trim$(gstrPForceConsignment.strDeliveryAdd1) & FC & _
        Trim$(gstrPForceConsignment.strDeliveryAdd2) & FC & _
        Trim$(gstrPForceConsignment.strDeliveryAdd3) & FC & _
        Trim$(gstrPForceConsignment.strDeliveryAdd4) & FC & _
        Trim$(gstrPForceConsignment.strDeliveryPostcode) & FC
    End With
    
    If Trim$(gstrPForceConsignment.strNotificationCode) <> "" Then
        PForceElectronicFileDetail = 2
        With gstrDetailSupplementRecord
            Print #lintFileNum, .strRecordTypeInd & FC & .strFileVersion & FC & _
            Trim$(gstrPForceConsignment.strConsignNum) & FC & _
            Trim$(gstrPForceConsignment.strNotificationCode) & FC & _
            Trim$(gstrPForceConsignment.strConsignRemark) & FC
        End With
    End If
    Close #lintFileNum
    
        
End Function
Function PForceManifestDetail(pstrFilename As String) As Long
Dim lintFileNum As Integer
Dim lstrDeliveryName As String
Dim lstrGrossWeight As String

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    With gstrPForceConsignment
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & _
            Trim$(.strDeliveryInitials) & " " & _
            Trim$(.strDeliverySurname))
        
        Print #lintFileNum, Spacer(Trim$(.strConsignNum), 9) & " " & _
            Spacer(lstrDeliveryName, 30) & " " & Spacer(Trim$(.strDeliveryPostcode), 8) & _
            " " & Spacer(Trim$("M" & .lngOrderNum & "/" & .lngCustNum), 14) & " " & _
                Spacer(Trim$(.intParcelItems), 2, "L") & " " & _
                Left$(Trim$(.strSpecialSatDel), 1) & " " & _
                Left$(Trim$(.strSpecialBookIn), 1) & " " & Left$(Trim$(.strSpecialProof), 1)
            PForceManifestDetail = .intParcelItems
        End With
    Close #lintFileNum
    
End Function

Sub PForceElectronicFileFooter(pstrFilename As String, plngRecordCount As Long)
Dim lintFileNum As Integer
Const FC = "+"

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
        
    With gstrTrailerRecord
        Print #lintFileNum, .strRecordTypeInd & FC & .strFileVersion & FC & _
        Trim$(plngRecordCount + 1) & FC
    End With
    
    
    Close #lintFileNum
    
End Sub
Function PForceManifestFooter(pstrFilename As String, plngItemCount As Long, plngRecordCount As Long) As Long
Dim lintFileNum As Integer
Dim llngLineNum As Long

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
        
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, "": llngLineNum = llngLineNum + 1
    Print #lintFileNum, Spacer("", 16) & "Number of Consignments" & Spacer("", 10) & plngRecordCount: llngLineNum = llngLineNum + 1
    Print #lintFileNum, Spacer("", 16) & "Number of Items" & Spacer("", 17) & plngItemCount: llngLineNum = llngLineNum + 1
    
    Close #lintFileNum
    
    PForceManifestFooter = llngLineNum
    
End Function
Sub PForceManifestSpacer(pstrFilename As String, pintNumOfLines As Integer)
Dim lintFileNum As Integer
Dim lintArrInc As Integer

    On Error Resume Next
    
    lintFileNum = FreeFile
    
    Open pstrFilename For Append As lintFileNum
    
    For lintArrInc = 1 To pintNumOfLines
        Print #lintFileNum, ""
    Next lintArrInc
    
    Close #lintFileNum
    
End Sub
Function PForceFileGeneral(pstrFilename As String) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim llngDetailRecordCount As Long

    PForceFileGeneral = True 
    
    ShowStatus 91

    On Error GoTo ErrHandler
    lstrSQL = "SELECT OrderNum, CustNum, ConsignNum, Status From " & gtblPForce & " WHERE (((Status)='P'));"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            GetPForceConsignment .Fields("OrderNum"), .Fields("CustNum"), .Fields("ConsignNum")
    
            If llngRecCount = 1 Then
                PopulateStaticItems ' Clear & populate static
                GetPFContractDetails
                PForceElectronicFileHeader pstrFilename 'Assume 2 records
            End If
            
            llngDetailRecordCount = llngDetailRecordCount + PForceElectronicFileDetail(pstrFilename)

            .MoveNext
        Loop
    End With
    
    If llngRecCount <> 0 Then
        PForceElectronicFileFooter pstrFilename, llngDetailRecordCount + 2
    Else
        MsgBox "No Consignment found!", , gconstrTitlPrefix & "General PF File Production"
        PForceFileGeneral = False
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PForceFileGeneral", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function

Sub PopulateStaticItems()

    ShowStatus 92
    
    With gstrParcelNumber
        .strAlphaPrefix = "" 'not PB, this is added before
        .strNumber6Dig = ""
        .strChkDigit = ""
    End With
    
    With gstrThermalLabel
        .lngCustNum = 0
        .lngOrderNum = 0
        .strSalutation = ""
        .strSurname = ""
        .strInitials = ""
        .strAdd1 = ""
        .strAdd2 = ""
        .strAdd3 = ""
        .strAdd4 = ""
        .strAdd5 = ""
        .strPostcode = ""
        .intNumberOfParcels = 0
    End With
    
    With gstrFileHeader
        .strRecordTypeInd = "0"
        .strFileVersion = "02"
        .strFileType = "SKEL"  'MANF, SKEL, COLL
    End With
    
    With gstrSenderRecord
        .strRecordTypeInd = "1"
        .strFileVersion = "02"
        .strSenderName = UCase$(gstrReferenceInfo.strCompanyName) 
        .strSenderAdd1 = UCase$(gstrReferenceInfo.strCompanyAddLine1)
        .strSenderAdd2 = UCase$(gstrReferenceInfo.strCompanyAddLine2)
        .strSenderAdd3 = UCase$(gstrReferenceInfo.strCompanyAddLine3)       
        .strSenderPostCode = UCase$(gstrReferenceInfo.strCompanyAddLine4)
    End With
    
    With gstrDetailRecord
        .strRecordTypeInd = "2"
        .strFileVersion = "02"
        .strWeekendHand = ""
        .strPostCodeKeyed = "" 'Unused
        .strPostTownDerived = ""  'Unused
    End With
    
    With gstrDetailSupplementRecord
        .strRecordTypeInd = "3"
        .strFileVersion = "02"
    End With
    
    With gstrTrailerRecord
        .strRecordTypeInd = "9"
        .strFileVersion = "02"
    End With
    
End Sub
Sub AddParcelForce(plngOrderNum As Long, plngCustNum As Long)
Dim lstrSQL As String
Dim pdatPrintedDate As Date
Dim pstrPrintedBy As String
Dim ldatDespatchDate As Date
Dim lstrDeliveryName As String
Dim lstrName As String

    ShowStatus 84
    On Error GoTo ErrHandler

    lstrSQL = " INSERT INTO " & gtblPForce & " ( CustNum, OrderNum, Status, ConsignNum, ServiceID, " & _
        "BatchNumber, DespatchDate, DeliverySalutation, DeliverySurname, DeliveryInitials, " & _
        "DeliveryAdd1, DeliveryAdd2, DeliveryAdd3, DeliveryAdd4, DeliveryAdd5, " & _
        "DeliveryPostcode, ParcelItems, GrossWeight, WeekendHandCode, PrepaidInd, " & _
        "NotificationCode, ConsignRemark, SpecialSatDel, SpecialBookIn, SpecialProof ) Select "
    With gstrAdviceNoteOrder
        ldatDespatchDate = CDate(Now())
        
        lstrSQL = lstrSQL & .lngCustNum & ", " & .lngOrderNum & ", 'A','" & _
            gstrParcelNumber.strAlphaPrefix & gstrParcelNumber.strNumber6Dig & _
            gstrParcelNumber.strChkDigit & "','"
        Select Case UCase$(Trim$(.strCourierCode))
        Case "PF 48", "", "PF"
            lstrSQL = lstrSQL & "SUP"
        Case Else
            lstrSQL = lstrSQL & Trim$(.strCourierCode)
        End Select
        lstrSQL = lstrSQL & "','" & gstrBatchIncrNum & "', #" & Format$(ldatDespatchDate, "DD/MMM/YYYY") & "#,'"
        
        lstrName = Trim$(Trim$(.strSalutation) & " " & _
            Trim$(.strInitials) & " " & Trim$(.strSurname))
        
        lstrDeliveryName = Trim$(Trim$(.strDeliverySalutation) & " " & _
            Trim$(.strDeliveryInitials) & " " & Trim$(.strDeliverySurname))
        
        If Trim$(lstrDeliveryName) = "" Then
            lstrSQL = lstrSQL & OneSpace(JetSQLFixup(.strSalutation)) & "','" & OneSpace(JetSQLFixup(.strDeliverySurname)) & "','" & _
            OneSpace(JetSQLFixup(.strDeliveryInitials))
        Else
            lstrSQL = lstrSQL & OneSpace(JetSQLFixup(.strDeliverySalutation)) & "','" & OneSpace(JetSQLFixup(.strDeliverySurname)) & "','" & _
            OneSpace(JetSQLFixup(.strDeliveryInitials))
        End If
        
        lstrSQL = lstrSQL & "','" & OneSpace(JetSQLFixup(.strDeliveryAdd1)) & "','" & _
        OneSpace(JetSQLFixup(.strDeliveryAdd2)) & "','" & OneSpace(JetSQLFixup(.strDeliveryAdd3)) & "','" & _
        OneSpace(JetSQLFixup(.strDeliveryAdd4)) & "','" & OneSpace(JetSQLFixup(.strDeliveryAdd5)) & "','" & _
        OneSpace(JetSQLFixup(.strDeliveryPostcode)) & "'," & .intNumOfParcels & "," & .lngGrossWeight & _
        ",' ',' ','"
        
        If .lngConsignRemarkNum > 0 Then
            GetRemark .lngConsignRemarkNum, gstrConsignmentNote
            lstrSQL = lstrSQL & "IRD1','" & Left$(Trim$(gstrConsignmentNote.strText), 100) & "'"
        Else
            lstrSQL = lstrSQL & " ',' '"
        End If
        lstrSQL = lstrSQL & ",'N','N','N';"
        
    End With

    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddParcelForce", "Central", True)
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
Sub UpdateAdviceParcelExtras(plngCustNum As Long, plngOrderNum As Long, pintParcelCount As Integer, plngGrossWeight As Long)
Dim lstrSQL As String

    ShowStatus 93
    
    On Error GoTo ErrHandler
    
    lstrSQL = "UPDATE " & gtblAdviceNotes & " SET NumOfParcels = " & pintParcelCount & _
        ", GrossWeight = " & plngGrossWeight & " WHERE (((CustNum)=" & plngCustNum & _
        ") AND ((OrderNum)=" & plngOrderNum & "));"
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateAdviceParcelExtras", "Central", True)
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
Sub UpdateIncrs()
Dim lstrSQL As String

    ShowStatus 94
    
    On Error GoTo ErrHandler
    lstrSQL = "UPDATE " & gtblSystem & " SET [Value] = '" & gstrParcelNumber.strNumber6Dig & _
    "' WHERE (((Item)='LastPFConsignNumIncr'));"
    gdatCentralDatabase.Execute lstrSQL

    lstrSQL = "UPDATE " & gtblSystem & " SET [Value] = '" & gstrBatchIncrNum & "' WHERE " & _
        "(((Item)='BatchIncr'));"
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateIncrs", "Central", True)
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
Sub UpdatePForceStatus(pstrStatus As String, pstrWhereParam As String)
Dim lstrSQL As String

    ShowStatus 95
    
    On Error GoTo ErrHandler
    lstrSQL = "UPDATE " & gtblPForce & " SET Status = '" & pstrStatus & "' WHERE (((Status)='" & pstrWhereParam & "'));"
    gdatCentralDatabase.Execute lstrSQL


Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdatePForceStatus", "Central", True)
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
Sub GetPForceConsignment(plngOrderNum As Long, plngCustNum As Long, pstrConsignNum As String)
Dim ltabPFConsign As Recordset

    ShowStatus 96
    On Error GoTo ErrHandler
    Set ltabPFConsign = gdatCentralDatabase.OpenRecordset(gtblPForce)
    ltabPFConsign.Index = "OrdCust"
    
    With ltabPFConsign
        .Seek "=", plngCustNum, plngOrderNum, pstrConsignNum
        
        If Not .NoMatch Then
            gstrPForceConsignment.lngCustNum = .Fields("CustNum")
            gstrPForceConsignment.lngOrderNum = .Fields("OrderNum")
            gstrPForceConsignment.strStatus = .Fields("Status")
            gstrPForceConsignment.strConsignNum = .Fields("ConsignNum")
            gstrPForceConsignment.strServiceID = .Fields("ServiceID")
            gstrPForceConsignment.strBatchNumber = .Fields("BatchNumber")
            gstrPForceConsignment.datDespatchDate = .Fields("DespatchDate")
            gstrPForceConsignment.strDeliverySalutation = .Fields("DeliverySalutation")
            gstrPForceConsignment.strDeliverySurname = .Fields("DeliverySurname")
            gstrPForceConsignment.strDeliveryInitials = .Fields("DeliveryInitials")
            gstrPForceConsignment.strDeliveryAdd1 = .Fields("DeliveryAdd1")
            gstrPForceConsignment.strDeliveryAdd2 = .Fields("DeliveryAdd2")
            gstrPForceConsignment.strDeliveryAdd3 = .Fields("DeliveryAdd3")
            gstrPForceConsignment.strDeliveryAdd4 = .Fields("DeliveryAdd4")
            gstrPForceConsignment.strDeliveryAdd5 = .Fields("DeliveryAdd5")
            gstrPForceConsignment.strDeliveryPostcode = .Fields("DeliveryPostcode")
            gstrPForceConsignment.intParcelItems = .Fields("ParcelItems")
            gstrPForceConsignment.lngGrossWeight = .Fields("GrossWeight")
            gstrPForceConsignment.strWeekendHandCode = .Fields("WeekendHandCode")
            gstrPForceConsignment.strPrepaidInd = .Fields("PrepaidInd")
            gstrPForceConsignment.strNotificationCode = .Fields("NotificationCode")
            gstrPForceConsignment.strConsignRemark = .Fields("ConsignRemark")
            gstrPForceConsignment.strSpecialSatDel = .Fields("SpecialSatDel")
            gstrPForceConsignment.strSpecialBookIn = .Fields("SpecialBookIn")
            gstrPForceConsignment.strSpecialProof = .Fields("SpecialProof")

        End If
        .Close

    End With

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetPForceConsignment", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function GetPFContractDetails()
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    ShowStatus 97
    On Error GoTo ErrHandler
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".* " & _
        "FROM " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & gtblListDetails & ".ListNum = " & _
        "" & gtblLists & ".ListNum WHERE " & gtblLists & ".ListName='PForce Contract Details' and " & _
        "" & gtblListDetails & ".InUse = True;"

    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1

            Select Case Trim$(.Fields("ListCode"))
            Case "ContractNo"
                gstrPFContractDetails.strContractNumber = Left$(.Fields("Description"), 7)
            Case "AcctNo" 'Not used anywhere in Program
                gstrPFContractDetails.strPFAcctNumber = Trim$(.Fields("Description"))
            Case "HalconCoID"
                gstrPFContractDetails.strHalconCollectionID = Left$(.Fields("Description"), 6)
            Case "HCustCode"
                gstrPFContractDetails.strHalconCustumerCode = Left$(.Fields("Description"), 8)
            End Select
            
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CreateNextConsignmentNum", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function


Sub PForceManifestGeneral(pstrFilename As String)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim llngRecCountOnThisPage As Long
Dim llngItemsCount As Long
Dim lstrLastServiceIDGrouping As String
Dim lbooJustPrintedHeader As Boolean
Dim llngPageNumber As Long
Dim llngLinesInHeader As Long
Dim llngLinesInFooter As Long
Dim llngLinesOnPage As Long
Dim lintArrInc As Integer

    ShowStatus 98
    
    Busy True
    On Error GoTo ErrHandler
    lstrSQL = "SELECT OrderNum, CustNum, Status, ServiceID, ConsignNum, ParcelItems From " & _
        "" & gtblPForce & " WHERE (((Status)='P')) ORDER BY ServiceID, ConsignNum;"
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        llngRecCount = 0
        llngPageNumber = 1
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            llngRecCountOnThisPage = llngRecCountOnThisPage + 1
            lbooJustPrintedHeader = False
            
            If llngRecCount = 1 Then
                PopulateStaticItems ' Clear & populate static
                GetPFContractDetails
            End If
            
            GetPForceConsignment .Fields("OrderNum"), .Fields("CustNum"), .Fields("ConsignNum")
            
            If lstrLastServiceIDGrouping = "" Or lstrLastServiceIDGrouping <> .Fields("ServiceID") Then
                gstrPForceServiceInd.strListName = "PForce Service Indicator"
                gstrPForceServiceInd.strListCode = gstrPForceConsignment.strServiceID
                GetListVarsAll gstrPForceServiceInd
                llngLinesInHeader = PForceManifestHeader(pstrFilename, llngPageNumber)
                llngRecCountOnThisPage = 1
                llngItemsCount = 0
            End If
                        
            llngItemsCount = llngItemsCount + PForceManifestDetail(pstrFilename)
            llngLinesOnPage = llngLinesInHeader + llngRecCountOnThisPage ' + 5
            If llngLinesOnPage + Val(.Fields("ParcelItems")) >= (llngPageNumber * gintNumberOfLinesAPage) Then
                
                PForceManifestSpacer pstrFilename, ((llngPageNumber * gintNumberOfLinesAPage) - llngLinesOnPage)
                llngPageNumber = llngPageNumber + 1
                llngLinesInHeader = PForceManifestHeader(pstrFilename, llngPageNumber)
            End If
                        
            If lstrLastServiceIDGrouping <> .Fields("ServiceID") And lstrLastServiceIDGrouping <> "" Then
                llngLinesInFooter = PForceManifestFooter(pstrFilename, llngItemsCount, llngRecCountOnThisPage)
                llngPageNumber = llngPageNumber + 1
                lbooJustPrintedHeader = True
                llngLinesOnPage = llngLinesInHeader + llngItemsCount + llngLinesInFooter
                
                PForceManifestSpacer pstrFilename, ((llngPageNumber * gintNumberOfLinesAPage) - llngLinesOnPage)
            
            End If
            lstrLastServiceIDGrouping = .Fields("ServiceID")

            .MoveNext
        Loop
    End With
    
    If llngRecCount <> 0 Then
        If lbooJustPrintedHeader = False Then
                PForceManifestFooter pstrFilename, llngItemsCount, llngRecCountOnThisPage
        End If
    Else
        MsgBox "No Consignment found!", , gconstrTitlPrefix & "General Manifest Production"
    End If
    
    lsnaLists.Close
     Busy False
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PForceFileGeneral", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
