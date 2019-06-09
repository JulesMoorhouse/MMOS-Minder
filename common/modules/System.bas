Attribute VB_Name = "modSystem"
Option Explicit

Global glngLastOrderPrintedInThisRun As Long

Global gdatCentralDatabase      As Database
Global gdatLocalDatabase        As Database
Global gdatStockImportDB        As Database

Global gwrkODBC As Workspace
Global gwrkJet As Workspace
    
Global Const gconstrCashbookSpecificCustomer = "SPECIFICCUSTOMER"

Global Const gconstrAdviceReportTypeRange = "RANGE"

Global gstrLPTPortNumber As String

Public Const gstrQUOTE$ = """"

Function Spacer(pstrString, pintLength As Integer, Optional pstrParam As Variant) As String
Dim lintWordLength As Integer

    
    'Next line would never happen! 13/07/01
    
    If IsMissing(pstrParam) Then
        pstrParam = ""
    End If
    
    If Len(pstrString) > pintLength Then
        pstrString = Left$(pstrString, pintLength)
    End If
    lintWordLength = Len(Trim$(pstrString))
    Select Case pstrParam
    Case "L"
        Spacer = Space(pintLength - lintWordLength) & Trim$(pstrString)
    Case Else
        Spacer = Trim$(pstrString) & Space(pintLength - lintWordLength)
    End Select

End Function
Sub UpdateUser(pstrUserID As String, pstrUserFullName As String, pstrPassword As String, plngUserLevel As Long, pstrNotes As String)
Dim lstrSQL As String
    
    ShowStatus 79
    On Error GoTo ErrHandler
    
    pstrPassword = Hash(pstrPassword)
    
    lstrSQL = "UPDATE " & gtblUsers & " SET UserPassword = '" & pstrPassword & _
        "', UserName = '" & pstrUserFullName & _
        "', UserLevel = " & plngUserLevel & _
        ", UserNotes = '" & pstrNotes & _
        "' WHERE (((UserID)='" & _
        pstrUserID & "'));"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateUser", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        'LastChanceCafe
    Case Else
        Resume Next
    End Select

End Sub
Function GetUser(pstrUserName As String, ByRef pintUserCount As Long, pbooNoStatus As String) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
'Converted table names to constants

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 81
    End If
    
    On Error GoTo ErrHandler
    
    'lstrSQL = "SELECT * from " & gtblUsers & " where UserId ='" & pstrUserName & "';"
    
    lstrSQL = "SELECT First(u1.UserID) AS UserID, " & _
        "First(u1.UserName) AS UserName, " & _
        "First(u1.UserPassword) AS UserPassword, " & _
        "First(u1.UserLevel) AS UserLevel, " & _
        "First(u1.UserNotes) AS UserNotes, " & _
        "Count(u2.UserID) AS Count " & _
        "FROM " & gtblUsers & " AS u1, " & gtblUsers & " AS u2 " & _
        "where u1.userid = '" & pstrUserName & "';"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            gstrGenSysInfo.strUserName = .Fields("UserID") & "" ' & "" for these 5 fields
            gstrGenSysInfo.strUserFullName = .Fields("UserName") & ""
            gstrGenSysInfo.strUserPassword = .Fields("UserPassword") & ""
            Dim lstrLevel As String: lstrLevel = .Fields("UserLevel") & ""
            If lstrLevel = "" Then
                lstrLevel = 0
            End If
            gstrGenSysInfo.lngUserLevel = CLng(lstrLevel)
            gstrGenSysInfo.strUserNotes = .Fields("UserNotes") & ""
            pintUserCount = .Fields("Count") & ""
        End If
    End With
    
    If pintUserCount = 0 Then
        GetUser = False
    Else
        GetUser = True
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetUser", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim$(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
           
            ' It's quoted.  Get rid of the quotes.
           
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
    
End Function
Function FindProgram(pstrProg As String) As String
Dim lstrProgToFind As String
    
    Select Case pstrProg
    Case "MSACCESS"
        lstrProgToFind = "C:\Program Files\Microsoft Office\Office\MSACCESS.EXE"
    Case "IEXPLORE"
        lstrProgToFind = "C:\Program Files\Internet Explorer\IEXPLORE.EXE"
    Case "QARAPID"
        If gstrSystemRoute = srStandardRoute Then pstrProg = "AddressProg"
        lstrProgToFind = "C:\QADDRESS\APPS\Qarapid.EXE"
    End Select
    
    If Dir(lstrProgToFind) <> "" Then
        FindProgram = lstrProgToFind
        SaveSetting gstrIniAppName, pstrProg, "Location", FindProgram
    
    Else
        If Dir(GetSetting(gstrIniAppName, pstrProg, "Location")) = "" Or _
        UCase$(Dir(GetSetting(gstrIniAppName, pstrProg, "Location"))) <> pstrProg & ".EXE" Then
            SaveSetting gstrIniAppName, pstrProg, "Location", _
                InputBox("Please specify the exact location of " & pstrProg & ".EXE," & _
                vbCrLf & "You may need help from you Technical support office with this!", _
                "Need to find the location of " & pstrProg, _
                lstrProgToFind)
            FindProgram = GetSetting(gstrIniAppName, pstrProg, "Location")
        End If
    End If

End Function

Function EstablishQuickAddress() As Boolean
Dim lstrQAExe As String
Dim lstrShellString As String
Dim lstrParams As String
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim intReturnValue As Integer


    If gstrSystemRoute = srStandardRoute Then
        lstrQAExe = FindProgram("AddressProg")
    ElseIf gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        lstrQAExe = FindProgram("QARAPID")
    End If
    
    ShowStatus 83
    On Error GoTo ErrorHandler
    
    If Trim$(lstrQAExe) = "" Or Dir(lstrQAExe) = "" Then
        MsgBox "You have not got Quick Address properly installed, you will not be able to use Quick Address!"
        EstablishQuickAddress = False
        Exit Function
    End If
    
    '-ini "Configuration file name"
    '-section "Section name"
    If gstrSystemRoute = srStandardRoute Then
        FileCopy gstrStatic.strServerPath & "QAMMOS.ini", Justpath(lstrQAExe) & "QAMMOS.INI"
        lstrParams = " -ini " & Justpath(lstrQAExe) & "QAMMOS.ini" & " -section MMOS"
    ElseIf gstrSystemRoute = srCompanyRoute Or gstrSystemRoute = srCompanyDebugRoute Then
        FileCopy gstrStatic.strServerPath & "QATMOS.ini", Justpath(lstrQAExe) & "QATMOS.INI"
        lstrParams = " -ini " & Justpath(lstrQAExe) & "QATMOS.ini" & " -section TMOS"
    End If
    
    lstrShellString = strUnQuoteString(lstrQAExe & lstrParams)
        
    start.cb = Len(start)

    ' Start the shelled application:
    intReturnValue = CreateProcessA(0&, lstrShellString, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
    EstablishQuickAddress = True
    Exit Function
ErrorHandler:
    EstablishQuickAddress = False
    Exit Function
    
End Function
Sub Wait(pintSeconds As Integer)
'Dim ldatCurrentTime As Date

    'ldatCurrentTime = DateAdd("S", pintSeconds, Now())
    'Debug.Print Now()
    'Do Until Now() >= ldatCurrentTime
    '    'wait
    'Loop
    'Debug.Print Now()

    Sleep pintSeconds * 1000
    
End Sub
Sub AddressFill(pstrCriteria As String)


    'EstablishQuickAddress
    'DoEvents
    'Shell "QuickAddress Rapid"
    On Error GoTo ErrHandler
    AppActivate "QuickAddress Rapid"
    
    Exit Sub
    
ErrHandler:
    Select Case Err.Number
    Case 5
        MsgBox "Quick Address could not be launched!" & vbCrLf & vbCrLf & _
            "Please ensure QA is installed and working" & vbCrLf & _
            "working correctly", vbInformation, "Error From - AddressFill"
    End Select
    
    'Wait 1
    'SendKeys "{ENTER}"
    'SendKeys "" & pstrCriteria & "{ENTER}"

End Sub
Function Justpath(pstrProgramExe As String) As String
Dim intRevLocSlash As Integer

    intRevLocSlash = InStr(Reverse(pstrProgramExe), "\")
    Justpath = Left$(pstrProgramExe, Len(pstrProgramExe) - intRevLocSlash) & "\"

End Function
Sub CheckForMessages()
Dim lstrMessID As String
Dim lstrMessage As String
Dim lstrMessageRead As String

    lstrMessID = Left$(Format(Now(), "DDMMMYYYYHHMM"), 12)
    
    'NEXT LINE FOR TESTING PURPOSES
    ''SetPrivateINI ServerPath & "Messages\TMOS.ini", "ALL USERS", lstrMessID, "Test"
    lstrMessageRead = ""
    lstrMessage = GetPrivateINI(gstrStatic.strServerPath & "Messages\TMOS.ini", "ALL USERS", lstrMessID)
    
    If Trim$(lstrMessage) <> "" Then
        lstrMessageRead = GetPrivateINI(gstrStatic.strServerPath & "Messages\READSYS.ini", _
            UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID)
        If lstrMessageRead <> "YES" Then
            MsgBox lstrMessage, , "SYSTEM WIDE MESSAGE FROM IT"
            SetPrivateINI gstrStatic.strServerPath & "Messages\READSYS.ini", _
            UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID, "YES"
        End If
    End If
        
    'NEXT LINE FOR TESTING PURPOSES
    ''lstrMessID = Left$(Format(Now(), "DDMMMYYYYHHMM"), 12)
    
    'NEXT LINE FOR TESTING PURPOSES
    ''SetPrivateINI ServerPath & "Messages\TMOS.ini", UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID, "Test"
    lstrMessageRead = ""
    
'    lstrMessage = GetPrivateINI(ServerPath & "Messages\TMOS.ini", UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID)
    lstrMessage = GetPrivateINI(gstrStatic.strServerPath & "Messages\TMOS.INI", UCase$(Trim$(CurrentMachineName)), lstrMessID)
    If Trim$(lstrMessage) <> "" Then
        lstrMessageRead = GetPrivateINI(gstrStatic.strServerPath & "Messages\READPERS.ini", _
            UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID)
        If lstrMessageRead <> "YES" Then
            MsgBox lstrMessage, , "PRIVATE SYSTEM MESSAGE FROM IT"
            SetPrivateINI gstrStatic.strServerPath & "Messages\READPERS.ini", _
            UCase$(Trim$(gstrGenSysInfo.strUserName)), lstrMessID, "YES"
        End If
    End If
    
    
End Sub
Function FormatMessageText(pstrMessagetext) As String
Dim lintReturnChar As Integer
Dim lintArrInc As Integer
    
    lintReturnChar = 1

    Do While lintReturnChar <> 0
        lintReturnChar = InStr(1, pstrMessagetext, Chr(10))
        If lintReturnChar <> 0 Then
            Mid(pstrMessagetext, lintReturnChar - 1, 1) = " "
            Mid(pstrMessagetext, lintReturnChar, 1) = vbCrLf
        End If
    Loop
    
    FormatMessageText = pstrMessagetext
    
End Function
Sub BatchFile(pstrFilename As String)
Dim lintFileNum As Integer

    
    Busy True ', Me
    lintFileNum = FreeFile
    Open pstrFilename & ".bat" For Append As lintFileNum
    
    Print #lintFileNum, "@cd\"
    Print #lintFileNum, "type " & Chr(34) & pstrFilename & _
        ".tmp" & Chr(34) & " > " & gstrLPTPortNumber & ":"
    'Print #lintFileNum, "pause"
    
    Close #lintFileNum
    
    Busy False ', Me
    
    'If gstrGenSysInfo.lngUserLevel = 99 Or _
    '    gstrGenSysInfo.lngUserLevel = 40 Then
    '    On Error Resume Next
    '    Shell "NOTEPAD " & pstrFileName & ".tmp", vbMinimizedNoFocus
    'End If
    
    RunNWait pstrFilename & ".bat"
    DoEvents
    On Error Resume Next
    Kill pstrFilename & ".bat"
    
End Sub
Sub LastChanceCafe()
Dim lstrString As String
Dim lintFileNum As Integer

    GoTo Quit
    lstrString = "This information has been generated as an Unexpected error has caused the program to end." & vbCrLf & _
        "Please print this file and report to incident to your Technical Support office." & vbCrLf & vbCrLf

    With gstrAdviceNoteOrder
        lstrString = lstrString & vbCrLf & "Customer Number = " & .lngCustNum
        lstrString = lstrString & vbCrLf & "Order Num = " & .lngOrderNum
        lstrString = lstrString & vbCrLf & "Media Code = " & .strMediaCode
        lstrString = lstrString & vbCrLf & "Delivery Date = " & .datDeliveryDate
        lstrString = lstrString & vbCrLf & "Courier Code = " & .strCourierCode
        lstrString = lstrString & vbCrLf & "Payment Method1 = " & .strPaymentType1
        lstrString = lstrString & vbCrLf & "Payment Method2 = " & .strPaymentType2
        lstrString = lstrString & vbCrLf & "Orcer Code = " & .strOrderCode
        lstrString = lstrString & vbCrLf & "Card Number =" & .strCardNumber
        lstrString = lstrString & vbCrLf & "Expiry Date =" & .datExpiryDate
        lstrString = lstrString & vbCrLf & "Donation = " & .strDonation
        lstrString = lstrString & vbCrLf & "Payment1 = " & .strPayment
        lstrString = lstrString & vbCrLf & "Payment2 = " & .strPayment2
        lstrString = lstrString & vbCrLf & "Underpayment = " & .strUnderpayment
        lstrString = lstrString & vbCrLf & "Reconcilliation = " & .strReconcilliation
        
        lstrString = lstrString & vbCrLf & "Postage = " & .strPostage
        lstrString = lstrString & vbCrLf & "TotalIncVat = " & .strTotalIncVat
        lstrString = lstrString & vbCrLf & "VAT = " & .strVAT
    End With


        lstrString = lstrString & vbCrLf

    With gstrCustomerAccount
        .lngCustNum = 0
        lstrString = lstrString & vbCrLf & "Surname = " & .strSurname
        lstrString = lstrString & vbCrLf & "Initals = " & .strInitials
        lstrString = lstrString & vbCrLf & "Address 1 = " & .strAdd1
        lstrString = lstrString & vbCrLf & "Address 2 = " & .strAdd2
        lstrString = lstrString & vbCrLf & "Address 3 = " & .strAdd3
        lstrString = lstrString & vbCrLf & "Address 4 = " & .strAdd4
        lstrString = lstrString & vbCrLf & "Address 5 = " & .strAdd5
        lstrString = lstrString & vbCrLf & "Postcode = " & .strPostcode
        lstrString = lstrString & vbCrLf & "Telephone = " & .strTelephoneNum
        lstrString = lstrString & vbCrLf & "Delivery Address 1 = " & .strDeliveryAdd1
        lstrString = lstrString & vbCrLf & "Delivery Address 2 = " & .strDeliveryAdd2
        lstrString = lstrString & vbCrLf & "Delivery Address 3 = " & .strDeliveryAdd3
        lstrString = lstrString & vbCrLf & "Delivery Address 4 = " & .strDeliveryAdd4
        lstrString = lstrString & vbCrLf & "Delivery Address 5 = " & .strDeliveryAdd5
        lstrString = lstrString & vbCrLf & "Delivery Postcode = " & .strDeliveryPostcode
        lstrString = lstrString & vbCrLf & "Account Type = " & .strAccountType
        lstrString = lstrString & vbCrLf & "Receive Mailings = " & .strReceiveMailings
        
        '.strAcctInUseByFlag
    End With

        lstrString = lstrString & vbCrLf

    With gstrConsignmentNote
        lstrString = lstrString & vbCrLf & .strType & " Note"
        lstrString = lstrString & vbCrLf & "Text = " & .strText
    End With
    
        lstrString = lstrString & vbCrLf
    
    With gstrInternalNote
        lstrString = lstrString & vbCrLf & .strType & " Note"
        lstrString = lstrString & vbCrLf & "Text = " & .strText
    End With


    lintFileNum = FreeFile
'    Open App.Path & "\" & App.EXEName & ".log" For Append As FreeFile
    Open "c:\windows\desktop\" & App.EXEName & ".txt" For Append As lintFileNum
    Print #lintFileNum, Now() & vbCrLf & lstrString
    Close lintFileNum

    RunNWait "NOTEPAD c:\windows\desktop\" & App.EXEName & ".txt"
Quit:
    On Error Resume Next
    gdatCentralDatabase.Close
    gdatLocalDatabase.Close
    Set gdatLocalDatabase = Nothing
    Set gdatCentralDatabase = Nothing
    End
    
End Sub
Function CheckField(pstrField As String, pstrFieldName As String, pstrType As String, pintMaxlength, pstrLineNum As String) As String
Dim lstrMessage As String
    
    Select Case pstrType
    Case "STRING"
        If Len(pstrField) < 1 Then
            lstrMessage = lstrMessage & vbCrLf & "Line " & pstrLineNum & " " & vbTab & pstrFieldName & " is not long enough."
        Else
            If Left$(pstrField, 1) <> Chr(34) And Right$(pstrField, 1) <> Chr(34) Then
                lstrMessage = lstrMessage & vbCrLf & "Line " & pstrLineNum & " " & vbTab & pstrFieldName & " is not a character field."
            ElseIf Len(strUnQuoteString(pstrField)) > pintMaxlength Then
                lstrMessage = lstrMessage & vbCrLf & "Line " & pstrLineNum & " " & vbTab & pstrFieldName & " is too long."
            Else
                'If Left$(pstrField, 1) <> Chr(34) And Right$(pstrField, 1) <> Chr(34) Then
                '    lstrMessage = lstrMessage & vbCrLf & "Line " & plngLineNum & " " & pstrFieldName & " doesn't appear to be a string."
                'End If
            End If
        End If
    Case "STRINGNULL"
        If Len(Trim$(strUnQuoteString(pstrField))) > pintMaxlength Then
            lstrMessage = lstrMessage & vbCrLf & "Line " & pstrLineNum & " " & vbTab & pstrFieldName & " is too long."
        Else
            'If Left$(pstrField, 1) <> Chr(34) And Right$(pstrField, 1) <> Chr(34) Then
            '    lstrMessage = lstrMessage & vbCrLf & "Line " & plngLineNum & " " & pstrFieldName & " doesn't appear to be a string."
            'End If
        End If
    Case "LONG"
        If IsNumeric(pstrField) Then
            lstrMessage = lstrMessage & vbCrLf & "Line " & pstrLineNum & " " & vbTab & pstrFieldName & " is not a number."
        End If
    End Select
    
    CheckField = lstrMessage
    
End Function
