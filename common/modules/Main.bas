Attribute VB_Name = "modGeneral"
Option Explicit

Public Type ListDetails
    strCode As String
    strDetail As String
End Type

Global gintActivityDelayCounter As Integer
Global gintOrderLineNumber As Integer

Global Const gconstrReportTypeTextPrint = "Text Print"
Global Const gconstrReportTypeSpreadsheet = "Spreadsheet"
Global Const gconstrReportTypeQuality = "Quality Printing"
Global Const gconstrReportTypeLabel = "Label Printing"
Global Const gconstrReportTypeCustom = "Custom Reporting"
Global Const gconstrReportTypeGrouped = "Grouped Reporting"
'Global Const gconstrReportTypeStockBatch = "Stock Report"
'Global Const gconstrReportTypeAccountType = "AcType Report"
'Global Const gconstrReportTypeUserID = "User Report"
Sub NameForm(pobjForm As Form, Optional pbooNoStatus As Variant)
Dim lintStatusID As Integer

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    lintStatusID = 77
    
    If pobjForm.Name = "frmMainReps" Or pobjForm.Name = "frmMain" Or _
        pobjForm.Name = "frmMainReps" Or pobjForm.Name = "frmAbout" Or _
        pobjForm.Name = "frmWarehouse" Then
        If DateDiff("n", gdatSystemStartTime, Now()) < 1 Then
            If glngNumOfColours < 257 Then
                lintStatusID = 134
            End If
        End If
    End If
                
    If pbooNoStatus = False Then
        ShowStatus lintStatusID
    End If
    
    On Error Resume Next
    
    With gstrCustomerAccount
            
        If .lngCustNum <> 0 Then
            If Trim$(.strSalutation & .strSurname) = "" Then
                pobjForm.Caption = " (M" & .lngCustNum & ")"
            Else
                If gstrAdviceNoteOrder.lngOrderNum = 0 Then
                    pobjForm.Caption = "Talking to: " & Trim$(.strSalutation) & " " & Trim(.strSurname) & " (M" & .lngCustNum & ")"
                Else
					pobjForm.Caption = "Talking to: " & Trim$(.strSalutation) & " " & Trim(.strSurname) & _
						" (M" & .lngCustNum & "/" & gstrAdviceNoteOrder.lngOrderNum & ")"
                End If
            End If   
        End If
    
    End With
    If gstrUserMode <> "" Then
        ModeChange pobjForm, gstrUserMode
    End If
        
End Sub
Sub ModeChange(pobjForm As Form, pstrParam As String)
Dim lintArrInc As Integer
Dim lvarBackColor As Variant

    Select Case UCase$(pstrParam)
    Case gconstrTestingMode
        lvarBackColor = &HFFFF80
    Case gconstrLiveMode
        lvarBackColor = vbButtonFace
    End Select

    On Error Resume Next

    For lintArrInc = 0 To pobjForm.Controls.Count - 1   ' Use the Controls collection
        If Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "txt" And _
             Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "cbo" And _
             Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "dbg" And _
              Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "lst" And _
              Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "tim" And _
              Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "cdg" And _
              Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "tab" And _
              Left$(pobjForm.Controls(lintArrInc).Name, 2) <> "sb" Then
            pobjForm.Controls(lintArrInc).BackColor = lvarBackColor
        End If
    Next

    pobjForm.BackColor = lvarBackColor

    DoEvents
    
End Sub
Function SystemPrice(pstrPrice As String) As String

    SystemPrice = Format(PriceVal(pstrPrice), gstrReferenceInfo.strDenomination & "0.00")

End Function
Function AdvicePrice(pstrPrice As String) As String

    If Trim$(gstrAdviceNoteOrder.strDenom) = "" Then
        MsgBox "Error gstrAdviceNoteOrder.strDenom not set!" & vbCrLf & _
            "Please report where this happen!", , gconstrTitlPrefix & "Advice Price"
    End If
    AdvicePrice = Format(PriceVal(pstrPrice), gstrAdviceNoteOrder.strDenom & "0.00")

End Function
Function FixTel(pstrTelNum As String) As String

Dim lintBracketPos As Integer
Dim lintLetterXPos As Integer
Dim lintIndex As Integer
Dim lstrChar As String
Dim lstrNewTel As String

Dim lintEndOfNum As Integer
Dim lintArrInc As Integer

    For lintArrInc = 1 To Len(pstrTelNum)
        If InStr("0123456789(- ", Mid$(pstrTelNum, lintArrInc, 1)) Then
        Else
            FixTel = pstrTelNum
            Exit Function
        End If
    Next lintArrInc
    
    pstrTelNum = UCase$(pstrTelNum)
    
    lintBracketPos = InStr(pstrTelNum, "(")
    lintLetterXPos = InStr(pstrTelNum, "X")
    
    lintIndex = 1
    lintEndOfNum = 0
    
    While lintIndex <= Len(pstrTelNum) And lintEndOfNum = 0
        lstrChar = Mid$(pstrTelNum, lintIndex, 1)
        lintEndOfNum = InStr("/\X", lstrChar)
        If InStr("0123456789", lstrChar) Then
            lstrNewTel = lstrNewTel + lstrChar
        Else
            If lintEndOfNum > 0 And lintBracketPos > lintIndex Then lintEndOfNum = 0
        End If
        lintIndex = lintIndex + 1
    Wend
    
    lstrNewTel = lstrNewTel & "                    "
    
    If Len(lstrNewTel) = 7 Or Len(lstrNewTel) = 8 Then
        FixTel = Left$(lstrNewTel, 3) & " " & Mid$(lstrNewTel, 4)
    ElseIf Len(lstrNewTel) = 9 Or Len(lstrNewTel) = 10 Then
        FixTel = Left$(lstrNewTel, 4) & " " & Mid$(lstrNewTel, 5)
    Else
        FixTel = Left$(lstrNewTel, 4) & " " & Mid$(lstrNewTel, 5, 3) & " " & Mid$(lstrNewTel, 8)
    End If

End Function
Function CheckCalendar(KeyCode As Integer, pstrDate As String) As String

Dim lstrDate As String
    
    CheckCalendar = pstrDate
    If KeyCode = vbKeyInsert Then
        If IsDate(pstrDate) Then
            lstrDate = Format$(pstrDate, "dd/mmm/yyyy")
        Else
            lstrDate = Format$(Date, "dd/mmm/yyyy")
        End If
    
        frmChildCalendar.CalDate = lstrDate
    
        frmChildCalendar.Show vbModal
        
        CheckCalendar = Format$(frmChildCalendar.CalDate, "dd/mmm/yyyy")
        
        Unload frmChildCalendar
        
        KeyCode = 0
        SendKeys "{tab}"
        DoEvents
    End If

End Function

Function FillList(pstrListName As String, pobjList As Object, pstrIndexArray() As String, _
Optional pstrUserDef1 As Variant, Optional pstrUserDef2 As Variant, Optional pstrParam As Variant, Optional pbooNoClear As Variant)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pstrParam) Then
        pstrParam = ""
    End If
    
    If IsMissing(pbooNoClear) Then
        pbooNoClear = False
    End If
    
    On Error GoTo ErrHandler
    
    If pbooNoClear = False Then
        pobjList.Clear
        llngRecCount = 0
    Else
        ReDim pstrIndexArray(pobjList.ListCount)
        llngRecCount = pobjList.ListCount
    End If
    
    pobjList.BackColor = vbWindowBackground
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".* " & _
        "FROM " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & gtblListDetails & ".ListNum = " & _
        "" & gtblLists & ".ListNum WHERE " & gtblLists & ".ListName='" & pstrListName & "' and " & _
        "" & gtblListDetails & ".InUse = True order by SequenceNum;"
        
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
    
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            ReDim Preserve pstrIndexArray(llngRecCount)
            
            Select Case pstrParam
            Case "CODE&DESC"
                pobjList.AddItem .Fields("ListCode") & " " & .Fields("Description")
            Case Else
                pobjList.AddItem .Fields("Description")
            End Select
            
            pstrIndexArray(llngRecCount - 1) = .Fields("ListCode")
            If Not IsMissing(pstrUserDef1) Then
                ReDim Preserve pstrUserDef1(llngRecCount)
                pstrUserDef1(llngRecCount - 1) = .Fields("UserDef1")
                If Not IsMissing(pstrUserDef2) Then
                    ReDim Preserve pstrUserDef2(llngRecCount)
                    pstrUserDef2(llngRecCount - 1) = .Fields("UserDef2")
                End If
            End If
            
            .MoveNext
        Loop

    End With
    
    If llngRecCount = 0 Then
        pobjList.AddItem "Update Lists!"
        pobjList.BackColor = vbActiveBorder
        ReDim pstrIndexArray(0)
        pstrIndexArray(0) = ""
        If Not IsMissing(pstrUserDef1) Then
            ReDim pstrUserDef1(0)
            pstrUserDef1(0) = ""
            If Not IsMissing(pstrUserDef2) Then
                ReDim pstrUserDef2(0)
                pstrUserDef2(0) = ""
            End If
        End If
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillLists", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function
Function FillCustomReps(pobjList As Object, pobjForm As Form, plngLength As Long, _
    Optional pbooDisplayAll As Variant, Optional pvarType As Variant)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lstrParam As String

    If IsMissing(pbooDisplayAll) Then
        pbooDisplayAll = False
    End If
    
    If IsMissing(pvarType) Then
        pvarType = ""
    End If
    
    On Error GoTo ErrHandler
    pobjList.Clear

    Select Case pbooDisplayAll
    Case False
        lstrSQL = "SELECT * FROM " & gtblCustomReports & " WHERE InUse = True "
        If pvarType <> "" Then lstrSQL = lstrSQL & " and "
    Case True
        lstrSQL = "SELECT * FROM " & gtblCustomReports & " "
        If pvarType <> "" Then lstrSQL = lstrSQL & "Where "
    End Select

    If pvarType <> "" Then
        lstrSQL = lstrSQL & "type = '" & pvarType & "' "
    End If
    
    lstrSQL = lstrSQL & "order by SequenceNum;"

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            
            Select Case Trim$(UCase$(.Fields("Param")))
            Case "CUSTOM"
                lstrParam = gconstrReportTypeCustom
            Case "LABEL"
                lstrParam = gconstrReportTypeLabel
            Case "GROUPED"
                lstrParam = gconstrReportTypeGrouped
            'Case "STOCKBATCH"
            '    lstrParam = gconstrReportTypeStockBatch
            'Case "ACCTTYPE"
            '    lstrParam = gconstrReportTypeAccountType
            'Case "USERID"
            '    lstrParam = gconstrReportTypeUserID
            End Select
            
            pobjList.AddItem ColLeveller(pobjForm, plngLength, Trim$(.Fields("CustRepName"))) & lstrParam
            .MoveNext
        Loop

    End With
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillCustomReps", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function
Sub FillGenericList(pobjList As Object, pstrIndexArray() As String, pstrSQL As String, _
    pstrCodeField As String, pstrDescField As String, pbooStar As Boolean, Optional pstrDB As Variant)

Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    If IsMissing(pstrDB) Then
        pstrDB = "CENTRAL"
    End If
    
    If pstrDB = "" Then
        pstrDB = "CENTRAL"
    End If
    
    On Error GoTo ErrHandler
    pobjList.Clear
    pobjList.BackColor = vbWindowBackground
    lstrSQL = pstrSQL
    
    llngRecCount = 0
    
    If pbooStar = True Then
        llngRecCount = 1
        ReDim Preserve pstrIndexArray(llngRecCount)
        pstrIndexArray(llngRecCount - 1) = "*"
        pobjList.AddItem "All"
    End If
    
    If pstrDB = "CENTRAL" Then
        Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    ElseIf pstrDB = "LOCAL" Then
        Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    End If
    
    With lsnaLists
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            ReDim Preserve pstrIndexArray(llngRecCount)
            pobjList.AddItem .Fields(pstrDescField)
            pstrIndexArray(llngRecCount - 1) = .Fields(pstrCodeField)
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
        ReDim pstrIndexArray(0)
        pstrIndexArray(0) = ""
    End If
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillGenericList", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function GetListCodeDesc(pstrListName, pstrListCode)
Dim lsnaListDetails As Recordset
Dim lstrSQL As String

    If IsBlank(pstrListCode) Then
        Exit Function
    End If
    
    On Error GoTo ErrHandler
        
    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".* " & _
        "FROM " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & gtblListDetails & ".ListNum = " & _
        "" & gtblLists & ".ListNum WHERE " & gtblLists & ".ListName='" & pstrListName & "' and " & _
        "" & gtblListDetails & ".InUse = True and " & gtblListDetails & ".ListCode='" & _
        Trim$(pstrListCode) & "'" & " order by " & gtblListDetails & ".ListCode; "

    Set lsnaListDetails = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    If Not (lsnaListDetails.BOF And lsnaListDetails.EOF) Then
        GetListCodeDesc = lsnaListDetails("Description")
    End If
    
    lsnaListDetails.Close

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetListCodeDesc", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub GetCustomRep(ByVal pobjSQL As Object, pstrRepId As String, ByRef plngSeqNum As Long, _
    ByRef pstrSysDB As String, ByRef pobjInUse As Integer, ByRef pbooLocked As Boolean, pintOrientation As Integer)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    
    ShowStatus 78
    If Val(pstrRepId) > 0 Then
        lstrSQL = "SELECT * FROM " & gtblCustomReports & " WHERE CRID = " & CLng(pstrRepId) & ";"
    Else
        lstrSQL = "SELECT * FROM " & gtblCustomReports & " WHERE CustRepName = '" & pstrRepId & "';"
    End If
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    With lsnaLists
        llngRecCount = 0
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            pobjSQL = Trim$(.Fields("ReportSQL"))
            plngSeqNum = .Fields("SequenceNum")
            pstrSysDB = Trim$(.Fields("SysDB"))

            If CBool(Trim$(.Fields("InUse"))) = True Then
                pobjInUse = 1
            Else
                pobjInUse = 0
            End If

            If Left$(.Fields("Settings"), 1) = "P" Then
                pintOrientation = vbPRORPortrait
            ElseIf Left$(.Fields("Settings"), 1) = "L" Then
                pintOrientation = vbPRORLandscape
            End If
            
            .MoveNext
        Loop
    End With
    
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetCustomRep", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub
Function GreetingTime(Optional ptimTime As Variant) As String

    If IsMissing(ptimTime) Then
        ptimTime = Now
    End If
    If IsBlank(ptimTime) Then
        ptimTime = Now
    End If
    
    If Format$(ptimTime, "hh") < 12 Then
        GreetingTime = "Good Morning"
    ElseIf Format$(ptimTime, "hh") < 18 Then
        GreetingTime = "Good Afternoon"
    Else
        GreetingTime = "Good Evening"
    End If

End Function
Function SelectListItem(pstrCode As String, pobjList As Object, plistIndexArray() As String)

Dim lintArrInc As Integer
On Error GoTo Err_Hand

    For lintArrInc = 0 To UBound(plistIndexArray)
        If plistIndexArray(lintArrInc) = pstrCode Then
            SelectListItem = lintArrInc
            pobjList.ListIndex = lintArrInc
            Exit For
        End If
    Next lintArrInc
    
    Exit Function
Err_Hand:
    If Err = 9 Then Exit Function

End Function
Function NotNull(pobjList As Object, pstrArray() As String) As String

    If pobjList.ListIndex = -1 Then
        NotNull = ""
    Else
        NotNull = pstrArray(pobjList.ListIndex)
    End If

End Function

Function CheckKeyAsciiTelNum(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiTelNum = pstrKeyAscii
    ElseIf InStr(" 0123456789", Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiTelNum = pstrKeyAscii
    Else
        CheckKeyAsciiTelNum = 0
    End If

End Function
Function CheckKeyAsciiValid(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiValid = pstrKeyAscii
    ElseIf InStr(" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890=-_!ï¿½$%^&*(){}[]:;@<>,.?/~#\|`ï¿½'" & Chr(34), Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiValid = pstrKeyAscii
    Else
        CheckKeyAsciiValid = 0
    End If

End Function
Function CheckKeyAsciiValidEmail(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiValidEmail = pstrKeyAscii

    ElseIf InStr(" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890=-_!ï¿½$%^&*(){}[]:;@<>,.?/~#\|`ï¿½'", Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiValidEmail = pstrKeyAscii
    Else
        CheckKeyAsciiValidEmail = 0
    End If

End Function
Function CheckKeyAsciiValidPostCode(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiValidPostCode = pstrKeyAscii
    ElseIf InStr(" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890", Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiValidPostCode = pstrKeyAscii
    Else
        CheckKeyAsciiValidPostCode = 0
    End If

End Function
Function CheckKeyAsciiValidAlphaNum(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiValidAlphaNum = pstrKeyAscii
    ElseIf InStr(" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-\'" & Chr(34), Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiValidAlphaNum = pstrKeyAscii
    Else
        CheckKeyAsciiValidAlphaNum = 0
    End If
End Function
Function CheckKeyAsciiValidNum(pstrKeyAscii As Integer)

    If Chr(pstrKeyAscii) = vbBack Then
        CheckKeyAsciiValidNum = pstrKeyAscii

    ElseIf InStr(" 1234567890" & Chr(34), Chr(pstrKeyAscii)) <> 0 Then
        CheckKeyAsciiValidNum = pstrKeyAscii
    Else
        CheckKeyAsciiValidNum = 0
    End If
    
End Function
Function SQLFixup(pstrTextIn As String) As String

    SQLFixup = ReplaceStr(pstrTextIn, "'", "''", 0)
    
End Function

Function JetSQLFixup(pstrTextIn As String) As String
Dim lstrTemp As String

    lstrTemp = ReplaceStr(pstrTextIn, "'", "''", 0)
    JetSQLFixup = ReplaceStr(lstrTemp, "|", "' & chr(124) & '", 0)
    
End Function
Function AmpersandDouble(pstrText As String) As String
Dim lstrNewMessage As String
Dim lstrOldMessage As String

    If InStr(pstrText, "&") > 0 Then
        lstrOldMessage = pstrText
        lstrNewMessage = ""
        Do Until InStr(lstrOldMessage, "&") = 0
            lstrNewMessage = lstrNewMessage & Mid$(lstrOldMessage, 1, InStr(lstrOldMessage, "&")) & "&"
            lstrOldMessage = Mid$(lstrOldMessage, InStr(lstrOldMessage, "&") + 1)
        Loop
        lstrNewMessage = lstrNewMessage & lstrOldMessage
    Else
       lstrNewMessage = pstrText
    End If
    
    AmpersandDouble = lstrNewMessage

End Function
Function Line(pintLength) As String
Dim lintArrInc As Integer

    For lintArrInc = 0 To pintLength
        Line = Line & "-"
    Next lintArrInc
    
End Function
Function PriceVal(pstrPrice As String) As Currency

    If Left$(Trim$(pstrPrice), 1) = gstrReferenceInfo.strDenomination Then
        PriceVal = CCur(Right$(Trim$(pstrPrice), Len(pstrPrice)))
    Else
        If IsNull(pstrPrice) Or IsBlank(pstrPrice) Then
            PriceVal = CCur(0)
        Else
            PriceVal = CCur(pstrPrice)
        End If
    End If
    
End Function
Sub ConvertContact(ByRef pstrContactString As String, ByRef pstrTitle As String, ByRef pstrFirstname As String, ByRef pstrSurname As String)
Dim lintNumOfSpaces As Integer
Dim lintCurrentPos As Integer
lintCurrentPos = 1
    
    If InStr(1, pstrContactString, "&") = 0 Then
        Do Until lintNumOfSpaces = 3
            If InStr(lintCurrentPos, pstrContactString, " ") = 0 Then
                Exit Do
            Else
                lintNumOfSpaces = lintNumOfSpaces + 1
                lintCurrentPos = InStr(lintCurrentPos, pstrContactString, " ") + 1
            End If
        Loop
        
        'two words
        If lintNumOfSpaces = 1 Then
            'Mrs mr
            Select Case UCase$(Left$(pstrContactString, 3))
            Case "MR "
                pstrTitle = "Mr"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 3)
                Exit Sub
            Case "DR "
                pstrTitle = "Dr"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 3)
                Exit Sub
            Case "MS "
                pstrTitle = "Ms"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 3)
                Exit Sub
            End Select
            
            Select Case UCase$(Left$(pstrContactString, 4))
            Case "MISS "
                pstrTitle = "Miss"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 4)
                Exit Sub
            Case "MRS "
                pstrTitle = "Mrs"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 4)
                Exit Sub
            End Select
            If UCase$(Left$(pstrContactString, 7)) = "SISTER " Then
                pstrTitle = "Sister"
                pstrSurname = Right$(pstrContactString, Len(pstrContactString) - 7)
                Exit Sub
            End If
            pstrFirstname = Left$(pstrContactString, InStr(1, pstrContactString, " "))
            pstrSurname = Right$(pstrContactString, Len(pstrContactString) - InStr(1, pstrContactString, " "))
            Exit Sub
        End If
        
        'three words
        If lintNumOfSpaces = 2 Then
            Select Case UCase$(Left$(pstrContactString, 3))
            Case "MR "
                pstrTitle = "Mr"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
              Exit Sub
            Case "DR "
                pstrTitle = "Dr"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
                 Exit Sub
            Case "MS "
                pstrTitle = "Ms"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
                 Exit Sub
            End Select
            
            Select Case UCase$(Left$(pstrContactString, 4))
            Case "MISS"
                pstrTitle = "Miss"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
                Exit Sub
            Case "MRS "
                pstrTitle = "Mrs"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
                Exit Sub
            End Select
            If UCase$(Left$(pstrContactString, 7)) = "SISTER " Then
                pstrTitle = "Sister"
                pstrFirstname = Mid$(pstrContactString, InStr(1, pstrContactString, " "), InStr(InStr(1, pstrContactString, " ") + 1, pstrContactString, " ") - InStr(1, pstrContactString, " "))
                pstrSurname = Reverse(Left$(Reverse(pstrContactString), InStr(1, Reverse(pstrContactString), " ")))
                Exit Sub
            End If
        End If
    End If

End Sub

Function Reverse(ByVal pstrString As String) As String
Dim lintArrInc As Integer

    For lintArrInc = 0 To Len(pstrString) - 1
        Reverse = Reverse & Right$(pstrString, 1)
        pstrString = Left$(pstrString, Len(pstrString) - 1)
    Next lintArrInc

End Function

Function ValNull(pvarString As Variant) As Long

    If IsNull(pvarString) Then
        ValNull = 0
    Else
        ValNull = CLng(pvarString)
    End If
    
End Function

Sub ErrorLogging(pstrSQLorString As String, Optional pstrFilename As Variant)
Dim lintFileNum As Integer
Dim lstrFileName As String

    If IsMissing(pstrFilename) Then
        pstrFilename = ""
    End If
    
    If Trim$(pstrFilename) = "" Then
        lstrFileName = gstrStatic.strServerPath & "Error.Log"
    Else
        lstrFileName = gstrStatic.strServerPath & pstrFilename
    End If
    
    lintFileNum = FreeFile
    
    Open lstrFileName For Append As lintFileNum
    Print #lintFileNum, pstrSQLorString
    Close #lintFileNum
    
End Sub
Sub ComingSoon()

    MsgBox "Coming Soon!", , gconstrTitlPrefix & "Coming Soon"
                    
End Sub
Function CA(pstrString As String) As String
'ConvertApostrophe
Dim lintApostrophePos As Integer
Dim lintArrInc As Integer
    
    For lintArrInc = 1 To Len(pstrString)
        lintApostrophePos = InStr(lintArrInc, pstrString, "'")
        
        If lintApostrophePos = 0 Then
            Exit For
        Else
            Mid$(pstrString, lintApostrophePos, 1) = "`"
        End If
    Next lintArrInc
    
    CA = pstrString
    
End Function
Sub CheckFile()
Dim lintFileNum As Integer
Dim lintFileNum2 As Integer
Dim lstrLineData As String
Dim lintLineNum As Integer
    lintFileNum = FreeFile
    
    Open "c:\windows\desktop\check.txt" For Append As lintFileNum
    lintFileNum2 = FreeFile
    Open "c:\manifest.txt" For Input As lintFileNum2
    While Not EOF(lintFileNum2)
        lintLineNum = lintLineNum + 1
        Line Input #lintFileNum2, lstrLineData
        If Len(lstrLineData) > 79 Then
            Print #lintFileNum, "Line=" & lintLineNum & " Len=" & Len(lstrLineData)
        End If
    Wend
    
    Close #lintFileNum2
    Close #lintFileNum

End Sub
Sub CheckActivity()

    On Error GoTo ErrorHandler
    
    gintActivityDelayCounter = gintActivityDelayCounter + 1
    
    If gintActivityDelayCounter >= 4 Then
        If FileDateTime(gstrStatic.strServerPath & gstrStatic.strPrograms(0).strProgram) > FileDateTime(AppPath & gstrStatic.strPrograms(0).strProgram) Then
            ShowStatus 37
            'MsgBox "newer"
        Else
            'MsgBox "Older"
        End If
        
        If NewStockAvailable Then
            ShowStatus 40
        End If
        gintActivityDelayCounter = 0
    End If
    
    CheckForMessages
    
ErrorHandler:
    Exit Sub
    
End Sub
Function CheckEmailAddress(email) As String
Dim FirstOcrDot
Dim FirstOcrAt
Dim lastOcrDot
    
    lastOcrDot = InStrRev(email, ".")
    FirstOcrDot = CInt(InStr(1, email, "."))
    FirstOcrAt = CInt(InStr(1, email, "@"))
    
    If Len(email) > 5 Then
        If FirstOcrAt > 0 Then
            If Mid(email, FirstOcrAt + 1, 1) <> "." And Mid(email, FirstOcrAt - 1, 1) <> "." Then
                If InStr(1, email, " ") = 0 Then
                    If InStrRev(email, ".") > CInt(FirstOcrAt) + 1 Then
                        If lastOcrDot < Len(email) - 1 Then
                            If FirstOcrAt < Len(email) - 1 And FirstOcrAt > 1 Then
                                If FirstOcrDot > 1 Then
                                    CheckEmailAddress = ""
                                    Exit Function
                                Else
                                    'Response.Write "Dot in starting<br>"
                                End If
                            Else
                                'Response.Write "@ Not between first and last<br>"
                            End If
                        Else
                            'Response.Write ". before last two chars<br>"
                        End If
                    Else
                        'Response.Write "Last Dot after @ missing or @ and . are not seperated by a Char<br>"
                    End If
                Else
                    'Response.Write "Space Not Allowed<br>"
                End If
            Else
                'Response.Write "A Dot Just before or after @ not allowed<br>"
            End If
        Else
            'Response.Write "A required Character Missing<br> "
        End If
    Else
        'Response.Write "Email cannot be less than 6 characters<br> "
    End If
        
    CheckEmailAddress = "This Email address does not appear to be valid!"
    
End Function
Function NumToWord(pintNumber As Integer) As String

    Select Case pintNumber
    Case 1: NumToWord = "One"
    Case 2: NumToWord = "Two"
    Case 3: NumToWord = "Three"
    Case 4: NumToWord = "Four"
    Case 5: NumToWord = "Five"
    Case 6: NumToWord = "Six"
    Case 7: NumToWord = "Seven"
    Case 8: NumToWord = "Eight"
    Case 9: NumToWord = "Nine"
    Case 10: NumToWord = "Ten"
    End Select
    
    NumToWord = UCase$(NumToWord)
    
End Function
Function ColLeveller(pobjForm As Form, plngLength As Long, pstrString As String) As String
Dim llngTextWidth As Long
Dim lstrPadding As String
    
    Do While llngTextWidth < plngLength
        lstrPadding = lstrPadding & " "
        llngTextWidth = pobjForm.TextWidth(pstrString & lstrPadding)
    Loop
    
    ColLeveller = pstrString & lstrPadding & vbTab
    
End Function
Function StripColLevelPadding(pstrString As String) As String
Dim lintTabPos As Integer

    lintTabPos = InStr(1, pstrString, vbTab) - 1
    StripColLevelPadding = Trim$(Left$(pstrString, lintTabPos))
    
End Function
Function GrapLine(pstrText As String, pintLine As Integer, pintLength As Integer)
Dim lintArrInc As Integer
Dim lstrCurrentChar As String
Dim lstrAccumSentence As String
Dim lstrSafeAccumSentence As String
Dim lstrSentArry() As String
Dim lintSafeArrInc As Integer
Dim lintArrCount As Integer
Dim lbooFin As Boolean

    lbooFin = False
    pstrText = Trim$(pstrText) & Space(pintLength)
    lintArrCount = 0
    
    Do Until lbooFin = True
        lintArrInc = lintArrInc + 1
        lstrCurrentChar = Mid$(pstrText, lintArrInc, 1)

        If Len(lstrAccumSentence) < pintLength Then
            lstrAccumSentence = lstrAccumSentence & lstrCurrentChar
            If lstrCurrentChar = " " Then
                lintSafeArrInc = lintArrInc
                lstrSafeAccumSentence = lstrAccumSentence
            End If
        End If
        If Len(lstrAccumSentence) >= pintLength Then
            lintArrInc = lintSafeArrInc
            If lintArrCount = 0 Then
                ReDim lstrSentArry(0)
            Else
                ReDim Preserve lstrSentArry(UBound(lstrSentArry) + 1)
            End If
            lintArrCount = lintArrCount + 1
            lstrSentArry(UBound(lstrSentArry)) = lstrSafeAccumSentence

            lstrAccumSentence = ""
            lstrSafeAccumSentence = ""
        End If
        If lintArrInc >= Len(pstrText) + 1 Then
            lbooFin = True
        End If
    Loop
    If pintLine <= UBound(lstrSentArry) Then
        GrapLine = lstrSentArry(pintLine)
    Else
        GrapLine = ""
    End If
    
End Function
Sub GetListDetailToArray(pstrListName As String, pstrIndexArray() As ListDetails)
Dim lstrSQL As String
Dim lstrMsg As String
Dim lsnaLists As Recordset
Dim llngRecCount As Long

    lstrSQL = "SELECT " & gtblLists & ".ListName, " & gtblListDetails & ".* " & _
    "FROM " & gtblListDetails & " INNER JOIN " & gtblLists & " ON " & gtblListDetails & ".ListNum = " & gtblLists & ".ListNum " & _
    "WHERE (((" & gtblLists & ".ListName)='" & pstrListName & "'));"

    On Error GoTo ErrHandler
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        llngRecCount = 0
        
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            ReDim Preserve pstrIndexArray(llngRecCount)
            pstrIndexArray(llngRecCount - 1).strCode = .Fields("ListCode")
            pstrIndexArray(llngRecCount - 1).strDetail = .Fields("Description")
            .MoveNext
        Loop
    End With
    lsnaLists.Close


Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetListDetailToArray", "Local")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Function MatchListDetailArray(pstrCode As String, pstrIndexArray() As ListDetails) As String
Dim lintArrInc As Integer

    For lintArrInc = 0 To UBound(pstrIndexArray)
        If pstrCode = pstrIndexArray(lintArrInc).strCode Then
            MatchListDetailArray = pstrIndexArray(lintArrInc).strDetail
            Exit For
        End If
    Next lintArrInc
    
End Function
Function LogUsage(pstrProcessType As String, pstrCategory As String, pstrItem As String)
Dim lstrSQL As String
Dim lsnaLists As Recordset
Dim llngUsage As Long
    
    lstrSQL = "SELECT * from " & gtblUsage & " where ProcessType = '" & pstrProcessType & "' and " & _
        "Item = '" & pstrItem & "' and UserId = '" & Trim$(gstrGenSysInfo.strUserName) & _
        "' and Category = '" & pstrCategory & "';"

    On Error GoTo ErrHandler
    'Get last Usage
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            llngUsage = .Fields("Usage")
        End If
    End With
    lsnaLists.Close
        
    'Add 1 or if not used before add new record.
    If llngUsage = 0 Then
        lstrSQL = "Insert Into " & gtblUsage & " ( UserID, ProcessType, Item, " & _
            "LastUsed, Creation, Usage, Category) Select '" & _
            Trim$(gstrGenSysInfo.strUserName) & "', '" & Trim$(pstrProcessType) & "', '" & Trim$(pstrItem) & "', #" & _
            Now() & "#, #" & Now() & "#,1,'" & Trim$(pstrCategory) & "';"
    Else
        lstrSQL = "UPDATE " & gtblUsage & " SET LastUsed = Now(), Usage = " & _
            (llngUsage + 1) & " where UserID = '" & Trim$(gstrGenSysInfo.strUserName) & _
            "' and ProcessType = '" & Trim$(pstrProcessType) & "' and Item = '" & _
            Trim$(pstrItem) & "' and Category = '" & Trim$(pstrCategory) & "';"
    End If
    gdatCentralDatabase.Execute lstrSQL
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "LogUsage", "Central", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
    
End Function
