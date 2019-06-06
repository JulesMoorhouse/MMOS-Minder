Attribute VB_Name = "modNetReps"
Option Explicit

Private Type GroupFieldNames
    strFieldName As String
End Type

Private Type GroupFieldValues
    strFieldValues(6) As String
End Type

Private Type GroupSQLType
    strSQL As String
    strBlockHeader As String
End Type
Global mstrMasterGroupSQL() As GroupSQLType

Function ChooseLayout(pstrLayoutType As String, pobjForm As Form) As Boolean
Dim lstrLayoutsFound As String
Dim lstrSourceFileSpec As String
    
    ChooseLayout = True
    
    On Error GoTo Err_Handler

	If Dir(App.Path & "\Layouts", vbDirectory) = "" Then
		MkDir App.Path & "\Layouts"
	End If
        
	'Copy all files from Server area
	lstrSourceFileSpec = Dir(gstrStatic.strServerPath & "Layouts\*.rpt")

	Do While lstrSourceFileSpec <> ""
		CopyFile gstrStatic.strServerPath & "Layouts\" & _
			lstrSourceFileSpec, App.Path & "\Layouts\" & lstrSourceFileSpec
		lstrSourceFileSpec = Dir
	Loop
	frmChildLayoutSel.LayoutsPath = App.Path & "\Layouts\"

    frmChildLayoutSel.LayoutType = pstrLayoutType
    
    Load frmChildLayoutSel
    
    lstrLayoutsFound = frmChildLayoutSel.LayoutsFound
    
    If lstrLayoutsFound <> "" Then
        ChooseLayout = False
        MsgBox lstrLayoutsFound, , gconstrTitlPrefix & "Layouts"
        Unload frmChildLayoutSel
        Exit Function
    End If
         
    frmChildLayoutSel.Show vbModal
    
    If InitPlotReport(Trim$(gstrReportLayout.strPaperSize), pobjForm) = False Then
        ChooseLayout = False
        Exit Function
    End If
    
    Debug.Print "LAYOU: " & Now()
    
    SetStandardReportDataValues gstrReportLayout.lngLayoutType
    
    Exit Function
Err_Handler:

    MsgBox Err.Number & " " & Err.Description, vbInformation, gconstrTitlPrefix & "ChooseLayout"
    ChooseLayout = False
    Exit Function
    
End Function
Sub AnalyseFields(pstrSQL As String, pstrSysDB As String, pobjObject As Object)
Const lcontintNumOfFields = 100
Dim lintArrInc As Integer
Dim fldTableDef As Field
Dim rstSQL As Recordset

    ReDim lstrFieldNames(0)
    If pstrSysDB = "CENTRAL" Then
        Set rstSQL = gdatCentralDatabase.OpenRecordset(pstrSQL)
    Else
        Set rstSQL = gdatLocalDatabase.OpenRecordset(pstrSQL)
    End If
    
    For lintArrInc = 0 To lcontintNumOfFields
        On Error Resume Next
        Set fldTableDef = rstSQL.Fields(lintArrInc)
        Select Case Err.Number
        Case 3265
            Exit For
        End Select
        
        If lintArrInc > 0 Then
            ReDim Preserve lstrFieldNames(UBound(lstrFieldNames) + 1)
        End If
        
        lstrFieldNames(UBound(lstrFieldNames)).strFieldName = fldTableDef.Name
        lstrFieldNames(UBound(lstrFieldNames)).lngFieldLength = pobjObject.TextWidth(fldTableDef.Name)
        lstrFieldNames(UBound(lstrFieldNames)).lngFieldType = fldTableDef.Type
        lstrFieldNames(UBound(lstrFieldNames)).strLongestField = fldTableDef.Name
        
    Next lintArrInc
    
End Sub
Function AnalyseSQL(pstrSQL As String, pstrSysDB As String, pobjObject As Object, Optional pstrBlockHeader As Variant)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
Dim lintArrInc As Integer
Dim lsngCurrentTextWidth As Single
Dim lstrLineOut As String
Dim lintFileNum As Integer
Dim lbooFoundOrderNumThisRecord As Boolean
Dim lstrLocalErrorStage As String
Dim llngOrderNum As Long
Dim lstrCurrencyString As String
Dim lstrTotalString As String
Dim lstrTotalEqualsString As String
Dim lstrTotalSpaceString As String

    For lintArrInc = 0 To UBound(lstrFieldNames)
        lstrFieldNames(lintArrInc).sngBlockFieldTotal = 0
    Next lintArrInc
    
    On Error GoTo ErrHandler

    ReDim Preserve glngTrackNUpdate(0)
    
    lintFileNum = FreeFile

    Open gstrReport.strDelimDetailsFile For Append As lintFileNum
    
    ReDim lstrReportData(0, UBound(lstrFieldNames))
    
    If pstrSysDB = "CENTRAL" Then
        Set lsnaLists = gdatCentralDatabase.OpenRecordset(pstrSQL, dbOpenSnapshot)
    Else
        Set lsnaLists = gdatLocalDatabase.OpenRecordset(pstrSQL, dbOpenSnapshot)
    End If
    
    With lsnaLists
    
        llngRecCount = 0
        
        Do Until .EOF
            
            For lintArrInc = 0 To UBound(lstrFieldNames)
                lsngCurrentTextWidth = pobjObject.TextWidth(Trim$(ChkNull(.Fields(lstrFieldNames(lintArrInc).strFieldName))))
                
                If lstrFieldNames(lintArrInc).lngFieldLength < lsngCurrentTextWidth Then
                    lstrFieldNames(lintArrInc).lngFieldLength = lsngCurrentTextWidth
                    lstrFieldNames(lintArrInc).strLongestField = Trim$(.Fields(lstrFieldNames(lintArrInc).strFieldName))
                End If
                
                'see if currency
                If IsNull(.Fields(lstrFieldNames(lintArrInc).strFieldName)) = False Then
                    If Left$(Trim$(.Fields(lstrFieldNames(lintArrInc).strFieldName)), 1) = gstrReferenceInfo.strDenomination And Len(Trim$(.Fields(lstrFieldNames(lintArrInc).strFieldName))) < 8 Then
                        lstrFieldNames(lintArrInc).lngFieldType = dbCurrency
                    End If
                
                    If lstrFieldNames(lintArrInc).lngFieldType = dbCurrency Then
                        lstrFieldNames(lintArrInc).sngFieldTotal = lstrFieldNames(lintArrInc).sngFieldTotal + .Fields(lstrFieldNames(lintArrInc).strFieldName)
                        gbooTotalLineRequired = True
                        'ensure curreny fields are big enough
                        If Len(lstrFieldNames(lintArrInc).strLongestField) < 6 Then
                        
                            lstrFieldNames(lintArrInc).strLongestField = "�99.99"
                        End If
                    End If
                End If
                
                lstrCurrencyString = ""
                If lstrFieldNames(lintArrInc).lngFieldType = dbCurrency Then
                    lstrCurrencyString = Trim$(ChkNull(.Fields(lstrFieldNames(lintArrInc).strFieldName)))
                    If lstrCurrencyString = "" Then lstrCurrencyString = "0"
                    lstrCurrencyString = FormatCurrency(lstrCurrencyString, 2)
                    lsngCurrentTextWidth = pobjObject.TextWidth(lstrCurrencyString)
                
                    If lstrFieldNames(lintArrInc).lngFieldLength < lsngCurrentTextWidth Then
                        lstrFieldNames(lintArrInc).lngFieldLength = lsngCurrentTextWidth
                        If Len(lstrFieldNames(lintArrInc).strLongestField) < Len(lstrCurrencyString) Then
                            lstrFieldNames(lintArrInc).strLongestField = lstrCurrencyString
                        End If
                    End If
                    lstrFieldNames(lintArrInc).sngBlockFieldTotal = CCur(lstrFieldNames(lintArrInc).sngBlockFieldTotal) + CCur(Trim$(ChkNull(.Fields(lstrFieldNames(lintArrInc).strFieldName))))
                End If
                
                lstrLineOut = lstrLineOut & Trim$(ChkNull(.Fields(lstrFieldNames(lintArrInc).strFieldName))) & vbTab
                
                lstrLocalErrorStage = "OUTSIDE"

                If lbooFoundOrderNumThisRecord = False Then
                    lstrLocalErrorStage = "INSIDE"
                    llngOrderNum = .Fields("OrdNo")
                    llngOrderNum = .Fields("OrderNum")
                    If llngOrderNum <> 0 Then
                        lbooFoundOrderNumThisRecord = True
                        If llngRecCount = 1 Then
                            glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = llngOrderNum
                        Else
                            ReDim Preserve glngTrackNUpdate(UBound(glngTrackNUpdate) + 1)
                            glngTrackNUpdate(UBound(glngTrackNUpdate)).lngOrderNum = llngOrderNum
                        End If
                    End If
                End If
                lstrLocalErrorStage = "OUTSIDE"
                llngOrderNum = 0
            Next lintArrInc
            lbooFoundOrderNumThisRecord = False
            Print #lintFileNum, lstrLineOut
            lstrLineOut = ""
            .MoveNext
        Loop
    End With
    
    If gstrReport.strReportType = rpTypeGroupings Then
        For lintArrInc = 0 To UBound(lstrFieldNames)
            If lstrFieldNames(lintArrInc).sngBlockFieldTotal <> 0 Then
                lstrTotalEqualsString = lstrTotalEqualsString & Char(Len(lstrFieldNames(lintArrInc).sngBlockFieldTotal) + 1, "=") & vbTab                
                lstrTotalString = lstrTotalString & lstrFieldNames(lintArrInc).sngBlockFieldTotal & vbTab
                lstrTotalSpaceString = lstrTotalSpaceString & Chr(160) & vbTab
            Else
                lstrTotalEqualsString = lstrTotalEqualsString & " " & vbTab
                lstrTotalString = lstrTotalString & " " & vbTab
                lstrTotalSpaceString = lstrTotalSpaceString & " " & vbTab
            End If
        Next lintArrInc
        
        If Not IsMissing(pstrBlockHeader) Then
            Print #lintFileNum, pstrBlockHeader
            gstrReport.lngTotalDetailLines = gstrReport.lngTotalDetailLines + 1
        End If
        
        Print #lintFileNum, lstrTotalEqualsString
        Print #lintFileNum, lstrTotalString
        Print #lintFileNum, lstrTotalEqualsString
        Print #lintFileNum, lstrTotalSpaceString
        gstrReport.lngTotalDetailLines = gstrReport.lngTotalDetailLines + 4
    End If
    
    Close #lintFileNum
    
    If llngRecCount <> 0 Then
        gstrReport.lngTotalDetailLines = gstrReport.lngTotalDetailLines + (llngRecCount - 1)
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case lstrLocalErrorStage
    Case "INSIDE"
        Select Case Err.Number
        Case 3265
            Resume Next
        End Select
    Case Else
        Select Case GlobalErrorHandler(Err.Number, "AnalyseSQL", "Central")
        Case gconIntErrHandRetry
            Resume
        Case gconIntErrHandExitFunction
            Exit Function
        Case Else
            Resume Next
        End Select
    End Select
End Function
Function ReadGroupsBlockIntoArray(pstrSQL As String, pstrGroupingsSQL As String) As Boolean
Const lcontintNumOfFields = 6
Dim lintArrInc As Integer
Dim lintArrInc2 As Integer
Dim fldTableDef As Field
Dim rstSQL As Recordset
Dim llngRecCount As Long
Dim lstrGroupFieldNames(6) As GroupFieldNames
Dim lstrGroupFieldValues() As GroupFieldValues
Dim lstrBlckHeadSufx As String
    
    ReadGroupsBlockIntoArray = True
    
    'Get group block field names
    Set rstSQL = gdatCentralDatabase.OpenRecordset(pstrSQL)
    For lintArrInc = 0 To lcontintNumOfFields
        On Error Resume Next
        Set fldTableDef = rstSQL.Fields(lintArrInc)
        Select Case Err.Number
        Case 3265
            Exit For
        End Select
        
        lstrGroupFieldNames(lintArrInc).strFieldName = fldTableDef.Name
    Next lintArrInc

    'Get group block field values for each field name
    Set rstSQL = gdatCentralDatabase.OpenRecordset(pstrSQL)
    With rstSQL
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            ReDim Preserve lstrGroupFieldValues(llngRecCount)
            For lintArrInc = 0 To UBound(lstrGroupFieldNames)
                If lstrGroupFieldNames(lintArrInc).strFieldName = "" Then Exit For
                lstrGroupFieldValues(llngRecCount - 1).strFieldValues(lintArrInc) = .Fields(lstrGroupFieldNames(lintArrInc).strFieldName)
            Next lintArrInc
            .MoveNext
        Loop
        .Close
    End With
    
    If llngRecCount = 0 Then
        ReadGroupsBlockIntoArray = False
        MsgBox "No matches found!", vbInformation, gconstrTitlPrefix & "Grouped Reporting"
        Exit Function
    End If
    
    'Generate group block SQL statements with group block where clauses populated
    ReDim mstrMasterGroupSQL(llngRecCount - 1)
    For lintArrInc2 = 0 To UBound(lstrGroupFieldValues)
        mstrMasterGroupSQL(lintArrInc2).strSQL = pstrGroupingsSQL
        For lintArrInc = 0 To UBound(lstrGroupFieldNames)
            If lstrGroupFieldNames(lintArrInc).strFieldName = "" Then Exit For
            mstrMasterGroupSQL(lintArrInc2).strSQL = ReplaceStr(mstrMasterGroupSQL(lintArrInc2).strSQL, "{" & lstrGroupFieldNames(lintArrInc).strFieldName & "}", lstrGroupFieldValues(lintArrInc2).strFieldValues(lintArrInc), 1)
            mstrMasterGroupSQL(lintArrInc2).strBlockHeader = mstrMasterGroupSQL(lintArrInc2).strBlockHeader & _
                lstrGroupFieldNames(lintArrInc).strFieldName & " = '" & _
                    lstrGroupFieldValues(lintArrInc2).strFieldValues(lintArrInc) & "'  "
        Next lintArrInc
    Next lintArrInc2
    
End Function

Function Char(pintNumber As Integer, pstrChar As String) As String
Dim lintArrInc As Integer

    For lintArrInc = 0 To pintNumber
        Char = Char & pstrChar
    Next lintArrInc
    
End Function
