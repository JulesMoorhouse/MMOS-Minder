Attribute VB_Name = "modSummary"
Option Explicit

Type SummaryArray
    strSQL          As String
    strSystem       As String
    strShortDescID  As String
    strLabelDesc    As String
    strValueName    As String
    strValue        As String
End Type
Dim mstrSummaryArray(11) As SummaryArray
Global Const gconlngRedSummary = 8421631
Global Const gconlngGreenSummary = 12648384

Function JustifyTwoText(pstrLeftText As String, pstrRightText As String, _
    pobjObject As Object, plngMaxWidth As Long) As String
Dim llngLeftTextWidth As Long
Dim llngRightTextWidth As Long
Dim llngMiddleSpacerWidth As Long
Dim lstrMiddleSpacer As String
Dim llngTotalWidth As Long

    llngLeftTextWidth = pobjObject.TextWidth(pstrLeftText)
    llngRightTextWidth = pobjObject.TextWidth(pstrRightText)
    llngMiddleSpacerWidth = 0
    
    llngTotalWidth = llngLeftTextWidth + llngMiddleSpacerWidth + llngRightTextWidth
    
    Do While llngTotalWidth <= plngMaxWidth
        lstrMiddleSpacer = lstrMiddleSpacer & " "
        llngTotalWidth = pobjObject.TextWidth(pstrLeftText & lstrMiddleSpacer & pstrRightText)
    Loop
    
    JustifyTwoText = pstrLeftText & lstrMiddleSpacer & pstrRightText
    
End Function

Sub PopulateSummaryArray()
Dim lintCounter As Integer
'Also removed unnecessary table. references

    With mstrSummaryArray(lintCounter) 'Total Orders waiting to be Packed
        .strShortDescID = "TOWP"
        .strLabelDesc = "Orders waiting to be packed"
        .strSQL = "SELECT Sum(1) AS [Counter] From " & gtblAdviceNotes & " WHERE (((OrderStatus)='P'));"
        .strValueName = "Counter": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Totals Advice Notes waiting to be printed
        .strShortDescID = "ANWP"
        .strLabelDesc = "Advice Notes waiting to be printed"
        .strSQL = "SELECT Sum(1) AS [Counter] From " & gtblAdviceNotes & " WHERE (((OrderStatus)='A'));"
        .strValueName = "Counter": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Total orders ready to be downloaded
        .strShortDescID = "OWTD"
        .strLabelDesc = "Orders waiting to be downloaded"
        .strSQL = "SELECT Sum(1) AS [Counter] From " & gtblAdviceNotes & " WHERE (((OrderStatus)='C' Or (OrderStatus)='B'));"
        .strValueName = "Counter": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Total Orders on hold awaiting Authorisation
        .strShortDescID = "OHWA"
        .strLabelDesc = "Orders on hold waiting authorisation"
        .strSQL = "SELECT Sum(1) AS [Counter] From " & gtblAdviceNotes & " WHERE (((OrderStatus)='H'));"
        .strValueName = "Counter": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Total consignments waiting to be despatched
        .strShortDescID = "CWTD"
        .strLabelDesc = "Consignments waiting to be despatched"
        .strSQL = "SELECT Sum(1) AS [Counter] From " & gtblPForce & " WHERE (((Status)='P'));"
        .strValueName = "Counter": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Totals Order Value so far Today , excluding X and S
        .strShortDescID = "TOVT"
        .strLabelDesc = "Total order value today (so far)"
        .strSQL = "SELECT Sum(TotalIncVat) AS PriceAmount From " & gtblAdviceNotes & " WHERE (((Format([CreationDate],'dd/mmm/yyyy'))=Format(Date(),'dd/mmm/yyyy')) AND ((OrderStatus)<>'X' And (OrderStatus)<>'S'));"
        .strValueName = "PriceAmount": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter) 'Total Lost Sales due to Out Of Stock products, todal
        .strShortDescID = "TLST"
        .strLabelDesc = "Total loss - items out of stock today"
        .strSQL = "SELECT Sum([Qty]-[DespQty]) AS OOS, Sum(([Qty]-[DespQty])*[Price]) AS PriceAmount FROM " & gtblAdviceNotes & " INNER JOIN " & gtblMasterOrderLines & " ON " & gtblAdviceNotes & ".OrderNum = " & gtblMasterOrderLines & ".OrderNum WHERE (((Format([CreationDate],'dd/mmm/yyyy'))=Format(Date(),'dd/mmm/yyyy')) AND ((" & gtblAdviceNotes & ".OrderStatus)<>'X' And (" & gtblAdviceNotes & ".OrderStatus)<>'S'));"
        .strValueName = "PriceAmount": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter)
        .strShortDescID = "TOVM"
        .strLabelDesc = "Total Order Value last month (calendar)"
        .strSQL = "SELECT Sum(TotalIncVat) AS PriceAmount From " & gtblAdviceNotes & " WHERE ((((Month([CreationDate])))=(Month(Date())-1)) AND ((OrderStatus)<>'X' And (OrderStatus)<>'S'));"
        .strValueName = "PriceAmount": .strSystem = "CLIENT"
    End With
    
    lintCounter = lintCounter + 1
    With mstrSummaryArray(lintCounter)
        .strShortDescID = "TLSA"
        .strLabelDesc = "Total loss - items out of stock"
        .strSQL = "SELECT Sum([Qty]-[DespQty]) AS OOS, Sum(([Qty]-[DespQty])*[Price]) AS PriceAmount FROM " & gtblAdviceNotes & " INNER JOIN " & gtblMasterOrderLines & " ON " & gtblAdviceNotes & ".OrderNum = " & gtblMasterOrderLines & ".OrderNum WHERE (((" & gtblAdviceNotes & ".OrderStatus)<>'X' And (" & gtblAdviceNotes & ".OrderStatus)<>'S'));"
        .strValueName = "PriceAmount": .strSystem = "CLIENT"
    End With   

End Sub
Function BuildSummaryItems(pobjLabel As Object, pobjShape As Object, _
    pstrSystem As String, plngMaxTextWidth As Long, pobjForm As Object) As Integer
Dim lintArrInc As Integer
Dim lintObjectCounter As Integer

    On Error Resume Next

    For lintArrInc = 0 To UBound(mstrSummaryArray)
        With mstrSummaryArray(lintArrInc)
            If .strSystem = pstrSystem Then
                If lintObjectCounter > 0 Then
                    Load pobjLabel(lintObjectCounter)
                    pobjLabel(lintObjectCounter).Visible = True
                    Load pobjShape(lintObjectCounter)
                    pobjShape(lintObjectCounter).Visible = True
                End If
                pobjLabel(lintObjectCounter).Caption = ""
                .strValue = .strValue
                pobjLabel(lintObjectCounter).Caption = JustifyTwoText(.strLabelDesc, .strValue, pobjForm, plngMaxTextWidth)
                lintObjectCounter = lintObjectCounter + 1
            
            End If
        End With
    Next lintArrInc
    
    BuildSummaryItems = lintObjectCounter
    
End Function

Function PopNCalcSysRecStr() As String
Dim lsnaQuery As Recordset
Dim lstrSQL As String
Dim lintArrInc As Integer
Dim lstrStrBuild As String
Const lconstrMsg = "Processing summary item "

    On Error GoTo ErrHandler
        
    For lintArrInc = 0 To UBound(mstrSummaryArray)
        On Error Resume Next
        Screen.ActiveForm.sbStatusBar.Panels(1).Text = _
            lconstrMsg & (lintArrInc + 1) & " of " & UBound(mstrSummaryArray) + 1
        DoEvents
        Err.Number = 0
        On Error GoTo ErrHandler
                
        With mstrSummaryArray(lintArrInc)
            If .strSQL = "" Then Exit Function
            Set lsnaQuery = gdatCentralDatabase.OpenRecordset(.strSQL, dbOpenSnapshot)
            
            If Not (lsnaQuery.BOF And lsnaQuery.EOF) Then
                If IsNull(lsnaQuery(.strValueName)) Then
                    Select Case .strValueName
                    Case "Amount", "Counter"
                        lstrStrBuild = lstrStrBuild & Chr(182) & .strShortDescID & "=0"
                    Case "PriceAmount"
                        lstrStrBuild = lstrStrBuild & Chr(182) & .strShortDescID & "=" & SystemPrice(0)
                    End Select
                Else
                    Select Case .strValueName
                    Case "PriceAmount"
                        lstrStrBuild = lstrStrBuild & Chr(182) & .strShortDescID & "=" & SystemPrice(lsnaQuery(.strValueName))
                    Case Else
                        lstrStrBuild = lstrStrBuild & Chr(182) & .strShortDescID & "=" & lsnaQuery(.strValueName)
                    End Select
                End If
            End If
        End With
    Next lintArrInc
    lsnaQuery.Close
    
    PopNCalcSysRecStr = lstrStrBuild
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "PopNCalcSysRecStr", "Central")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub AddValuestoSummaryArray(pstrSystemString As String)
Dim lintArrInc As Integer
Dim lstrRawItem As String

    For lintArrInc = 2 To UBound(mstrSummaryArray) + 2 '3
        With mstrSummaryArray(lintArrInc - 2)
            lstrRawItem = ReturnNthStr(pstrSystemString, lintArrInc, Chr(182))
            If Left$(lstrRawItem, 4) = .strShortDescID Then
                .strValue = Mid$(lstrRawItem, 6)
            End If
        End With
    Next lintArrInc
    
End Sub
Function GetSummaryString(ByRef pstrString As String) As String
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long

    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT * From " & gtblSystem & " WHERE Item='SysSummary';"
    
    Set lsnaLists = gdatLocalDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not lsnaLists.EOF Then
            GetSummaryString = .Fields("DateCreated")
            pstrString = .Fields("Value")
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "GetSummaryString", "Local", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Sub DebugSummary()
Dim lintArrInc As Integer

    For lintArrInc = 0 To UBound(mstrSummaryArray)
        With mstrSummaryArray(lintArrInc)
            Debug.Print lintArrInc & "#" & .strShortDescID & "#" & _
                .strLabelDesc & "#" & .strValue & "#" & .strValueName
        End With
    Next lintArrInc
    
End Sub
