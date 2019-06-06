Attribute VB_Name = "modPrintObj"
Option Explicit

Type LabelPage
    intLabelsAcross As Integer
    intLabelsDown As Integer
    lngVertGap As Long
    lngHorizGap As Long
    lngTopMargin As Long
    lngLeftMargin As Long
End Type
Global gstrLabelPage As LabelPage

Type OrderLines
    strCatCode          As String * 10 
    strBinLoc           As String * 50
    strQtyOrd           As String * 4 
    strQtyDesp          As String * 30
    strDesc             As String * 50 
    strUnitPrice        As String * 25
    strTaxCode          As String * 1 
    strAmount           As String * 9 
    strVatAmt           As String * 9 
End Type

Public Type OrderDetail
    strProductLines     As String
    lngNumberOfLines    As String
    booLightParcel      As Boolean
    lngParcelBoxNumber  As Long
    strLines(30)        As OrderLines
End Type
Type BatchLines
    strSQLWhereClause   As String
    strOrderNumberHeader As String
    strBatchTotal       As String
End Type

Global glngGrossWeight  As Long
Global gstrAdviceServiceInd As ListVars
Global glngTotalParcel As Long

Type OrderLineParcelNumber
    lngOrderNum         As Long
    lngParcelBoxNumber  As Long
    strCatNum           As String * 10
End Type

Global gstrOrderLineParcelNumbers() As OrderLineParcelNumber
Global gintNumberOfLinesAPage As Integer
Global glngItemsWouldLikeToPrint As Long
Public Type Box
    lngYPos             As Long
    lngXPos             As Long
    lngYPos2            As Long
    lngXPos2            As Long
    strBoxStyle         As String * 2 
    booRelativeY1         As Boolean
    booRelativeY2         As Boolean
End Type
Global gstrBoxArray() As Box

Global glngGap   As Long

Global lstrFieldNames() As FieldData

Type FieldData
    strFieldName As String
    lngFieldLength As Long
    lngFieldType As Long
    strLongestField As String
    sngFieldTotal As Single
    sngBlockFieldTotal As Single 'ADDED FOR GROUPINGS 
End Type

Global lstrFieldDataArr() As String

Public Type PrinterPageMargins
    lngTopMargin As Long
    lngBottomMargin As Long
    lngLeftMargin As Long
    lngRightMargin As Long
    lngPrintableWidth As Long
    lngPrintableHeight As Long
    lngNonPrintableLeftMargin As Long
    lngNonPrintableTopMargin As Long
    lngStandardRule As Long
End Type

Global glngPageAdjustHeight As Long
Global glngPageAdjustWidth As Long

Global gstrPMarg As PrinterPageMargins
Global glngLeftMargin As Long
Global glngTopMargin As Long

Global glngScaleModeWidth As Long
Global glngScaleModeHeight As Long

Global Const rpScale300 = 0.333333333
Global Const rpScale200 = 0.5      
Global Const rpScale150 = 0.666666667
Global Const rpScale100 = 1          
Global Const rpScale87_5 = 1.142857143
Global Const rpScale75 = 1.29999
Global Const rpScale60 = 1.666666667
Global Const rpScale50 = 2.5       

Global gintScaleFactor As Single


Public Enum rpMargins
    rpMarginNarrow = 520
    rpMarginWide = 520 * 2
End Enum

Public Enum rpSpacing
    rpSpacingSingle = 1
    rpSpacingDouble = 2
End Enum

Public Enum rpFontFactor
    rpFontFactorSmall = 0.5
    rpFontFactorNormal = 1
    rpFontFactorLarge = 2
End Enum
Const lconFontFactorSmall = 0.5
Const lconFontFactorNormal = 1
Const lconFontFactorLarge = 2

Public Enum rpType
    rpTypeDetails = 1
    rpTypePlot = 2
    rpTypeLabels = 3
    rpTypeGroupings = 4 'ADDED FOR GROUPINGS 
End Enum

Public Type ReportDetails
    strReportType           As rpType
    strReportName           As String * 35
    strStartRangeDate       As Date
    strEndRangeDate         As Date
    lngTotalDetailLines     As Long
    intLinesOnAPageAvail    As Long
    intDetailLinesOnAPage   As Long
    intPagesInReport        As Long
    
    booBarsOn               As Boolean
    sngFontSize             As rpFontFactor
    intSpacing              As rpSpacing
    lngMargins              As rpMargins
    
    strDelimDetailsFile     As String
    booShowPageSetup        As Boolean
    booShowOptions          As Boolean
    booOptEnableBars        As Boolean
    booOptEnableLineSpace   As Boolean
    booOptEnableMargins     As Boolean
    booDontDeleteDelim      As Boolean
    booHideZoom             As Boolean 
    booNeedStartDate        As Boolean 
    booNeedEndDate          As Boolean 
End Type
Global gstrReport As ReportDetails
Global glngLastField As Long
Global gbooTotalLineRequired As Boolean

Public Type TextLine
    strText As String
    lngTextWidth As Long
    VarColour As Variant
    booBold As Boolean
End Type

Public Type FontType
    strName             As String * 15
    intSize             As Integer
    booBold             As Boolean
    varColor            As Variant
End Type

Public Enum Alignment
    alRight = 1
    alLeft = 2
    alCenter = 3
End Enum

Type ReportLayoutItem
    lngYPos             As Long
    lngXPos             As Long
    strFontSpecific     As FontType
    lngAlignment        As Long
    intMaxChars         As Integer
End Type

Type ReportDataItem
    strValue            As String * 60
End Type

Type ReportDataDetailItem
    strValue            As String * 30
End Type

Type ReportLayout
    strLayoutName       As String 
    lngLayoutType       As Long 
    strPaperSize        As String * 15
    booFooterBinded         As Boolean 
    strLayoutFileName   As String * 8
    booHasDetails       As Boolean
    strFontStandard     As FontType
    strHeaders(46)      As ReportLayoutItem 
    strDetails(11)       As ReportLayoutItem
    strFooters(33)      As ReportLayoutItem 
    lngTopMargAdj       As Long 
    lngLeftMargAdj      As Long 
End Type

Type ReportData
    strHeaders(46)      As ReportDataItem 
    strDetails(11, 50)   As ReportDataDetailItem
    strFooters(33)      As ReportDataItem 
End Type

Global gintCurrentReportPageNum As Integer
Global gstrReportLayout As ReportLayout
Global gstrReportData As ReportData

Type FontClear
    strCurrentFont As String
    varCurrentFontColour As Variant
    intCurrentFontSize As Integer
    booCurrentFontBold As Boolean
End Type

'Y Pos params
Global Const gconUnder = -1
Global Const gconUnder1Space = -2
Global Const gconSameAsPrev = -3
Global Const gconInvisible = -4

Global Const gconFooterTop = -1

'X Pos Params
Global Const gconAfter = -1
Global Const gconAfterNSpace = -2

Function Bar(pobjObject As Object, plngX As Long, plngY As Long, pstrColour As Variant, pstrWord As String)
Dim lvarCurrentColour As Variant
Dim llngWidth As Long
Dim llngHeight As Long

    If gstrReport.booBarsOn = True Then
        llngWidth = (gstrPMarg.lngPrintableWidth - gstrPMarg.lngStandardRule) / gintScaleFactor
        
        lvarCurrentColour = pobjObject.ForeColor
        
        llngHeight = pobjObject.TextHeight(pstrWord)
        
        pobjObject.ForeColor = pstrColour
        
        pobjObject.Line (plngX, plngY)- _
            ((llngWidth), (plngY + llngHeight)), , BF
        pobjObject.ForeColor = lvarCurrentColour
        
    End If
    
    pobjObject.CurrentX = plngX
    pobjObject.CurrentY = plngY
    
End Function
Sub DrawBoxes(pobjObject As Object, Optional plngYStartPos As Long)
Dim lvarCurrentColour As Variant
Dim llngYPosValue As Long
Dim llngYPos2Value As Long

    Dim lintArrInc As Integer
    
    lvarCurrentColour = pobjObject.ForeColor
    On Error GoTo NextBit
        
    For lintArrInc = 0 To UBound(gstrBoxArray)
    With gstrBoxArray(lintArrInc)
        
        pobjObject.DrawMode = 9 
        
        If .booRelativeY1 = True Then
            llngYPosValue = (.lngYPos / gintScaleFactor) + plngYStartPos
        Else
            llngYPosValue = .lngYPos / gintScaleFactor
        End If
    
        If .booRelativeY2 = True Then
            llngYPos2Value = (.lngYPos2 / gintScaleFactor) + plngYStartPos
        Else
            llngYPos2Value = .lngYPos2 / gintScaleFactor
        End If
        
        Select Case Trim$(.strBoxStyle)
        Case "B"
            pobjObject.Line (.lngXPos / gintScaleFactor, llngYPosValue)-(.lngXPos2 / gintScaleFactor, llngYPos2Value), , B
        Case "BF"
            pobjObject.ForeColor = &H80000008
            pobjObject.Line (.lngXPos / gintScaleFactor, llngYPosValue)-(.lngXPos2 / gintScaleFactor, llngYPos2Value), , BF
            pobjObject.ForeColor = lvarCurrentColour
        Case ""
            pobjObject.Line (.lngXPos / gintScaleFactor, llngYPosValue)-(.lngXPos2 / gintScaleFactor, llngYPos2Value)
        Case "D"
            pobjObject.DrawWidth = 2
            pobjObject.Line (.lngXPos / gintScaleFactor, llngYPosValue)-(.lngXPos2 / gintScaleFactor, llngYPos2Value), , B
            pobjObject.DrawWidth = 1
        End Select
    End With
    Next lintArrInc
NextBit:

End Sub
Function ChopStringintoArray(pUnFormatedString As String, pDelimetingChar As String, ByRef pvarArray() As String) As Integer
'Function to chop up a given string by an given character and to return as an array
Dim lXchar, lLastXchar, lPos As Integer
Dim lMyWord, lMyLastWord As String
Dim lArrInc As Integer
Dim lstrQuoteStrip As String

    lXchar = 0: lPos = 1: lArrInc = 0
    lLastXchar = 0
    
    Do While lPos <= Len(pUnFormatedString)
        lXchar = InStr(lPos, pUnFormatedString, pDelimetingChar, 1)
        If lXchar <> lLastXchar And lXchar > lLastXchar Then
            lMyWord = Mid(pUnFormatedString, lLastXchar + 1, lXchar - lLastXchar - 1)
            ReDim Preserve pvarArray(lArrInc + 1)
            lstrQuoteStrip = Mid(pUnFormatedString, lLastXchar + 1, lXchar - lLastXchar - 1)
    
            If Left(lstrQuoteStrip, 1) = Chr(34) Then lstrQuoteStrip = Right(lstrQuoteStrip, Len(lstrQuoteStrip) - 1)
            If Right(lstrQuoteStrip, 1) = Chr(34) Then lstrQuoteStrip = Left(lstrQuoteStrip, Len(lstrQuoteStrip) - 1)
            pvarArray(lArrInc) = lstrQuoteStrip
            lLastXchar = lXchar
            lArrInc = lArrInc + 1
                ElseIf lLastXchar > lXchar And lXchar = 0 Then
            lXchar = Len(pUnFormatedString) + 1
            lMyLastWord = Mid(pUnFormatedString, lLastXchar + 1, lXchar - lLastXchar - 1)
        End If
        lPos = lPos + 1
    Loop
    
    ReDim Preserve pvarArray(lArrInc + 1)
    
    If Left(lMyLastWord, 1) = Chr(34) Then lMyLastWord = Right(lMyLastWord, Len(lMyLastWord) - 1)
    If Right(lMyLastWord, 1) = Chr(34) Then lMyLastWord = Left(lMyLastWord, Len(lMyLastWord) - 1)
    
    pvarArray(lArrInc) = lMyLastWord
    ChopStringintoArray = lArrInc

End Function
Function PrintNPreview(plngPgNum As Long, pobjObject As Object)
Dim lintArrInc As Integer
Dim lbooItsAPrinter As Boolean
Dim llngHeadNFootTotLines As Long
Dim llngKeptYPos As Long
Dim llngKeptXPos As Long

    lbooItsAPrinter = False
    
    glngGap = 240
    glngGap = glngGap / gintScaleFactor
    
    With pobjObject
        On Error Resume Next
        .Cls
        If Err.Number = 438 Then
            lbooItsAPrinter = True
            gintScaleFactor = 1
        End If
        On Error GoTo 0
        .Font.Name = "Arial"
        .Font.Weight = 400
        
        Select Case gintScaleFactor
        Case rpScale200
            .Font.Size = 21.75
            glngGap = glngGap + 55
        Case rpScale150
            .Font.Size = 17
            glngGap = glngGap - 35
        Case rpScale100
            .Font.Size = 11
        Case rpScale87_5
             .Font.Size = 9.8
             glngGap = glngGap - 20
        Case rpScale75
            .Font.Size = 8.24
            glngGap = glngGap + 4
        Case rpScale60
            .Font.Size = 6
            .Font.Weight = 600
            glngGap = glngGap - 30
        Case rpScale50
            .Font.Size = 4.7
            glngGap = glngGap + 5
        End Select
        
        Select Case gstrReport.sngFontSize
        Case lconFontFactorLarge
            .Font.Size = .Font.Size * 2
        Case lconFontFactorNormal
        
        Case lconFontFactorSmall
            .Font.Size = .Font.Size / 2
        End Select
        
        glngPageAdjustHeight = 6:       glngPageAdjustWidth = 6
        
        CalcPrintableArea
        
        If lbooItsAPrinter = False Then
            .Height = (gstrPMarg.lngPrintableHeight + glngPageAdjustHeight) / gintScaleFactor
            .Width = (gstrPMarg.lngPrintableWidth + glngPageAdjustWidth) / gintScaleFactor
            glngLeftMargin = gstrPMarg.lngStandardRule
            glngTopMargin = gstrPMarg.lngStandardRule
        Else
            glngLeftMargin = gstrPMarg.lngLeftMargin + gstrPMarg.lngStandardRule
            glngTopMargin = gstrPMarg.lngTopMargin + gstrPMarg.lngStandardRule
        End If
            
        Select Case gstrReport.strReportType
        Case rpTypeLabels
            CalcPrintableArea
            LabelPagePrint pobjObject, plngPgNum
        Case rpTypeDetails, rpTypeGroupings 
            RecalcTextWidths pobjObject
            CalcPrintableArea
            
            llngHeadNFootTotLines = 0
        
            llngHeadNFootTotLines = ReportHeader(pobjObject)
    
            llngKeptXPos = pobjObject.CurrentX
            llngKeptYPos = pobjObject.CurrentY
        
            llngHeadNFootTotLines = llngHeadNFootTotLines + 3
            If gstrReport.lngMargins = rpMarginWide Then
                llngHeadNFootTotLines = llngHeadNFootTotLines + 3
            End If
            
            With gstrReport
                .intDetailLinesOnAPage = .intLinesOnAPageAvail - llngHeadNFootTotLines
                If .intSpacing = rpSpacingDouble Then
                     .intDetailLinesOnAPage = .intDetailLinesOnAPage / 2
                End If
                If gbooTotalLineRequired = True Then
                    .lngTotalDetailLines = .lngTotalDetailLines + 2
                End If
                
                .intPagesInReport = .lngTotalDetailLines \ .intDetailLinesOnAPage
                If (.lngTotalDetailLines / .intDetailLinesOnAPage) - .intPagesInReport Then 
                    .intPagesInReport = .intPagesInReport + 1
                End If
                
                If .intPagesInReport = 0 Then .intPagesInReport = 1
            End With
            
            ReportFooter pobjObject, plngPgNum
            
            pobjObject.CurrentX = llngKeptXPos
            pobjObject.CurrentY = llngKeptYPos
            
            ReportDetails pobjObject, plngPgNum
            
            If gbooTotalLineRequired = True Then
                If gstrReport.intPagesInReport = plngPgNum Then
                    ReportTotals pobjObject
                End If
            End If
            
        Case rpTypePlot
            CalcPrintableArea
            
            ReadReportPageData gstrReport.strDelimDetailsFile, plngPgNum
            PlotReport pobjObject
                        
        End Select
    End With
    
    
End Function
Sub LabelPagePrint(pobjObject As Object, plngPage As Long)
Dim lintArrIncAcross As Integer
Dim lintArrIncDownLine As Integer
Dim lintArrIncDown As Integer
Dim llngCurrentXPos As Long
Dim llngCurrentYPos As Long

Dim lintFileNum As Integer
Dim lstrLineData As String
Dim lintLineNum As Integer
Dim llngStartLine As Long
Dim llngArrInc As Long
Dim llnglabelCounter As Long
Dim lstrPrintString As String

    If plngPage > 1 Then
        llnglabelCounter = ((gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown) * plngPage) - 1
    Else
        llnglabelCounter = 0
    End If
    
    lintFileNum = FreeFile
    
    llngStartLine = (gstrLabelPage.intLabelsAcross * gstrLabelPage.intLabelsDown) * (plngPage - 1)
    
    Open gstrReport.strDelimDetailsFile For Input As lintFileNum
    
    'First Starting line
    For llngArrInc = 0 To llngStartLine - 1
        If Not EOF(lintFileNum) Then
            lintLineNum = lintLineNum + 1
            Line Input #lintFileNum, lstrLineData
        End If
    Next llngArrInc
    
    With gstrLabelPage
    
        For lintArrIncDown = 1 To .intLabelsDown
            For lintArrIncAcross = 1 To .intLabelsAcross
                If lintArrIncAcross = 1 Then
                    llngCurrentXPos = .lngLeftMargin
                    If lintArrIncDown = 1 Then
                        llngCurrentYPos = .lngTopMargin
                    Else
                        llngCurrentYPos = pobjObject.CurrentY + .lngVertGap
                    End If
                Else
                    llngCurrentXPos = llngCurrentXPos + .lngHorizGap
                End If
    
        If Not EOF(lintFileNum) Then
            lintLineNum = lintLineNum + 1
            Line Input #lintFileNum, lstrLineData
            
            ChopStringintoArray lstrLineData, vbTab, lstrFieldDataArr()
            
            glngLastField = UBound(lstrFieldNames) '- 1
    
            pobjObject.CurrentY = llngCurrentYPos
            
                For lintArrIncDownLine = 0 To glngLastField
                    pobjObject.CurrentX = llngCurrentXPos
                    lstrPrintString = FormatField(pobjObject, CLng(lintArrIncDownLine))
                    pobjObject.Print lstrPrintString
                Next lintArrIncDownLine
                llnglabelCounter = llnglabelCounter + 1
        
        End If
            Next lintArrIncAcross
        Next lintArrIncDown
    
    End With
    Close #lintFileNum
    
End Sub
Sub RecalcTextWidths(pobjObject As Object)
Dim lintArrInc As Integer
Dim lsngCurrentTextWidth As Single

    For lintArrInc = 0 To UBound(lstrFieldNames)
        lsngCurrentTextWidth = pobjObject.TextWidth(lstrFieldNames(lintArrInc).strLongestField)
        
        lstrFieldNames(lintArrInc).lngFieldLength = lsngCurrentTextWidth
        
    Next lintArrInc
            
End Sub

Sub ReportDetails(pobjObject As Object, plngPage As Long)
Dim lintFileNum As Integer
Dim lstrLineData As String
Dim lintLineNum As Integer
Dim llngStartLine As Long
Dim llngArrInc As Long
Dim llngArrInc2 As Long
Dim llngCurrentYPos As Long
Dim llngLastXPos As Long
Dim lbooThisRowBarPainted As Boolean
Dim lvarBarColour As Variant
Dim lconvarGreen As Variant

    lconvarGreen = RGB(206, 255, 193)
        
    lintFileNum = FreeFile
    lbooThisRowBarPainted = False

    lvarBarColour = lconvarGreen
    
    llngStartLine = gstrReport.intDetailLinesOnAPage * (plngPage - 1)
    
    Open gstrReport.strDelimDetailsFile For Input As lintFileNum
    
    'First Starting line
    For llngArrInc = 0 To llngStartLine - 1
        If Not EOF(lintFileNum) Then
            lintLineNum = lintLineNum + 1
            Line Input #lintFileNum, lstrLineData
        End If
    Next llngArrInc
    
    'Find required lines
    For llngArrInc = llngStartLine To (llngStartLine + gstrReport.intDetailLinesOnAPage) - 1
        If Not EOF(lintFileNum) Then
            lintLineNum = lintLineNum + 1
            Line Input #lintFileNum, lstrLineData
            
            ChopStringintoArray lstrLineData, vbTab, lstrFieldDataArr()
            
            pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
    
            llngCurrentYPos = pobjObject.CurrentY
            
            For llngArrInc2 = 0 To glngLastField
                pobjObject.CurrentY = llngCurrentYPos
                llngLastXPos = pobjObject.CurrentX + lstrFieldNames(llngArrInc2).lngFieldLength
                
                If lbooThisRowBarPainted = False Then
                    Bar pobjObject, pobjObject.CurrentX, pobjObject.CurrentY, lvarBarColour, Trim$(ChkNull(lstrFieldDataArr(llngArrInc2)))
                    lbooThisRowBarPainted = True
                End If
                
                pobjObject.Print FormatField(pobjObject, llngArrInc2)
                If gstrReport.intSpacing = rpSpacingDouble Then
                    pobjObject.Print ""
                    lvarBarColour = vbWhite
                End If
                pobjObject.CurrentX = llngLastXPos + glngGap
            Next llngArrInc2
            lbooThisRowBarPainted = False
            If lvarBarColour = lconvarGreen Then
                lvarBarColour = vbWhite
            Else
                lvarBarColour = lconvarGreen
            End If
        End If
    Next llngArrInc
    
    Close #lintFileNum
    
End Sub

Function ChkNull(pvar As Variant) As String

    If IsNull(pvar) Then
        ChkNull = ""
    Else
        ChkNull = pvar
    End If
    
End Function
Function ReportFooter(pobjObject As Object, plngPageNum As Long) As Integer
Dim llngLastXPos As Long
Dim lintArrInc As Integer
Dim llngCurrentYPos As Long
Dim lstrFTxt() As TextLine
Dim llngKeptYPos As Long
Dim llngKeptXPos As Long
Const lconSpacer = "     "

    llngKeptXPos = pobjObject.CurrentX
    llngKeptYPos = pobjObject.CurrentY
    
    pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
    llngCurrentYPos = gstrPMarg.lngBottomMargin / gintScaleFactor
    pobjObject.CurrentY = llngCurrentYPos
    
    pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
    
    ReDim lstrFTxt(7)
    With lstrFTxt(2)
        .strText = "Date: ": .booBold = True: .VarColour = RGB(192, 192, 192)
    End With
    With lstrFTxt(3)
        .strText = Format$(Now(), "DD/MM/YY HH:MM") & lconSpacer: .booBold = False:  .VarColour = vbBlack
    End With
    With lstrFTxt(4)
        .strText = "Page: ": .booBold = True:   .VarColour = RGB(192, 192, 192)
    End With
    With lstrFTxt(5)
        .strText = plngPageNum:   .booBold = False:    .VarColour = vbBlack
    End With
    With lstrFTxt(6)
        .strText = " of  ": .booBold = True:   .VarColour = RGB(192, 192, 192)
    End With
    With lstrFTxt(7)
        .strText = gstrReport.intPagesInReport & lconSpacer:    .booBold = False:    .VarColour = vbBlack
    End With
    
    With lstrFTxt(0)
        .strText = "Developed by: ": .booBold = True:   .VarColour = RGB(192, 192, 192)
    End With
    With lstrFTxt(1)
        .strText = "Mindwarp Consultancy Ltd " & Chr(169) & "2002" & lconSpacer: .booBold = False:   .VarColour = vbBlack
    End With
    
    TextLineBuild pobjObject, lstrFTxt()

    pobjObject.CurrentX = llngKeptXPos
    pobjObject.CurrentY = llngKeptYPos
    
    
End Function
Function ReportHeader(pobjObject As Object) As Integer
Dim llngLastXPos As Long
Dim lintArrInc As Integer
Dim llngCurrentYPos As Long
Dim lstrHTxt() As TextLine
    
    pobjObject.Font.Underline = False
    llngCurrentYPos = glngTopMargin / gintScaleFactor
    pobjObject.CurrentY = llngCurrentYPos
    pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
    
    ReDim lstrHTxt(5)
    With lstrHTxt(0)
        .strText = "Report: ": .booBold = True: .VarColour = RGB(192, 192, 192)
    End With
    With lstrHTxt(1)
        .strText = gstrReport.strReportName & "     ": .booBold = False:   .VarColour = vbBlack
    End With
    With lstrHTxt(2)
        If gstrReport.booNeedStartDate = True Then 
            .strText = "From: ": .booBold = True:   .VarColour = RGB(192, 192, 192)
        End If
    End With
    With lstrHTxt(3)
        If gstrReport.booNeedStartDate = True Then 
            .strText = Format(gstrReport.strStartRangeDate, "DD MMMM YYYY") & "     ":   .booBold = False:    .VarColour = vbBlack
        End If
    End With
    With lstrHTxt(4)
        If gstrReport.booNeedEndDate = True Then 
            .strText = "To: ": .booBold = True:   .VarColour = RGB(192, 192, 192)
        End If
    End With
    With lstrHTxt(5)
        If gstrReport.booNeedEndDate = True Then 
            .strText = Format$(gstrReport.strEndRangeDate, "DD MMMM YYYY"): .booBold = False:   .VarColour = vbBlack
        End If
    End With
    
    TextLineBuild pobjObject, lstrHTxt()
    
    pobjObject.Print ""
    pobjObject.Print ""
    
    llngCurrentYPos = pobjObject.CurrentY
    pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
        
    pobjObject.Font.Underline = True
    glngLastField = 0
    
    For lintArrInc = 0 To UBound(lstrFieldNames)
        pobjObject.CurrentY = llngCurrentYPos
        llngLastXPos = (pobjObject.CurrentX + lstrFieldNames(lintArrInc).lngFieldLength)
        If (pobjObject.CurrentX + pobjObject.TextWidth(lstrFieldNames(lintArrInc).strFieldName)) < ((gstrPMarg.lngPrintableWidth - gstrPMarg.lngStandardRule) / gintScaleFactor) Then
            pobjObject.Print lstrFieldNames(lintArrInc).strFieldName
        Else
            glngLastField = lintArrInc - 1
            Exit For
        End If
        pobjObject.CurrentX = ((llngLastXPos) + glngGap)
    Next lintArrInc
    
    'If everything fits on apage then
    If glngLastField = 0 Then
        glngLastField = UBound(lstrFieldNames)
        pobjObject.CurrentY = llngCurrentYPos
    End If
    
    pobjObject.Font.Underline = False
    
    pobjObject.Print ""
    pobjObject.Print ""

    ReportHeader = 7
    
End Function
Sub CalcPrintableArea()
Dim TotalPrtAreaVert As Long, TotalPrtAreaHorz As Long
Dim MarginTop As Long, MarginLeft As Long

    TotalPrtAreaHorz = GetDeviceCaps(Printer.hdc, HORZRES)
    TotalPrtAreaVert = GetDeviceCaps(Printer.hdc, VERTRES)
    
    MarginLeft = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
    MarginTop = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
          
    With gstrPMarg
        .lngStandardRule = gstrReport.lngMargins
        .lngRightMargin = .lngStandardRule
        .lngPrintableWidth = TotalPrtAreaHorz * Printer.TwipsPerPixelX
        .lngPrintableHeight = TotalPrtAreaVert * Printer.TwipsPerPixelY
        .lngNonPrintableLeftMargin = (MarginLeft * Printer.TwipsPerPixelX)
        .lngNonPrintableTopMargin = (MarginTop * Printer.TwipsPerPixelY)
        .lngTopMargin = .lngNonPrintableTopMargin
        .lngLeftMargin = .lngNonPrintableLeftMargin
        If gstrReport.lngMargins = rpMarginNarrow Then
            .lngBottomMargin = .lngPrintableHeight - (.lngStandardRule * 1.5)
        Else
            .lngBottomMargin = .lngPrintableHeight - (.lngStandardRule)
        End If
    
    End With
    Printer.Font.Size = 11
    
    gstrReport.intLinesOnAPageAvail = ((((TotalPrtAreaVert - (MarginTop))) * Printer.TwipsPerPixelY) / Printer.TextHeight("Test"))
    
End Sub
Sub SetFont(pstrValue As ReportLayoutItem, pobjObject As Object, pstrFontClear As FontClear)
    
    With pstrValue
        If Not IsBlank(.strFontSpecific.strName) Then
            pstrFontClear.strCurrentFont = pobjObject.Font.Name
            pstrFontClear.varCurrentFontColour = pobjObject.ForeColor
            pstrFontClear.intCurrentFontSize = pobjObject.Font.Size
            pstrFontClear.booCurrentFontBold = pobjObject.FontBold
            
            pobjObject.Font.Name = Trim$(.strFontSpecific.strName)
            If pstrFontClear.varCurrentFontColour <> "" Then
                pobjObject.ForeColor = .strFontSpecific.varColor
            End If
            pobjObject.Font.Size = .strFontSpecific.intSize / gintScaleFactor
            pobjObject.FontBold = .strFontSpecific.booBold
        End If
    End With
    
End Sub
Sub ResetFont(pstrValue As ReportLayoutItem, pobjObject As Object, pstrFontClear As FontClear)

    With pstrValue
        If Trim$(.strFontSpecific.strName) <> "" Then
            pobjObject.Font.Name = pstrFontClear.strCurrentFont
            pobjObject.ForeColor = pstrFontClear.varCurrentFontColour
            pobjObject.Font.Size = pstrFontClear.intCurrentFontSize
            pobjObject.FontBold = pstrFontClear.booCurrentFontBold
        End If
    End With
    
End Sub
Function SwitchToSpace(pobjObject As Object, pstrString As String) As String
Dim llngTextWidth As Long

    Do While llngTextWidth < pobjObject.TextWidth(pstrString)
        SwitchToSpace = SwitchToSpace & " "
        llngTextWidth = pobjObject.TextWidth(SwitchToSpace)
    Loop
    
End Function


Sub TextLineBuild(pobjObject As Object, ByRef pstrText() As TextLine)
Dim llngCurrentYPos As Long
Dim llngCurrentXPos As Long
Dim lintArrInc As Integer
Dim llngCalcdXpos As Long

    llngCurrentXPos = pobjObject.CurrentX
    llngCurrentYPos = pobjObject.CurrentY
    llngCalcdXpos = llngCurrentXPos
    
    For lintArrInc = 0 To UBound(pstrText)
        If Printer.ColorMode = 2 Then
            pobjObject.ForeColor = pstrText(lintArrInc).VarColour
        End If
        pobjObject.Font.Bold = pstrText(lintArrInc).booBold
        pstrText(lintArrInc).lngTextWidth = pobjObject.TextWidth(pstrText(lintArrInc).strText)
        
        If lintArrInc > 0 Then
            llngCalcdXpos = llngCalcdXpos + pstrText(lintArrInc - 1).lngTextWidth
        End If

        pobjObject.CurrentX = llngCalcdXpos
        pobjObject.Print pstrText(lintArrInc).strText
        pobjObject.CurrentY = llngCurrentYPos
    Next lintArrInc

    pobjObject.CurrentX = llngCurrentXPos
    pobjObject.CurrentY = llngCurrentYPos
        
End Sub

Function FormatField(pobjObject As Object, plngIndex As Long) As String
Dim lintStringLen As Integer
Dim lintFieldLen As Integer
Dim lstrString As String

    lstrString = Trim$(ChkNull(lstrFieldDataArr(plngIndex)))
    
    If Left$(lstrString, 1) = Left$(FormatCurrency("0"), 1) And Len(lstrString) < 8 Then
        If lstrFieldNames(plngIndex).lngFieldType <> dbCurrency Then
            lstrString = RightAlign(pobjObject, plngIndex, "")
        End If
    End If
    
    Select Case lstrFieldNames(plngIndex).lngFieldType
    Case dbBigInt
    Case dbBinary
    Case dbBoolean
    Case dbByte
    Case dbChar
    Case dbCurrency
        lstrString = RightAlign(pobjObject, plngIndex, "PRICE")
    Case dbDate
        lstrString = RightAlign(pobjObject, plngIndex, "")
    Case dbDecimal
    Case dbDouble
        lstrString = RightAlign(pobjObject, plngIndex, "")
    Case dbFloat
    Case dbGUID
    Case dbInteger
        lstrString = RightAlign(pobjObject, plngIndex, "")
    Case dbLong
        lstrString = RightAlign(pobjObject, plngIndex, "")
    Case dbLongBinary
    Case dbMemo
    Case dbNumeric
    Case dbSingle
        
    Case dbTime
    Case dbTimeStamp
    Case dbVarBinary
    End Select
    
    FormatField = lstrString
End Function
Function RightAlign(pobjObject As Object, plngIndex As Long, pstrParam As Variant) As String
Dim lstrString As String
Dim lstrLeadingSpaces As String
Dim llngTextWidth As Long

    Select Case pstrParam
    Case "", "PRICE"
        lstrString = Trim$(ChkNull(lstrFieldDataArr(plngIndex)))
    Case "TOTAL"
        lstrString = Trim$(ChkNull(lstrFieldNames(plngIndex).sngFieldTotal))
    End Select
            
    Select Case pstrParam
    Case "PRICE", "TOTAL"
        
        If Trim$(lstrString) = "" Then 
            lstrString = Chr(160) 
        End If
        
        If IsPrice(lstrString) And Left$(lstrString, 1) <> "=" Then   'ADDED FOR GROUPINGS
            lstrString = FormatCurrency(lstrString, 2) 
        End If 'ADDED FOR GROUPINGS
        
        If lstrString = "0" Then 
            lstrString = FormatCurrency(lstrString, 2) 
        End If
    End Select

    Do While llngTextWidth <= lstrFieldNames(plngIndex).lngFieldLength
        lstrLeadingSpaces = lstrLeadingSpaces & " "
        llngTextWidth = pobjObject.TextWidth(lstrLeadingSpaces & lstrString)
    Loop
    
    RightAlign = lstrLeadingSpaces & lstrString
End Function
Function IsPrice(pstrValue As Variant) As Currency
Dim lcurPrice As Currency

    On Error Resume Next
    
    lcurPrice = CCur(pstrValue)
    
    Select Case Err.Number
    Case 13
        IsPrice = 0
        Exit Function
    End Select
    
    IsPrice = Left$(FormatCurrency("0"), 1) & lcurPrice
    
End Function
Function ReportTotals(pobjObject As Object)
Dim llngCurrentYPos As Long
Dim llngLastXPos As Long
Dim lintArrInc As Integer
Dim llngPreviousX As Long

    pobjObject.Font.Underline = True
    llngCurrentYPos = pobjObject.CurrentY
    pobjObject.CurrentY = llngCurrentYPos
    pobjObject.CurrentX = glngLeftMargin / gintScaleFactor
    
    For lintArrInc = 0 To UBound(lstrFieldNames)
        pobjObject.CurrentY = llngCurrentYPos
        llngLastXPos = (pobjObject.CurrentX + lstrFieldNames(lintArrInc).lngFieldLength)
        If (pobjObject.CurrentX + pobjObject.TextWidth(lstrFieldNames(lintArrInc).strFieldName)) < ((gstrPMarg.lngPrintableWidth - gstrPMarg.lngStandardRule) / gintScaleFactor) Then
                
            If lstrFieldNames(lintArrInc).lngFieldType = dbCurrency Then
                llngPreviousX = pobjObject.CurrentX
                pobjObject.Print SwitchToSpace(pobjObject, RightAlign(pobjObject, CLng(lintArrInc), "TOTAL"))
                pobjObject.CurrentX = llngPreviousX
                pobjObject.Print RightAlign(pobjObject, CLng(lintArrInc), "TOTAL")
                
            End If
        End If
        pobjObject.CurrentX = ((llngLastXPos) + glngGap)
    Next lintArrInc
End Function
Sub ReplaceParams(pobjTextBox As Object, ByVal pobjTextBoxThisQuery As Object, _
    pstrStartDate As String, pstrEndDate As String, _
    Optional pstrParam As Variant, Optional pstrValue As Variant)
Dim llngPos As Long
    
    If IsMissing(pstrParam) Then
        pstrParam = ""
    End If
    
    gstrReport.booNeedStartDate = False 
    gstrReport.booNeedEndDate = False 
    
    If InStr(1, pobjTextBox, "[Start Date]") > 0 Then 
        gstrReport.booNeedStartDate = True
    End If
    If InStr(1, pobjTextBox, "[End Date]") > 0 Then 
        gstrReport.booNeedEndDate = True
    End If
    
    pobjTextBoxThisQuery = pobjTextBox
    pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[Start Date]", "#" & Format$(pstrStartDate, "DD/MMM/YYYY") & "#", 1)
    pobjTextBoxThisQuery = ReplaceStr(pobjTextBoxThisQuery, "[End Date]", "#" & Format$(pstrEndDate, "DD/MMM/YYYY") & "#", 1)
    
End Sub
Sub ClearReportBuffer()

    With gstrReport
        .booBarsOn = False
        .intDetailLinesOnAPage = 0
        .intLinesOnAPageAvail = 0
        .intPagesInReport = 0
        .intSpacing = 0
        .lngMargins = 0
        .lngTotalDetailLines = 0
        .sngFontSize = 0
        
        If gstrReport.booDontDeleteDelim = False Then
            .strDelimDetailsFile = ""
        End If
        .strEndRangeDate = 0
        .strReportName = ""
        .strStartRangeDate = 0
        .strReportType = 0
        .booOptEnableBars = True
        .booOptEnableLineSpace = True
        .booOptEnableMargins = True
        .booShowOptions = True
        .booShowPageSetup = True
    End With
    
End Sub
Sub ReadReportPageData(pstrDelimFileName As String, plngPageNumber As Long)
Dim lintFreeFile As Integer
Dim lintRecCounter As Integer

    Debug.Print "B4REA: " & Now()
    
    lintFreeFile = FreeFile

    Open pstrDelimFileName For Random As #lintFreeFile Len = Len(gstrReportData)
    Get #lintFreeFile, plngPageNumber, gstrReportData
    Close #lintFreeFile

    Debug.Print "AFREA: " & Now()

End Sub
Sub PlotReport(pobjObject As Object)
Dim lintDetailItemInc As Integer
Dim lstrFontClear As FontClear
Dim llngCurrentY As Long
Dim llngPreviousYPos As Long
Dim llngHeaderPreviousYPos As Long
Dim llngDetailPreviousYPos As Long
Dim llngHeaderAfterLastXPos As Long
Dim llngFooterAfterLastXPos As Long
Dim llngKeptYPos As Long

    Debug.Print "B4PLO: " & Now

On Error GoTo ErrHandler
    
    With gstrReportLayout
        pobjObject.Font.Size = .strFontStandard.intSize / gintScaleFactor
        pobjObject.Font.Bold = .strFontStandard.booBold
        pobjObject.Font.Name = Trim$(.strFontStandard.strName)
                
        pobjObject.Font.Underline = False
        
        For lintDetailItemInc = 0 To UBound(.strHeaders)
            If Trim$(.strHeaders(lintDetailItemInc).strFontSpecific.strName) <> "" Then SetFont .strHeaders(lintDetailItemInc), pobjObject, lstrFontClear
            pobjObject.ForeColor = vbBlack 
            Select Case .strHeaders(lintDetailItemInc).lngYPos
            Case gconUnder
            Case gconUnder1Space
                pobjObject.Print ""
            Case gconSameAsPrev
                pobjObject.CurrentY = llngHeaderPreviousYPos
            Case gconInvisible
                pobjObject.ForeColor = vbWhite 
            Case 0 ' No value can be zero
                ' Do nothing!
            Case Else
                pobjObject.CurrentY = (.strHeaders(lintDetailItemInc).lngYPos + .lngTopMargAdj) / gintScaleFactor
            End Select
            
            Select Case .strHeaders(lintDetailItemInc).lngXPos
            Case gconAfter
                pobjObject.CurrentX = llngHeaderAfterLastXPos
            Case gconAfterNSpace
                pobjObject.CurrentX = llngHeaderAfterLastXPos + pobjObject.TextWidth(" ")
            Case Else
                pobjObject.CurrentX = (.strHeaders(lintDetailItemInc).lngXPos + .lngLeftMargAdj) / gintScaleFactor
            End Select
            If Not IsBlank(gstrReportData.strHeaders(lintDetailItemInc).strValue) Then
                llngHeaderPreviousYPos = pobjObject.CurrentY
                llngHeaderAfterLastXPos = pobjObject.CurrentX + pobjObject.TextWidth(Trim$(gstrReportData.strHeaders(lintDetailItemInc).strValue))
   
                pobjObject.Print Align(pobjObject, _
                    Trim$(gstrReportData.strHeaders(lintDetailItemInc).strValue), _
                    .strHeaders(lintDetailItemInc).lngAlignment, .strHeaders(lintDetailItemInc).intMaxChars)
                llngCurrentY = pobjObject.CurrentY
            End If
            If Trim$(.strHeaders(lintDetailItemInc).strFontSpecific.strName) <> "" Then ResetFont .strHeaders(lintDetailItemInc), pobjObject, lstrFontClear
        Next lintDetailItemInc
    
        Dim lintDetailLineInc As Integer
        Dim llngUnderYPos As Long
        Dim lintCurrentFontSize As Integer
        
        lintCurrentFontSize = pobjObject.Font.Size
        pobjObject.Font.Size = 5
        
        pobjObject.Print ""
        pobjObject.Font.Size = lintCurrentFontSize
                
        llngUnderYPos = pobjObject.CurrentY
        Dim lstrValueToPrint As String
        If .booHasDetails = True Then
        Dim llngDetailAfterLastXPos As Long
        For lintDetailLineInc = 0 To UBound(gstrReportData.strDetails, 2)
            For lintDetailItemInc = 0 To UBound(.strDetails)
                If Trim$(.strDetails(lintDetailItemInc).strFontSpecific.strName) <> "" Then SetFont .strDetails(lintDetailItemInc), pobjObject, lstrFontClear
                pobjObject.ForeColor = vbBlack 
                Select Case .strDetails(lintDetailItemInc).lngYPos
                Case gconUnder
                    pobjObject.CurrentY = llngUnderYPos
                Case gconUnder1Space
                    pobjObject.CurrentY = llngUnderYPos
                    pobjObject.Print ""
                Case gconSameAsPrev
                    pobjObject.CurrentY = llngDetailPreviousYPos
                Case gconInvisible
                    pobjObject.ForeColor = vbWhite 
                Case Else
                    pobjObject.CurrentY = (.strDetails(lintDetailItemInc).lngYPos + .lngTopMargAdj) / gintScaleFactor
                End Select
                
                Select Case .strDetails(lintDetailItemInc).lngXPos
                Case gconAfter
                    pobjObject.CurrentX = llngDetailAfterLastXPos
                Case gconAfterNSpace
                    pobjObject.CurrentX = llngDetailAfterLastXPos + pobjObject.TextWidth(" ")
                Case Else
                    pobjObject.CurrentX = (.strDetails(lintDetailItemInc).lngXPos + .lngLeftMargAdj) / gintScaleFactor
                End Select
                If Not IsBlank(gstrReportData.strDetails(lintDetailItemInc, lintDetailLineInc).strValue) Then
                    llngDetailPreviousYPos = pobjObject.CurrentY
                    llngDetailAfterLastXPos = pobjObject.CurrentX + pobjObject.TextWidth(Trim$(gstrReportData.strDetails(lintDetailItemInc, lintDetailLineInc).strValue))
                    
                    lstrValueToPrint = Align(pobjObject, _
                        gstrReportData.strDetails(lintDetailItemInc, lintDetailLineInc).strValue, _
                        .strDetails(lintDetailItemInc).lngAlignment, .strDetails(lintDetailItemInc).intMaxChars)
                    
                    pobjObject.Print lstrValueToPrint
                    llngUnderYPos = pobjObject.CurrentY
                End If
                If Trim$(.strDetails(lintDetailItemInc).strFontSpecific.strName) <> "" Then ResetFont .strDetails(lintDetailItemInc), pobjObject, lstrFontClear
            Next lintDetailItemInc
        Next lintDetailLineInc
        End If
        Dim llngYPosBeforeFooterDetailGap As Long
        pobjObject.CurrentY = llngUnderYPos
    
        llngYPosBeforeFooterDetailGap = pobjObject.CurrentY
        If pobjObject.CurrentY <> 0 Then
            lintCurrentFontSize = pobjObject.Font.Size
            pobjObject.Font.Size = 5
            
            pobjObject.Print ""
            pobjObject.Font.Size = lintCurrentFontSize
            
            llngCurrentY = pobjObject.CurrentY
        End If
        
        llngKeptYPos = pobjObject.CurrentY
        DrawBoxes pobjObject, llngKeptYPos - (llngKeptYPos - llngYPosBeforeFooterDetailGap) / 2
        pobjObject.CurrentY = llngKeptYPos
        
        For lintDetailItemInc = 0 To UBound(.strFooters)
            If Trim$(.strFooters(lintDetailItemInc).strFontSpecific.strName) <> "" Then SetFont .strFooters(lintDetailItemInc), pobjObject, lstrFontClear
            pobjObject.ForeColor = vbBlack 
            Select Case .strFooters(lintDetailItemInc).lngYPos
            Case gconUnder
            Case gconUnder1Space
                pobjObject.Print ""
            Case gconSameAsPrev
                pobjObject.CurrentY = llngPreviousYPos
            Case gconInvisible
                pobjObject.ForeColor = vbWhite 
            Case Else
                'Would have been else!
                If gstrReportLayout.booFooterBinded = False Then
                    pobjObject.CurrentY = (((.strFooters(lintDetailItemInc).lngYPos + .lngTopMargAdj) / gintScaleFactor) + llngCurrentY)
                Else
                    pobjObject.CurrentY = (((.strFooters(lintDetailItemInc).lngYPos + .lngTopMargAdj))) / gintScaleFactor
                End If
            End Select
            
            Select Case .strFooters(lintDetailItemInc).lngXPos
            Case gconAfter
                pobjObject.CurrentX = llngFooterAfterLastXPos
            Case gconAfterNSpace
                pobjObject.CurrentX = llngFooterAfterLastXPos + pobjObject.TextWidth(" ")
            Case Else
                pobjObject.CurrentX = (.strFooters(lintDetailItemInc).lngXPos + .lngLeftMargAdj) / gintScaleFactor
            End Select
                        
            If Not IsBlank(gstrReportData.strFooters(lintDetailItemInc).strValue) Then
                llngPreviousYPos = pobjObject.CurrentY
                llngFooterAfterLastXPos = pobjObject.CurrentX + pobjObject.TextWidth(Trim$(gstrReportData.strFooters(lintDetailItemInc).strValue))
                
                pobjObject.Print Align(pobjObject, _
                    Trim$(gstrReportData.strFooters(lintDetailItemInc).strValue), _
                    .strFooters(lintDetailItemInc).lngAlignment, .strFooters(lintDetailItemInc).intMaxChars)
            End If
            If Trim$(.strFooters(lintDetailItemInc).strFontSpecific.strName) <> "" Then ResetFont .strFooters(lintDetailItemInc), pobjObject, lstrFontClear
        Next lintDetailItemInc
    End With
    
    Debug.Print "AFPLO: " & Now
    
    Exit Sub
    
ErrHandler:
Select Case Err.Number
Case 480
    MsgBox "The Amount of system resources required to display this " & vbCrLf & _
            "report at the selected Zoom factor was not available, " & vbCrLf & _
            "please select a lower Zoom factor!", , gconstrTitlPrefix & "Plot Report"
    Exit Sub
Case Else
    MsgBox "An unsual plotting error has occured!" & vbCrLf & _
        "Please report it! Error Number=" & Err.Number, , gconstrTitlPrefix & "Plot Report"
        
End Select

End Sub
Sub ClearReportingDataType(Optional pstrParam As Variant)
Dim lintDetailItemInc As Integer
Dim lintDetailLineInc As Integer

    If IsMissing(pstrParam) Then pstrParam = ""
    
    With gstrReportData
        If pstrParam = "" Then
            For lintDetailItemInc = 0 To UBound(.strHeaders)
                 .strHeaders(lintDetailItemInc).strValue = ""
            Next lintDetailItemInc
    
    
            For lintDetailItemInc = 0 To UBound(.strFooters)
                .strFooters(lintDetailItemInc).strValue = ""
            Next lintDetailItemInc
        End If
        
        If pstrParam = "" Or pstrParam = "D" Then
            For lintDetailItemInc = 0 To UBound(.strDetails)
                For lintDetailLineInc = 0 To UBound(.strDetails, 2)
                    .strDetails(lintDetailItemInc, lintDetailLineInc).strValue = ""
                Next lintDetailLineInc
            Next lintDetailItemInc
        End If
    End With


End Sub
Sub ClearReportingLayoutType()
Dim lintDetailItemInc As Integer
Dim lintDetailLineInc As Integer
    
    With gstrReportLayout
        .strLayoutFileName = ""
        .lngLayoutType = 0
        .strLayoutName = ""
        .booFooterBinded = False
        .lngLeftMargAdj = 0 
        .lngTopMargAdj = 0 
    End With
    
    For lintDetailItemInc = 0 To UBound(gstrReportLayout.strHeaders)
        With gstrReportLayout.strHeaders(lintDetailItemInc)
            .intMaxChars = 0
            .lngAlignment = 0
            .lngXPos = 0
            .lngYPos = 0
            .strFontSpecific.booBold = False
            .strFontSpecific.intSize = 0
            .strFontSpecific.strName = ""
            .strFontSpecific.varColor = 0
        End With
    Next lintDetailItemInc


    For lintDetailItemInc = 0 To UBound(gstrReportLayout.strFooters)
        With gstrReportLayout.strFooters(lintDetailItemInc)
            .intMaxChars = 0
            .lngAlignment = 0
            .lngXPos = 0
            .lngYPos = 0
            .strFontSpecific.booBold = False
            .strFontSpecific.intSize = 0
            .strFontSpecific.strName = ""
            .strFontSpecific.varColor = 0
        End With
    
    Next lintDetailItemInc

    For lintDetailItemInc = 0 To UBound(gstrReportLayout.strDetails)
        With gstrReportLayout.strDetails(lintDetailItemInc)
            .intMaxChars = 0
            .lngAlignment = 0
            .lngXPos = 0
            .lngYPos = 0
            .strFontSpecific.booBold = False
            .strFontSpecific.intSize = 0
            .strFontSpecific.strName = ""
            .strFontSpecific.varColor = 0
        End With
    Next lintDetailItemInc


End Sub
Function Align(pobjObject As Object, pstrValue As String, plngAlignment As Alignment, pintMaxChar As Integer) As String
Dim lstrLeadingSpaces As String
Dim llngTextWidth As Long
Dim llngMaxTextWidth As Long
    
    Select Case plngAlignment
    Case alRight, alCenter
        pstrValue = Trim$(pstrValue)
        If Len(pstrValue) < 9 Then
            llngMaxTextWidth = pobjObject.TextWidth(String(pintMaxChar, "9"))
            Do While llngTextWidth < llngMaxTextWidth
                lstrLeadingSpaces = lstrLeadingSpaces & " "
                llngTextWidth = pobjObject.TextWidth(lstrLeadingSpaces & pstrValue)
            Loop
        End If
        If plngAlignment = alCenter Then
            Align = Left(lstrLeadingSpaces, Fix(Len(lstrLeadingSpaces) / 2)) & pstrValue
        Else
            Align = lstrLeadingSpaces & pstrValue
     
        End If
    Case Else
        Align = pstrValue
    End Select
    

End Function
Sub ShowPlotReport()
    
    With gstrReport
        .intPagesInReport = gintCurrentReportPageNum
    End With
    
    If gintCurrentReportPageNum = 0 Then
        MsgBox "There is nothing to print! Process stopped!", , gconstrTitlPrefix & "Quality Printing"
        
        If gstrReport.booDontDeleteDelim = False Then
            Kill gstrReport.strDelimDetailsFile
        End If
        ClearReportBuffer
        ClearReportingDataType
        ClearReportingLayoutType
        ReDim gstrBoxArray(0)
        Exit Sub
    End If
    
    frmPrintPreview.Show vbModal
    Set frmPrintPreview = Nothing
    DoEvents
    
    If gstrReport.booDontDeleteDelim = False Then
        Kill gstrReport.strDelimDetailsFile
    End If
    ClearReportBuffer
    ClearReportingDataType
    ClearReportingLayoutType
    ReDim gstrBoxArray(0)
    gintCurrentReportPageNum = 0

End Sub
Function InitPlotReport(pstrReportType As String, pobjForm As Form) As Boolean
Dim llngWidth As Long
Dim llngHeight As Long
Dim lintHPos As Integer
Dim lintAPos As Integer
Dim llngHValue As Long
Dim llngAValue As Long
Const lconSettingTractor = "Trying to set tractor feed capability, please check that your printer/driver supports tractor feed (fanfold paper printing)"
Const lconSettingA3 = "Trying to set paper size to A3, please check that your printer/driver supports A3 paper size."
Const lconSettingCustom = "Trying to set custom paper size, please check that your printer/driver supports custom paper size."
Dim lconCurrentSetting As String

    InitPlotReport = True
    On Error GoTo Err_Handler 
    
    Printer.TrackDefault = False

    Printer.NewPage
    Printer.KillDoc
    
    Printer.PaperSize = vbPRPSA4
    
    llngWidth = Printer.Width
    llngHeight = Printer.Height
    
    If Left$(pstrReportType, 1) = "V" Then
        lintHPos = InStr(1, pstrReportType, "H")
        lintAPos = InStr(2, pstrReportType, "A")
        
        llngHValue = Mid(pstrReportType, lintHPos + 1, Len(pstrReportType) - lintAPos)
        llngAValue = Mid(pstrReportType, lintAPos + 1, Len(pstrReportType) - lintAPos)
        
        lconCurrentSetting = lconSettingCustom
        Printer.PaperSize = vbPRPSUser
        lconCurrentSetting = ""
        Printer.Orientation = 1
        Printer.Height = llngHValue * (llngHeight / llngAValue)
        Printer.Width = llngWidth
        lconCurrentSetting = lconSettingTractor
        Printer.PaperBin = vbPRBNTractor
        lconCurrentSetting = ""
    Else
        Select Case pstrReportType
        Case "LISTINGA4"
            Printer.PaperSize = vbPRPSLetter
            On Error Resume Next
            lconCurrentSetting = lconSettingTractor
            Printer.PaperBin = vbPRBNTractor
            On Error GoTo Err_Handler
            lconCurrentSetting = ""
            Printer.Orientation = 1
        Case "CHEQUE"
            lconCurrentSetting = lconSettingCustom
            Printer.PaperSize = vbPRPSUser
            lconCurrentSetting = ""
            Printer.Orientation = 1
            Printer.Height = 101 * (llngHeight / 297)
            Printer.Width = llngWidth
            lconCurrentSetting = lconSettingTractor
            Printer.PaperBin = vbPRBNTractor
            lconCurrentSetting = ""
        Case "LISTINGA3"
            lconCurrentSetting = lconSettingA3
            Printer.PaperSize = vbPRPSA3
            lconCurrentSetting = lconSettingTractor
            Printer.PaperBin = vbPRBNTractor
            lconCurrentSetting = ""
            Printer.Orientation = 2
        Case Else
            Printer.PaperSize = vbPRPSA4
            Printer.Orientation = 1
        End Select
    End If
    
    gstrReport.strReportType = rpTypePlot 
    
    CalcPrintableArea
    
    gintNumberOfLinesAPage = gstrReport.intLinesOnAPageAvail
    
    If gintNumberOfLinesAPage > 66 Then
        MsgBox "This Report can not have more than 66 lines!" & vbCrLf & vbCrLf & _
            "Process Stopped!" & vbCrLf & vbCrLf & _
            "This was unexpected, please report where this occured!", , gconstrTitlPrefix & "Quality Printing"
        InitPlotReport = False
        Exit Function
    End If
    
    gintScaleFactor = 1
    
    pobjForm.Font.Name = "Arial"
    pobjForm.Font.Size = 11 / gintScaleFactor
    
    With gstrReport
        .booShowPageSetup = True
        .booShowOptions = True
        .booOptEnableBars = True
        .booOptEnableLineSpace = True
        .booOptEnableMargins = True
        .strReportType = rpTypePlot
        .intSpacing = rpSpacingSingle
        .lngMargins = rpMarginNarrow
        .sngFontSize = rpFontFactorNormal
        .strDelimDetailsFile = GetTempDir & "a" & Format(Now(), "MMDDSSN") & ".tmp"
    End With
    
    Exit Function
Err_Handler: 

    Dim lstrDebugMsg As String
    Dim lstrRaisedError As String
    Dim lstrErrNum As String
    
    lstrErrNum = Err.Number
    If lstrErrNum <> "0" Then
        lstrRaisedError = lstrErrNum & " " & Err.Description
    End If
    
    If DebugVersion Then
        lstrDebugMsg = vbCrLf & vbCrLf & "Please proceed Debug User!"
    End If
     
    If lconCurrentSetting <> "" Then
        MsgBox lstrRaisedError & vbCrLf & vbCrLf & lconCurrentSetting & vbCrLf & vbCrLf & _
        "You may just need to try a different printer driver!" & lstrDebugMsg, vbInformation, gconstrTitlPrefix & "InitPlotReport - " & pstrReportType
    Else
        MsgBox lstrRaisedError & lstrDebugMsg, vbInformation, gconstrTitlPrefix & "InitPlotReport - " & pstrReportType
    End If
    
    If DebugVersion Then
        Resume Next
    Else
        InitPlotReport = False
        Exit Function
    End If
End Function
Function ReadTempReportLayoutFile(pstrFilename As String) As Boolean
Dim lintFreeFile As Integer
    
    ReadTempReportLayoutFile = True
    On Error GoTo ErrHandler
    lintFreeFile = FreeFile

    Open pstrFilename For Random As #lintFreeFile Len = Len(gstrReportLayout)
    Get #lintFreeFile, 1, gstrReportLayout
    Get #lintFreeFile, 2, gstrBoxArray
    Close #lintFreeFile
    
    Exit Function
ErrHandler:

    Select Case Err.Number
    Case 458
        MsgBox "The report layout files are out-of-date, please contact your IT" & vbCrLf & _
            "support office and inform them!", vbInformation, gconstrTitlPrefix & "Reading Layout File"
        Close #lintFreeFile
        ReadTempReportLayoutFile = False
        Exit Function
    Case Else
        On Error GoTo 0
        Resume
    End Select
    
    
End Function
Sub SetStandardReportDataValues(pltReport As LayoutType)

    With gstrReportData

        Select Case pltReport
        Case ltAdviceNote
            
            If UCase$(App.ProductName) = "LITE" Then
                .strHeaders(lconLiteNag1).strValue = "Generated by MMOS Lite / Demo version."
                .strHeaders(lconLiteNag2).strValue = "This is a limited evaluation software version. "
            Else
                .strHeaders(lconLiteNag1).strValue = vbTab
                .strHeaders(lconLiteNag2).strValue = vbTab
                gstrBoxArray(18).strBoxStyle = "X"
            End If
            
            .strHeaders(lconTitlePageNum).strValue = "Page"
            .strHeaders(lconTitleCustNum).strValue = "Customer Number"
            .strHeaders(lconTitleOrderNum).strValue = "Order Number"
            .strHeaders(lconTitleOrderDate).strValue = "Order Date"
            .strHeaders(lconTitleShipDate).strValue = "Ship Date"
            .strHeaders(lconTitleParcelNum).strValue = "Parcel Number"
        
            .strHeaders(lconTitlProdCode).strValue = "Cat Code"
            .strHeaders(lconTitlBinLoc).strValue = "Bin"
            .strHeaders(lconTitlQty).strValue = "Qty Ord"
            .strHeaders(lconTitlDespQty).strValue = "Qty Desp"
            .strHeaders(lconTitlDesc).strValue = "Product Description"
            .strHeaders(lconTitlUnit).strValue = "Unit Price"
            .strHeaders(lconTitlTax).strValue = "Vat Code"
            .strHeaders(lconTitlAccum).strValue = "Amount"
            
            .strHeaders(lconTitlDelivServ).strValue = "Delivery Service"
            
            .strFooters(lconTitleGoodsNVatTot).strValue = "Goods & VAT"
            .strFooters(lconTitlePostNPack).strValue = "P&P"
            
            If gstrReferenceInfo.booDonationAvail = True Then 
                .strFooters(lconTitleDonation).strValue = "Donation"
            Else
                .strFooters(lconTitleDonation).strValue = vbTab 
                'Hide box around donation
                gstrBoxArray(7).strBoxStyle = "X" 
            End If
            
            .strFooters(lconTitleTotal).strValue = "Total"
            .strFooters(lconTitleOverPay).strValue = "Over Payment"
            .strFooters(lconTitleUnderPay).strValue = "Under Payment"
            .strFooters(lconTitleTotRefund).strValue = "Total Refund"
            
            .strFooters(lconTitleTaxTopSummary).strValue = "Tax Summary"
            .strFooters(lconTitleTaxTopGoodsNPP).strValue = "Goods + P&P"
            .strFooters(lconTitleTaxTopVat).strValue = "Vat"
            .strFooters(lconTitleTaxStnd).strValue = "Standard"
            .strFooters(lconTitleTaxZero).strValue = "Zero"
        
            .strFooters(lconTitleTaxZeroPrcnt).strValue = "   0%"
            
            .strFooters(lconCopyrightNote1).strValue = Chr(169) & " " & gstrOurCompany & " 2002.   Produced by Mindwarp Mai" 
            .strFooters(lconCopyrightNote2).strValue = "l Order System.   "
            
        Case ltParcelForceManifest
            .strHeaders(lconTitlDateDesp).strValue = "Date Despatched :"
            .strHeaders(lconTitlPage).strValue = "Page"
            
            .strHeaders(lconTitleManifest).strValue = "PARCELFORCE MANIFEST"
            
            .strHeaders(lconTitlContractNum).strValue = "Contract No."
            
            .strHeaders(lconTitlDetsConsNum1).strValue = "Consignment"
            .strHeaders(lconTitlDetsConsNum2).strValue = "Number"
            .strHeaders(lconTitlDetsDeliverName).strValue = "Delivery Name"
            .strHeaders(lconTitlDetsPostCode).strValue = "Postcode"
            .strHeaders(lconTitlDetsSendrRef1).strValue = "Senders"
            .strHeaders(lconTitlDetsSendrRef2).strValue = "Reference"
            .strHeaders(lconTitlDetsItems).strValue = "Items"
            .strHeaders(lconTitlDetsSpeHand).strValue = "SpeHand"
            .strHeaders(lconTitlDetsSHS).strValue = "S"
            .strHeaders(lconTitlDetsSHB).strValue = "B"
            .strHeaders(lconTitlDetsSHP).strValue = "P"
            
            .strFooters(lconTitlNumCons).strValue = "Number of Consignments"
            .strFooters(lconTitleNumItems).strValue = "Number of Items"
        Case ltRefundCheques
            'MsgBox " "
        Case ltBatchPickings
            .strHeaders(lconBPTitlBatchPick).strValue = "Batch Picking Summary"
            .strHeaders(lconBPTitlOrderNums).strValue = "Order Nos. "
            .strHeaders(lconBPTitlDetsCatCode).strValue = "Cat. Code"
            .strHeaders(lconBPTitlDetsBin).strValue = "Face/Bin"
            .strHeaders(lconBPTitlDetsQty).strValue = "Quantity"
            .strHeaders(lconBPTitlDetsProd).strValue = "Product"
            .strHeaders(lconBPTitlDetsWeight).strValue = "Weight"
        Case ltCreditCardClaims
            .strHeaders(lconCCCTitlAsAt).strValue = "as at"
            .strHeaders(lconCCCTitlDetsCardNum).strValue = "Credit Card Number"
            .strHeaders(lconCCCTitlDetsAmount).strValue = "Amount"
            .strHeaders(lconCCCTitlDetsAuthCode).strValue = "Auth.Code"
            .strHeaders(lconCCCTitlDetsOrderNum).strValue = "Order No."
            .strHeaders(lconCCCTitlDetsDespDate).strValue = "Desp Date"
            .strHeaders(lconCCCTitlDetsCustomer).strValue = "Customer"
            .strFooters(lconCCCTitlPageTotal).strValue = "Page Total"
            .strFooters(lconCCCTitlPageNum).strValue = "Page No."
            .strFooters(lconCCCTitlGrandTotal).strValue = "Grand Total"
        Case ltInvoice
            .strHeaders(lconSheetTitlName).strValue = "INVOICE"
            .strHeaders(lconSheetTitlPage).strValue = "Page"
            .strHeaders(lconProdTitlServDets).strValue = "Service Details"
            .strHeaders(lconProdTitlNetAmt).strValue = "Net Amount"
            .strHeaders(lconProdTitlVatAmt).strValue = "VAT Amount"
            .strFooters(lconTotTitlNetAmt).strValue = "Total Net Amount"
            .strFooters(lconTotTitlVatAmt).strValue = "Total VAT Amount"
            .strFooters(lconTotTitlCarriage).strValue = "Carriage"
            .strFooters(lconTotTitlInvoice).strValue = "Invoice Total"
            .strFooters(lconBanner1).strValue = "Produced by InvoicePrinter."
            .strFooters(lconBanner2).strValue = Chr(169) & " Mindwarp Consultancy Ltd.  2002"
        End Select
    End With
    
   
End Sub
