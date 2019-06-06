Attribute VB_Name = "modQualityLayouts"
Option Explicit

Public Enum LayoutType
    ltAdviceNote = 1
    ltParcelForceManifest = 2
    ltRefundCheques = 3
    ltBatchPickings = 4
    ltCreditCardClaims = 5
    ltInvoice = 6
    ltAdviceWithAddress = 7
    ltHeaderReport = 8
End Enum

Global Const glngTotalNumOfLayouts = 7

'Advice Note Constants Start
'Header Constants
Global Const lconAdviceTitle = 0:      Global Const lconTitlePageNum = 1: Global Const lconPageNum = 2
Global Const lconDeliverTitle = 3

Global Const lconTitleCustNum = 4:     Global Const lconTitleOrderNum = 5:     Global Const lconTitleOrderDate = 6:    Global Const lconTitleShipDate = 7:
Global Const lconTitleParcelNum = 8:

Global Const lconCustNum = 9:          Global Const lconOrderNum = 10:     Global Const lconOrderDate = 11:    Global Const lconShipDate = 12:
Global Const lconParcelNum = 13

Global Const lconCustAddName = 14
Global Const lconCustAddLine1 = 15:    Global Const lconCustAddLine2 = 16: Global Const lconCustAddLine3 = 17
Global Const lconCustAddLine4 = 18:    Global Const lconCustAddLine5 = 18: Global Const lconCustAddPCode = 19

Global Const lconDeliverAddName = 21
Global Const lconDeliverAddLine1 = 22: Global Const lconDeliverAddLine2 = 23: Global Const lconDeliverAddLine3 = 24
Global Const lconDeliverAddLine4 = 25: Global Const lconDeliverAddLine5 = 26: Global Const lconDeliverAddPCode = 27

Global Const lconLiteNag1 = 28: Global Const lconLiteNag2 = 29

Global Const lconCompName = 30: Global Const lconCompAdd1 = 31
Global Const lconCompAdd2 = 32: Global Const lconCompAdd3 = 33
Global Const lconCompAdd4 = 34: Global Const lconCompAdd5 = 35

Global Const lconMediaInfo = 36

Global Const lconTitlDelivServ = 37: Global Const lconDelivServ = 38

Global Const lconTitlProdCode = 39:    Global Const lconTitlBinLoc = 40
Global Const lconTitlQty = 41:       Global Const lconTitlDespQty = 42
Global Const lconTitlDesc = 43:        Global Const lconTitlUnit = 44
Global Const lconTitlTax = 45:         Global Const lconTitlAccum = 46

'Detail Constants
Global Const lconPrdCatCode = 0:       Global Const lconPrdBinLoc = 1
Global Const lconPrdQtyOrd = 2:        Global Const lconPrdQtyDesp = 3
Global Const lconPrdDesc = 4:          Global Const lconPrdUnitPrice = 5
Global Const lconPrdTaxCode = 6:       Global Const lconPrdAmount = 7
Global Const lconPrdVatAmt = 8

'Footer Constants
Global Const lconTitleGoodsNVatTot = 0:    Global Const lconTitlePostNPack = 1
Global Const lconTitleDonation = 2:        Global Const lconTitleTotal = 3
Global Const lconTitleOverPay = 4:         Global Const lconTitleUnderPay = 5
Global Const lconTitleTotRefund = 6

Global Const lconGoodsNVatTot = 7:         Global Const lconPostNPack = 8
Global Const lconDonation = 9:             Global Const lconTotal = 10
Global Const lconOverPay = 11:             Global Const gconUnderPay = 12
Global Const lconTotRefund = 13

Global Const lconPayReceived = 14:         Global Const lconPay1 = 15
Global Const lconPay2 = 16

Global Const lconTitleTaxTopSummary = 17:  Global Const lconTitleTaxStnd = 18
Global Const lconTitleTaxStndPrcnt = 19:   Global Const lconTitleTaxZero = 20
Global Const lconTitleTaxZeroPrcnt = 21

Global Const lconTitleTaxTopGoodsNPP = 22: Global Const lconTitleTaxStndGoodsNPP = 23
Global Const lconTitleTaxZeroGoodsNPP = 24

Global Const lconTitleTaxTopVat = 25:      Global Const lconTitleTaxStndVat = 26
Global Const lconTitleTaxZeroVat = 27

Global Const lconStndNoteLn1 = 28:         Global Const lconStndNoteLn2 = 29
Global Const lconConsNoteLn1 = 30:             Global Const lconConsNoteLn2 = 31
Global Const lconCopyrightNote1 = 32: Global Const lconCopyrightNote2 = 33

'Advice Note Constants End
'Parcel Force Manifest Constants Start
Global Const lconTitlDateDesp = 0: Global Const lconDateDesp = 1: Global Const lconTitlPage = 2
Global Const lconPFPageNum = 3: Global Const lconTitleManifest = 4
Global Const lconTitlCompAddressName = 5: Global Const lconTitlCompAddress1 = 6: Global Const lconTitlCompAddress2 = 7
Global Const lconTitlCompAddress3 = 8: Global Const lconTitlCompAddress4 = 9: Global Const lconTitlCompAddress5 = 10

Global Const lconServDesc = 11: Global Const lconTitlContractNum = 12
Global Const lconContractNum = 13

Global Const lconTitlDetsConsNum1 = 14: Global Const lconTitlDetsSendrRef1 = 15: Global Const lconTitlDetsSpeHand = 16
Global Const lconTitlDetsConsNum2 = 17: Global Const lconTitlDetsDeliverName = 18: Global Const lconTitlDetsPostCode = 19
Global Const lconTitlDetsSendrRef2 = 20: Global Const lconTitlDetsItems = 21: Global Const lconTitlDetsSHS = 22
Global Const lconTitlDetsSHB = 23: Global Const lconTitlDetsSHP = 24

'Detail Constants
Global Const lconDetsConsNum = 0: Global Const lconDetsDeliverName = 1: Global Const lconDetsPostCode = 2
Global Const lconDetsSendrRef = 3: Global Const lconDetsItems = 4: Global Const lconDetsSHS = 5
Global Const lconDetsSHB = 6: Global Const lconDetsSHP = 7

'Footer Constants
Global Const lconTitlNumCons = 0
Global Const lconNumCons = 1
Global Const lconTitleNumItems = 2
Global Const lconNumItems = 3
'Parcel Force Manifest Constants End

'Refund Cheque Constants Start
'Header Constants
Global Const lconChqDate = 0
Global Const lconChqOrderNum = 1
Global Const lconChqName = 2
Global Const lconChqAmount = 3
Global Const lconChqTensOfThousands = 4
Global Const lconChqThousands = 5
Global Const lconChqHundreds = 6
Global Const lconChqTens = 7
Global Const lconChqUnits = 8
'Refund Cheque Constants End

'Batch Pick Constanst Start
'Header Constants
Global Const lconBPTitlBatchPick = 0
Global Const lconBPDate = 1
Global Const lconBPTitlOrderNums = 2
Global Const lconBPOrderNums = 3

Global Const lconBPTitlDetsCatCode = 4
Global Const lconBPTitlDetsBin = 5
Global Const lconBPTitlDetsQty = 6
Global Const lconBPTitlDetsProd = 7
Global Const lconBPTitlDetsWeight = 8

'Detail Constants
Global Const lconBPDetsCatCode = 0
Global Const lconBPDetsBin = 1
Global Const lconBPDetsQty = 2
Global Const lconBPDetsProd = 3
Global Const lconBPDetsWeight = 4

'Footer Constants
Global Const lconBPTotalWeight = 0
'Batch Pick Constanst End

'Credit Card Claim Start
'Header Constants
Global Const lconCCCHead1A = 0
Global Const lconCCCHead1B = 1
Global Const lconCCCTitlAsAt = 2
Global Const lconCCCAsAt = 3
Global Const lconCCCHead2A = 4
Global Const lconCCCTitlDetsCardNum = 5
Global Const lconCCCTitlDetsAmount = 6
Global Const lconCCCTitlDetsAuthCode = 7
Global Const lconCCCTitlDetsOrderNum = 8
Global Const lconCCCTitlDetsDespDate = 9
Global Const lconCCCTitlDetsCustomer = 10

'Detail Constants
Global Const lconCCCDetsCardNum = 0
Global Const lconCCCDetsAmount = 1
Global Const lconCCCDetsAuthCode = 2
Global Const lconCCCDetsOrderNum = 3
Global Const lconCCCDetsDespDate = 4
Global Const lconCCCDetsCustomerName = 5
Global Const lconCCCDetsCustomerLine1 = 6
Global Const lconCCCDetsCustomerLine2 = 7
Global Const lconCCCDetsCustomerLine3 = 8
Global Const lconCCCDetsCustomerLine4 = 9
Global Const lconCCCDetsCustomerLine5 = 10
Global Const lconCCCDetsCustomerPostCode = 11

'Footer Constants
'Const lcurPageTotal = 0
Global Const lconCCCTitlPageTotal = 0
Global Const lconCCCPageTotal = 1

Global Const lconCCCTitlPageNum = 2
Global Const lconCCCPageNum = 3

Global Const lconCCCTitlGrandTotal = 4
Global Const lconCCCGrandTotal = 5
'Credit Card Claim End

'Start Adams Advice
Global Const lconCompAddName = 0
Global Const lconCompAddLine1 = 1
Global Const lconCompAddLine2 = 2
Global Const lconCompAddLine3 = 3
Global Const lconCompAddLine4 = 4
Global Const lconCompAddLine5 = 5
Global Const lconCompAddPostCode = 6
Global Const lconCompAddTel = 7
Global Const lconCompAddFax = 8
Global Const lconSheetTitlName = 9
Global Const lconSheetTitlPage = 10
Global Const lconSheetPageNum = 11
Global Const lconCustAddrContactName = 12
Global Const lconCustAddrName = 13
Global Const lconCustAddrLine1 = 14
Global Const lconCustAddrLine2 = 15
Global Const lconCustAddrLine3 = 16
Global Const lconCustAddrLine4 = 17
Global Const lconCustAddrLine5 = 18
Global Const lconCustAddPostCode = 19
Global Const lconSheetTitlInvoiceNum = 20
Global Const lconSheetTitlTaxDate = 21
Global Const lconSheetTitlOrderNum = 22
Global Const lconSheetTitlAcctNum = 23
Global Const lconSheetInvoiceNum = 24
Global Const lconSheetTaxDate = 25
Global Const lconSheetOrderNum = 26
Global Const lconSheetAcctNum = 27
Global Const lconProdTitlProdCode = 28
Global Const lconProdTitlBinLoc = 29
Global Const lconProdTitlProdQty = 30
Global Const lconProdTitlDespQty = 31
Global Const lconProdTitlServDets = 32
Global Const lconProdTitlNetAmt = 33
Global Const lconProdTitlVatAmt = 34
Global Const lconProdTitlTaxCode = 35
Global Const lconProdTitlTitAmt = 36

Global Const lconProdServDets1 = 0
Global Const lconProdServDets2 = 1
Global Const lconProdServDets3 = 2
Global Const lconProdNetAmt = 3
Global Const lconProdVatAmt = 4

Global Const lconTotTitlNetAmt = 0
Global Const lconTotTitlVatAmt = 1
Global Const lconTotTitlCarriage = 2
Global Const lconTotTitlInvoice = 3
Global Const lconTotNetAmt = 4
Global Const lconTotVatAmt = 5
Global Const lconTotCarriage = 6
Global Const lconTotInvoice = 7
Global Const lconInvComment1 = 8
Global Const lconInvComment2 = 9
Global Const lconDelivLine1 = 10
Global Const lconDelivLine2 = 11
Global Const lconDelivLine3 = 12
Global Const lconDelivLine4 = 13
Global Const lconDelivLine5 = 14
Global Const lconDelivLine6 = 15
Global Const lconBanner1 = 16
Global Const lconBanner2 = 17

Const X = ""


Sub xSetLayout(pltReport As Long)
    
    Select Case pltReport
    Case ltAdviceNote, ltAdviceWithAddress
        With gstrReportLayout
            If pltReport = ltAdviceNote Then
                .strLayoutName = "Advice Note - pre-printed stationary"
                .lngLayoutType = ltAdviceNote
                .strLayoutFileName = "StdAdvice"
            Else
                .strLayoutName = "Advice Note - with company address"
                .lngLayoutType = ltAdviceNote
                .strLayoutFileName = "StdAdvAdd"
            End If
            .strPaperSize = "LISTINGA4"
            .strFontStandard.intSize = 10
            .strFontStandard.booBold = False
            .strFontStandard.strName = "Arial"
            .booHasDetails = True
            .booFooterBinded = False
        End With
        With gstrReportLayout.strHeaders(lconAdviceTitle):      .lngXPos = 1:    .lngYPos = 1:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 24: .strFontSpecific.booBold = True: End With
        Const lclInfTitlY = 3800
        Const lclInfDetlY = 4075 '4200
        Const lcCAddX = 1100
        Const lcDAddX = 5600
        
        With gstrReportLayout.strHeaders(lconCustAddName):      .lngXPos = lcCAddX: .lngYPos = 1600:   End With
        'Rest of address can be just printed after this.
        With gstrReportLayout.strHeaders(lconCustAddLine1):     .lngXPos = lcCAddX: .lngYPos = gconUnder: End With
        With gstrReportLayout.strHeaders(lconCustAddLine2):     .lngXPos = lcCAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconCustAddLine3):     .lngXPos = lcCAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconCustAddLine4):     .lngXPos = lcCAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconCustAddLine5):     .lngXPos = lcCAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconCustAddPCode):     .lngXPos = lcCAddX: .lngYPos = gconUnder:   End With
        
        With gstrReportLayout.strHeaders(lconTitleCustNum):     .lngXPos = 150:  .lngYPos = lclInfTitlY: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitleOrderNum):    .lngXPos = 2100: .lngYPos = lclInfTitlY: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitleOrderDate):   .lngXPos = 3750: .lngYPos = lclInfTitlY: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitleShipDate):    .lngXPos = 5250: .lngYPos = lclInfTitlY:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitleParcelNum):   .lngXPos = 6650: .lngYPos = lclInfTitlY:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlDelivServ):   .lngXPos = 8600: .lngYPos = lclInfTitlY:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlePageNum):     .lngXPos = 10700: .lngYPos = lclInfTitlY: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
                
        With gstrReportLayout.strHeaders(lconCustNum):          .lngXPos = 600:  .lngYPos = lclInfDetlY:   End With
        With gstrReportLayout.strHeaders(lconOrderNum):         .lngXPos = 2600: .lngYPos = lclInfDetlY:   End With
        With gstrReportLayout.strHeaders(lconOrderDate):        .lngXPos = 3750: .lngYPos = lclInfDetlY:   End With
        With gstrReportLayout.strHeaders(lconShipDate):         .lngXPos = 5100: .lngYPos = lclInfDetlY:   End With
        With gstrReportLayout.strHeaders(lconParcelNum):        .lngXPos = 7200: .lngYPos = lclInfDetlY:   End With
        With gstrReportLayout.strHeaders(lconDelivServ):        .lngXPos = 8400: .lngYPos = lclInfDetlY:  End With
        With gstrReportLayout.strHeaders(lconPageNum):          .lngXPos = 10800: .lngYPos = lclInfDetlY: End With

        With gstrReportLayout.strHeaders(lconDeliverTitle):     .lngXPos = 5300: .lngYPos = 1300:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddName):   .lngXPos = lcDAddX: .lngYPos = gstrReportLayout.strHeaders(lconCustAddName).lngYPos:  End With
        With gstrReportLayout.strHeaders(lconDeliverAddLine1):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddLine2):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddLine3):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddLine4):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddLine5):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconDeliverAddPCode):  .lngXPos = lcDAddX: .lngYPos = gconUnder:   End With
        
        With gstrReportLayout.strHeaders(lconLiteNag1):
            .lngXPos = 700:    .lngYPos = 900:
            .strFontSpecific.strName = "Arial":
            .strFontSpecific.intSize = 9:
            .strFontSpecific.booBold = True: End With
            
        With gstrReportLayout.strHeaders(lconLiteNag2):
            .lngXPos = 700:    .lngYPos = gconUnder:
            .strFontSpecific.strName = "Arial":
            .strFontSpecific.intSize = 9:
            .strFontSpecific.booBold = True: End With
        
        If pltReport = ltAdviceWithAddress Then
            With gstrReportLayout.strHeaders(lconCompName):         .lngXPos = 8500: .lngYPos = 1: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 11: .strFontSpecific.booBold = True: End With
            With gstrReportLayout.strHeaders(lconCompAdd1):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd2):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd3):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd4):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd5):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
        Else
            With gstrReportLayout.strHeaders(lconCompName):         .lngXPos = 8500: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 11: .strFontSpecific.booBold = True: End With
            With gstrReportLayout.strHeaders(lconCompAdd1):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd2):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd3):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd4):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
            With gstrReportLayout.strHeaders(lconCompAdd5):         .lngXPos = gstrReportLayout.strHeaders(lconCompName).lngXPos: .lngYPos = gconInvisible: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = False: End With
        End If
        
        With gstrReportLayout.strHeaders(lconMediaInfo):        .lngXPos = 3000: .lngYPos = 4375:   End With
        Const lcPrdTitlY = 4650
        With gstrReportLayout.strHeaders(lconTitlProdCode):     .lngXPos = 40:   .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlBinLoc):       .lngXPos = 1060: .lngYPos = lcPrdTitlY:     .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlQty):          .lngXPos = 2020: .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlDespQty):      .lngXPos = 2900: .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlDesc):         .lngXPos = 3850: .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlUnit):         .lngXPos = 8350: .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlTax):          .lngXPos = 9500: .lngYPos = lcPrdTitlY:      .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlAccum):        .lngXPos = 10600: .lngYPos = lcPrdTitlY:     .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        
        With gstrReportLayout.strDetails(lconPrdCatCode):       .lngXPos = gstrReportLayout.strHeaders(lconTitlProdCode).lngXPos:   .lngYPos = gconUnder:       End With
        With gstrReportLayout.strDetails(lconPrdBinLoc):        .lngXPos = gstrReportLayout.strHeaders(lconTitlBinLoc).lngXPos:     .lngYPos = gconSameAsPrev:  End With
        With gstrReportLayout.strDetails(lconPrdQtyOrd):        .lngXPos = gstrReportLayout.strHeaders(lconTitlQty).lngXPos:        .lngYPos = gconSameAsPrev: .lngAlignment = alCenter:            .intMaxChars = 6: End With
        With gstrReportLayout.strDetails(lconPrdQtyDesp):       .lngXPos = gstrReportLayout.strHeaders(lconTitlDespQty).lngXPos:    .lngYPos = gconSameAsPrev: .lngAlignment = alCenter:            .intMaxChars = 7: End With
        With gstrReportLayout.strDetails(lconPrdDesc):          .lngXPos = gstrReportLayout.strHeaders(lconTitlDesc).lngXPos:       .lngYPos = gconSameAsPrev: End With

        With gstrReportLayout.strDetails(lconPrdUnitPrice):     .lngXPos = gstrReportLayout.strHeaders(lconTitlUnit).lngXPos:       .lngYPos = gconSameAsPrev:   .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strDetails(lconPrdTaxCode):       .lngXPos = gstrReportLayout.strHeaders(lconTitlTax).lngXPos:        .lngYPos = gconSameAsPrev: .lngAlignment = alCenter:            .intMaxChars = 7: End With
        With gstrReportLayout.strDetails(lconPrdAmount):        .lngXPos = 10400:      .lngYPos = gconSameAsPrev:  .lngAlignment = alRight:            .intMaxChars = 9: End With                                    
        'Y = zero, will be determined by gstrReport.booFooterBinded
        With gstrReportLayout.strFooters(lconTitleGoodsNVatTot): .lngXPos = gstrReportLayout.strHeaders(lconTitlUnit).lngXPos: .lngYPos = 0:  End With
        With gstrReportLayout.strFooters(lconTitlePostNPack):   .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder:  End With
        With gstrReportLayout.strFooters(lconTitleDonation):    .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder1Space:  End With
        With gstrReportLayout.strFooters(lconTitleTotal):       .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder1Space: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strFooters(lconTitleOverPay):     .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder1Space:  End With
        With gstrReportLayout.strFooters(lconTitleUnderPay):    .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder:  End With
        With gstrReportLayout.strFooters(lconTitleTotRefund):   .lngXPos = gstrReportLayout.strFooters(lconTitleGoodsNVatTot).lngXPos: .lngYPos = gconUnder: End With
                
        With gstrReportLayout.strFooters(lconGoodsNVatTot):     .lngXPos = 10400: .lngYPos = 0:           .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(lconPostNPack):        .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(lconDonation):         .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder1Space:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(lconTotal):            .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder1Space:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(lconOverPay):          .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder1Space:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(gconUnderPay):         .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
        With gstrReportLayout.strFooters(lconTotRefund):        .lngXPos = gstrReportLayout.strFooters(lconGoodsNVatTot).lngXPos:            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 9:        End With
                
        With gstrReportLayout.strFooters(lconPayReceived):      .lngXPos = 40: .lngYPos = 0:  End With
        With gstrReportLayout.strFooters(lconPay1):             .lngXPos = gstrReportLayout.strFooters(lconPayReceived).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconPay2):             .lngXPos = gstrReportLayout.strFooters(lconPayReceived).lngXPos: .lngYPos = gconUnder:   End With
        
        With gstrReportLayout.strFooters(lconTitleTaxTopSummary): .lngXPos = 3120:            .lngYPos = 800: End With
        With gstrReportLayout.strFooters(lconTitleTaxStnd):     .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos:            .lngYPos = gconUnder: End With
        With gstrReportLayout.strFooters(lconTitleTaxStndPrcnt): .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos + (1200 / gintScaleFactor):           .lngYPos = gconSameAsPrev:            .lngAlignment = alRight:            .intMaxChars = 6:        End With
        With gstrReportLayout.strFooters(lconTitleTaxZero):     .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos:            .lngYPos = gconUnder: End With
        With gstrReportLayout.strFooters(lconTitleTaxZeroPrcnt): .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos + (1200 / gintScaleFactor):            .lngYPos = gconSameAsPrev:            .lngAlignment = alRight:            .intMaxChars = 6:        End With
        With gstrReportLayout.strFooters(lconTitleTaxTopGoodsNPP): .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos + 2000: .lngYPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngYPos: End With
        With gstrReportLayout.strFooters(lconTitleTaxStndGoodsNPP): .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopGoodsNPP).lngXPos + (300 / gintScaleFactor):           .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 8:        End With
        With gstrReportLayout.strFooters(lconTitleTaxZeroGoodsNPP): .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopGoodsNPP).lngXPos + (300 / gintScaleFactor):            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 8:        End With

        With gstrReportLayout.strFooters(lconTitleTaxTopVat):   .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngXPos + 3700: .lngYPos = gstrReportLayout.strFooters(lconTitleTaxTopSummary).lngYPos: End With
        With gstrReportLayout.strFooters(lconTitleTaxStndVat):  .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopVat).lngXPos:            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 8:        End With
        With gstrReportLayout.strFooters(lconTitleTaxZeroVat):  .lngXPos = gstrReportLayout.strFooters(lconTitleTaxTopVat).lngXPos:            .lngYPos = gconUnder:            .lngAlignment = alRight:            .intMaxChars = 8:        End With
             
        With gstrReportLayout.strFooters(lconStndNoteLn1)
            .lngXPos = 10
            .lngYPos = gconUnder1Space
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9
        End With
        
        With gstrReportLayout.strFooters(lconStndNoteLn2)
            .lngXPos = gstrReportLayout.strFooters(lconStndNoteLn1).lngXPos
            .lngYPos = gconUnder
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9
        End With
        
        With gstrReportLayout.strFooters(lconConsNoteLn1)
            .lngXPos = 50
            .lngYPos = 2500
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9:
        End With
        
        With gstrReportLayout.strFooters(lconConsNoteLn2)
            .lngXPos = gconAfterNSpace
            .lngYPos = gconSameAsPrev
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9
        End With
        
        With gstrReportLayout.strFooters(lconCopyrightNote1)
            .lngXPos = 200
            .lngYPos = 2800
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 8: .strFontSpecific.booBold = True
        End With
        
        With gstrReportLayout.strFooters(lconCopyrightNote2)
            .lngXPos = gconAfter
            .lngYPos = gconSameAsPrev
            .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 8: .strFontSpecific.booBold = True
        End With
        
        Dim lintBoxInc As Integer
        ReDim gstrBoxArray(18)
        'Box around product titles
        Const lcProdTitlBoxTop = 4625
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = lcProdTitlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 11500
            .lngYPos2 = 4900
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = lcProdTitlBoxTop
            .lngXPos2 = 975
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Quantity ordered Column
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 1905
            .lngYPos = lcProdTitlBoxTop
            .lngXPos2 = 2835
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'product description column
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 3735
            .lngYPos = lcProdTitlBoxTop
            .lngXPos2 = 8070
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'Right hand line of product box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(0).lngXPos2
            .lngYPos = lcProdTitlBoxTop
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'Line At bottom of product box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(0).lngXPos
            .booRelativeY1 = True
            .lngYPos = 0
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'totals box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(3).lngXPos2
            .booRelativeY1 = True
            .lngYPos = 0
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 2500
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'totals box line above donation
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(3).lngXPos2
            .booRelativeY1 = True
            .lngYPos = 700
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = gstrBoxArray(8).lngYPos
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'totals box line above total
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(3).lngXPos2
            .booRelativeY1 = True
            .lngYPos = 1100
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = gstrBoxArray(8).lngYPos
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'totals box line above over payment
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(3).lngXPos2
            .booRelativeY1 = True
            .lngYPos = 1600
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = gstrBoxArray(8).lngYPos
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1

        'Tax Summary Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 2900
            .booRelativeY1 = True
            .lngYPos = 800
            .lngXPos2 = gstrBoxArray(3).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = gstrBoxArray(9).lngYPos
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
                
        'Payments box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .booRelativeY1 = True
            .lngXPos2 = gstrBoxArray(3).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 800
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        Const lcDetlBoxTop = 3800
        Const lcDetlBoxBot = 4350
        'Box around Custnum to parcel num
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = lcDetlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 11500
            .lngYPos2 = lcDetlBoxBot
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Box Around Order Num
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 1900 '2000
            .lngYPos = lcDetlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 3500
            .lngYPos2 = lcDetlBoxBot
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Box around Ship Date
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 4900
            .lngYPos = lcDetlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 6400
            .lngYPos2 = lcDetlBoxBot
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Delivery Service Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 8200
            .lngYPos = lcDetlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 10300
            .lngYPos2 = lcDetlBoxBot
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Box around detail titles
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = lcDetlBoxTop
            .booRelativeY1 = False
            .lngXPos2 = 11500
            .lngYPos2 = 4075
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Media code box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = lcDetlBoxBot
            .booRelativeY1 = False
            .lngXPos2 = 11500
            .lngYPos2 = 4625
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Nag Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 650
            .lngYPos = 850
            .booRelativeY1 = False
            .lngXPos2 = 4725
            .lngYPos2 = 1400
            .booRelativeY2 = False
            .strBoxStyle = "D"
        End With: lintBoxInc = lintBoxInc + 1
                
    Case ltParcelForceManifest
        With gstrReportLayout
            .strLayoutName = "Parcel Force Manifest"
            .lngLayoutType = ltParcelForceManifest
            .strLayoutFileName = "PFManifes"
            .strPaperSize = "NORMALA4"
            .strFontStandard.intSize = 11
            .strFontStandard.booBold = False
            .strFontStandard.strName = "Arial"
            .booHasDetails = True
            .booFooterBinded = False
        End With
        With gstrReportLayout.strHeaders(lconTitlDateDesp):         .lngXPos = 50: .lngYPos = 50:   End With
        With gstrReportLayout.strHeaders(lconDateDesp):             .lngXPos = 2000: .lngYPos = gconSameAsPrev:  End With
        With gstrReportLayout.strHeaders(lconTitlPage):             .lngXPos = 8500: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconPFPageNum):            .lngXPos = 9200: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconTitleManifest):        .lngXPos = 4000: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconTitlCompAddressName):   .lngXPos = 1000: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strHeaders(lconTitlCompAddress1):      .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:  End With
        With gstrReportLayout.strHeaders(lconTitlCompAddress2):      .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconTitlCompAddress3):      .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconTitlCompAddress4):      .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconTitlCompAddress5):      .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        
        With gstrReportLayout.strHeaders(lconServDesc):             .lngXPos = 3500: .lngYPos = gconUnder1Space:   End With
        With gstrReportLayout.strHeaders(lconTitlContractNum):      .lngXPos = 50: .lngYPos = gconUnder1Space:   End With
        With gstrReportLayout.strHeaders(lconContractNum):          .lngXPos = 1800: .lngYPos = gconSameAsPrev:   End With

        With gstrReportLayout.strHeaders(lconTitlDetsConsNum1):     .lngXPos = 50: .lngYPos = gconUnder1Space: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsSendrRef1):    .lngXPos = 7300: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsSpeHand):      .lngXPos = 8800: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsConsNum2):     .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsConsNum1).lngXPos: .lngYPos = gconUnder: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsDeliverName):  .lngXPos = 2000: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsPostCode):     .lngXPos = 5000: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsSendrRef2):    .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSendrRef1).lngXPos: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsItems):        .lngXPos = 6500: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strHeaders(lconTitlDetsSHS):          .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSpeHand).lngXPos: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlDetsSHB):          .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSHS).lngXPos + 300: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconTitlDetsSHP):          .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSHB).lngXPos + 300: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With

        With gstrReportLayout.strDetails(lconDetsConsNum):          .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsConsNum1).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strDetails(lconDetsDeliverName):      .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsDeliverName).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsPostCode):         .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsPostCode).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsSendrRef):         .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSendrRef1).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsItems):            .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsItems).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsSHS):              .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSHS).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsSHB):              .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSHB).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconDetsSHP):              .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSHP).lngXPos: .lngYPos = gconSameAsPrev:   End With

        With gstrReportLayout.strFooters(lconTitlNumCons):          .lngXPos = 4000: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconNumCons):              .lngXPos = 6700: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strFooters(lconTitleNumItems):        .lngXPos = 4000: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconNumItems):             .lngXPos = 6700: .lngYPos = gconSameAsPrev:   End With
            
        ReDim gstrBoxArray(5)
        lintBoxInc = 0
        
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = 2000
            .booRelativeY1 = False
            .lngXPos2 = 9700
            .lngYPos2 = 2500
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Consign column
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = gstrBoxArray(0).lngYPos
            .lngXPos2 = gstrReportLayout.strHeaders(lconTitlDetsDeliverName).lngXPos - 200
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'postcode column
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsPostCode).lngXPos - 200
            .lngYPos = gstrBoxArray(0).lngYPos
            .lngXPos2 = gstrReportLayout.strHeaders(lconTitlDetsItems).lngXPos - 200
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Sender Ref column
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrReportLayout.strHeaders(lconTitlDetsSendrRef1).lngXPos - 200
            .lngYPos = gstrBoxArray(0).lngYPos
            .lngXPos2 = gstrReportLayout.strHeaders(lconTitlDetsSpeHand).lngXPos - 200
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Right Line
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrBoxArray(0).lngXPos2
            .lngYPos = gstrBoxArray(0).lngYPos
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Bottom Line
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .booRelativeY1 = True
            .lngYPos = 0
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .booRelativeY2 = True
            .lngYPos2 = 0
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
    Case ltRefundCheques
        With gstrReportLayout
            .strLayoutName = "Company cheque"
            .lngLayoutType = ltRefundCheques
            .strLayoutFileName = "StdCheque"
            .strPaperSize = "VH101A297"
            .strFontStandard.intSize = 12
            .strFontStandard.booBold = False
            .strFontStandard.strName = "Arial"
            .booHasDetails = False
            .booFooterBinded = False
        End With
        With gstrReportLayout.strHeaders(lconChqDate):              .lngXPos = 8200: .lngYPos = 2000:   End With
        With gstrReportLayout.strHeaders(lconChqOrderNum):          .lngXPos = 100: .lngYPos = 2650:   End With
        With gstrReportLayout.strHeaders(lconChqName):              .lngXPos = 1350: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconChqAmount):            .lngXPos = 8500: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconChqTensOfThousands):   .lngXPos = 50: .lngYPos = 3900:   End With
        With gstrReportLayout.strHeaders(lconChqThousands):         .lngXPos = 950: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconChqHundreds):          .lngXPos = 1850: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconChqTens):              .lngXPos = 2750: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconChqUnits):             .lngXPos = 3650: .lngYPos = gconSameAsPrev:   End With
    Case ltBatchPickings
        With gstrReportLayout
            .strLayoutName = "Batch picking"
            .lngLayoutType = ltBatchPickings
            .strLayoutFileName = "StdBatPic"
            .strPaperSize = "NORMALA4"
            .strFontStandard.intSize = 10
            .strFontStandard.booBold = False
            .strFontStandard.strName = "Arial"
            .booHasDetails = True
            .booFooterBinded = False
        End With

        With gstrReportLayout.strHeaders(lconBPTitlBatchPick):      .lngXPos = 50: .lngYPos = 50:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 20: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconBPDate):               .lngXPos = 8000: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconBPTitlOrderNums):      .lngXPos = 50: .lngYPos = 1000:   End With
        With gstrReportLayout.strHeaders(lconBPOrderNums):          .lngXPos = 2000: .lngYPos = gconSameAsPrev:   End With

        With gstrReportLayout.strHeaders(lconBPTitlDetsCatCode):    .lngXPos = 50: .lngYPos = 1500:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconBPTitlDetsBin):        .lngXPos = 2000: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconBPTitlDetsQty):        .lngXPos = 4000: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconBPTitlDetsProd):       .lngXPos = 5500: .lngYPos = gconSameAsPrev:  .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconBPTitlDetsWeight):     .lngXPos = 10000: .lngYPos = gconSameAsPrev: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With

        With gstrReportLayout.strDetails(lconBPDetsCatCode):        .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsCatCode).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strDetails(lconBPDetsBin):            .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsBin).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconBPDetsQty):            .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsQty).lngXPos: .lngYPos = gconSameAsPrev:  .lngAlignment = alCenter:            .intMaxChars = 12: End With
        With gstrReportLayout.strDetails(lconBPDetsProd):           .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsProd).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconBPDetsWeight):         .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsWeight).lngXPos: .lngYPos = gconSameAsPrev:  .lngAlignment = alRight:            .intMaxChars = 12: End With

        With gstrReportLayout.strFooters(lconBPTotalWeight):        .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsWeight).lngXPos: .lngYPos = gconUnder: .lngAlignment = alRight:            .intMaxChars = 12:  End With
        
        ReDim gstrBoxArray(5)
        lintBoxInc = 0
        
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = gstrReportLayout.strHeaders(lconBPTitlDetsCatCode).lngYPos
            .booRelativeY1 = False
            .lngXPos2 = 11500
            .lngYPos2 = 1800
            .booRelativeY2 = False
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Cat Code Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = gstrReportLayout.strHeaders(lconBPTitlDetsCatCode).lngYPos
            .booRelativeY1 = False
            .lngXPos2 = gstrReportLayout.strHeaders(lconBPTitlDetsBin).lngXPos - 100
            .lngYPos2 = 0
            .booRelativeY2 = True
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Qty Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsQty).lngXPos - 100
            .lngYPos = gstrBoxArray(1).lngYPos
            .booRelativeY1 = False
            .lngXPos2 = gstrReportLayout.strHeaders(lconBPTitlDetsProd).lngXPos - 100
            .lngYPos2 = 0
            .booRelativeY2 = True
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Weight Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsWeight).lngXPos - 100
            .lngYPos = gstrBoxArray(1).lngYPos
            .booRelativeY1 = False
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .lngYPos2 = 0
            .booRelativeY2 = True
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        
        'Bottom Line
        With gstrBoxArray(lintBoxInc)
            .lngXPos = 0
            .lngYPos = 0
            .booRelativeY1 = True
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .lngYPos2 = 0
            .booRelativeY2 = True
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
        
        'Total Box
        With gstrBoxArray(lintBoxInc)
            .lngXPos = gstrReportLayout.strHeaders(lconBPTitlDetsWeight).lngXPos - 100
            .lngYPos = 0
            .booRelativeY1 = True
            .lngXPos2 = gstrBoxArray(0).lngXPos2
            .lngYPos2 = 300
            .booRelativeY2 = True
            .strBoxStyle = "B"
        End With: lintBoxInc = lintBoxInc + 1
    Case ltCreditCardClaims
        With gstrReportLayout
            .strLayoutName = "HSBC Card Serv"
            .lngLayoutType = ltCreditCardClaims
            .strLayoutFileName = "HSCardSer"
            .strPaperSize = "LISTINGA3"
            .strFontStandard.intSize = 12
            .strFontStandard.booBold = True
            .strFontStandard.strName = "Courier New"
            .booHasDetails = True
            .booFooterBinded = False
        End With

        With gstrReportLayout.strHeaders(lconCCCHead1A):            .lngXPos = 50: .lngYPos = 50:   End With
        With gstrReportLayout.strHeaders(lconCCCHead1B):            .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconCCCTitlAsAt):          .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconCCCAsAt):              .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconCCCHead2A):            .lngXPos = 50: .lngYPos = gconUnder:  End With
        
        With gstrReportLayout.strHeaders(lconCCCTitlDetsCardNum):   .lngXPos = 50: .lngYPos = gconUnder1Space:   End With
        With gstrReportLayout.strHeaders(lconCCCTitlDetsAmount):    .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsCardNum).lngXPos + 3000: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconCCCTitlDetsAuthCode):  .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAmount).lngXPos + 1200: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strHeaders(lconCCCTitlDetsOrderNum):  .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAuthCode).lngXPos + 1600: .lngYPos = gconSameAsPrev:  End With
        With gstrReportLayout.strHeaders(lconCCCTitlDetsDespDate):  .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsOrderNum).lngXPos + 1500: .lngYPos = gconSameAsPrev:    End With
        With gstrReportLayout.strHeaders(lconCCCTitlDetsCustomer):  .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsDespDate).lngXPos + 1600: .lngYPos = gconSameAsPrev:    End With

'Detail
        With gstrReportLayout.strDetails(lconCCCDetsCardNum):       .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsCardNum).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strDetails(lconCCCDetsAmount):        .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAmount).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsAuthCode):      .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAuthCode).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsOrderNum):      .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsOrderNum).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsDespDate):      .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsDespDate).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerName):  .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsCustomer).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerLine1): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerLine2): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerLine3): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerLine4): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerLine5): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strDetails(lconCCCDetsCustomerPostCode): .lngXPos = gconAfterNSpace: .lngYPos = gconSameAsPrev:  End With
        
'Footer
        With gstrReportLayout.strFooters(lconCCCTitlPageTotal):     .lngXPos = 50: .lngYPos = gconUnder1Space:   End With
        With gstrReportLayout.strFooters(lconCCCPageTotal):         .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAmount).lngXPos: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strFooters(lconCCCTitlPageNum):       .lngXPos = 5000: .lngYPos = gconUnder1Space:   End With
        With gstrReportLayout.strFooters(lconCCCPageNum):           .lngXPos = 6500: .lngYPos = gconSameAsPrev:   End With
        With gstrReportLayout.strFooters(lconCCCTitlGrandTotal):    .lngXPos = 50: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconCCCGrandTotal):        .lngXPos = gstrReportLayout.strHeaders(lconCCCTitlDetsAmount).lngXPos: .lngYPos = gconSameAsPrev:   End With
        
    Case ltInvoice
        With gstrReportLayout
            .strLayoutName = "Adams Invoice"
            .lngLayoutType = ltInvoice
            .strLayoutFileName = "AdamInvoi"
            .strPaperSize = "NORMALA4"
            .strFontStandard.intSize = 10
            .strFontStandard.booBold = False
            .strFontStandard.strName = "Arial"
            .booHasDetails = True
            .booFooterBinded = True
        End With
        With gstrReportLayout.strHeaders(lconCompAddName):      .lngXPos = 750: .lngYPos = 1000:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddLine1):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddLine2):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddLine3):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddLine4):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddLine5):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddPostCode):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddTel):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCompAddFax):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        
        With gstrReportLayout.strHeaders(lconSheetTitlName):     .lngXPos = 7000: .lngYPos = 1925:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 14: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconSheetTitlPage):     .lngXPos = 9000: .lngYPos = gconSameAsPrev:   End With
        
        With gstrReportLayout.strHeaders(lconSheetPageNum):     .lngXPos = 9700: .lngYPos = gconSameAsPrev:   End With
        
        With gstrReportLayout.strHeaders(lconCustAddrContactName):     .lngXPos = gstrReportLayout.strHeaders(lconCompAddName).lngXPos + 400: .lngYPos = 3900: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddName):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddLine1):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddLine2):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddLine3):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddLine4):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddLine5):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconCustAddPostCode):     .lngXPos = gstrReportLayout.strHeaders(lconCustAddrContactName).lngXPos: .lngYPos = gconUnder:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 10: .strFontSpecific.booBold = True: End With
        
        With gstrReportLayout.strHeaders(lconSheetInvoiceNum):     .lngXPos = gstrReportLayout.strHeaders(lconSheetTitlPage).lngXPos: .lngYPos = 3850:   End With
        With gstrReportLayout.strHeaders(lconSheetTaxDate):     .lngXPos = gstrReportLayout.strHeaders(lconSheetInvoiceNum).lngXPos: .lngYPos = 4350:   End With
        With gstrReportLayout.strHeaders(lconSheetOrderNum):     .lngXPos = gstrReportLayout.strHeaders(lconSheetInvoiceNum).lngXPos: .lngYPos = 4800:   End With
        With gstrReportLayout.strHeaders(lconSheetAcctNum):     .lngXPos = gstrReportLayout.strHeaders(lconSheetInvoiceNum).lngXPos: .lngYPos = 5300:   End With
        
        With gstrReportLayout.strHeaders(lconProdTitlServDets):     .lngXPos = 600: .lngYPos = 6380:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strHeaders(lconProdTitlNetAmt):     .lngXPos = 7400: .lngYPos = gconSameAsPrev:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strHeaders(lconProdTitlVatAmt):     .lngXPos = 9450: .lngYPos = gconSameAsPrev:   .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: .lngAlignment = alRight:            .intMaxChars = 9: End With

        With gstrReportLayout.strDetails(lconProdServDets1):
        .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos:   .lngYPos = gconUnder:        End With
        
        With gstrReportLayout.strDetails(lconProdServDets2):
        .lngXPos = gconAfter:   .lngYPos = gconSameAsPrev:        End With
        
        With gstrReportLayout.strDetails(lconProdServDets3):
        .lngXPos = gconAfter:   .lngYPos = gconSameAsPrev:        End With
        
        With gstrReportLayout.strDetails(lconProdNetAmt):
        .lngXPos = gstrReportLayout.strHeaders(lconProdTitlNetAmt).lngXPos:   .lngYPos = gconSameAsPrev:       .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strDetails(lconProdVatAmt):
        .lngXPos = gstrReportLayout.strHeaders(lconProdTitlVatAmt).lngXPos:   .lngYPos = gconSameAsPrev:       .lngAlignment = alRight:            .intMaxChars = 9: End With
        
        With gstrReportLayout.strFooters(lconTotTitlNetAmt):     .lngXPos = 7050: .lngYPos = 13700:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strFooters(lconTotTitlVatAmt):     .lngXPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos + 500: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With

        With gstrReportLayout.strFooters(lconTotTitlVatAmt):     .lngXPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos + 500: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True: End With
        With gstrReportLayout.strFooters(lconTotTitlCarriage):   .lngXPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos + 1000: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        With gstrReportLayout.strFooters(lconTotTitlInvoice):    .lngXPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos + 1500: .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 9: .strFontSpecific.booBold = True:  End With
        
        With gstrReportLayout.strFooters(lconTotNetAmt):        .lngXPos = 9500: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos: .lngAlignment = alRight:             .intMaxChars = 9: End With
        With gstrReportLayout.strFooters(lconTotVatAmt):        .lngXPos = gstrReportLayout.strFooters(lconTotNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlVatAmt).lngYPos:   .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strFooters(lconTotCarriage):        .lngXPos = gstrReportLayout.strFooters(lconTotNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlCarriage).lngYPos:   .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strFooters(lconTotInvoice):        .lngXPos = gstrReportLayout.strFooters(lconTotNetAmt).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlInvoice).lngYPos:   .lngAlignment = alRight:            .intMaxChars = 9: End With
        With gstrReportLayout.strFooters(lconInvComment1):        .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos - 1050: End With
        With gstrReportLayout.strFooters(lconInvComment2):        .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder: End With
        
        With gstrReportLayout.strFooters(lconDelivLine1):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gstrReportLayout.strFooters(lconTotTitlNetAmt).lngYPos + 350:   End With
        With gstrReportLayout.strFooters(lconDelivLine2):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder:: End With
        With gstrReportLayout.strFooters(lconDelivLine3):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconDelivLine4):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconDelivLine5):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder:   End With
        With gstrReportLayout.strFooters(lconDelivLine6):     .lngXPos = gstrReportLayout.strHeaders(lconProdTitlServDets).lngXPos: .lngYPos = gconUnder:   End With

        With gstrReportLayout.strFooters(lconBanner1):     .lngXPos = 700: .lngYPos = 15650:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 8: End With
        With gstrReportLayout.strFooters(lconBanner2):     .lngXPos = 3000: .lngYPos = 15650:    .strFontSpecific.strName = "Arial": .strFontSpecific.intSize = 8: End With
    Case ltHeaderReport

        Dim lintArrInc As Integer
        For lintArrInc = 0 To UBound(gstrReportLayout.strHeaders)
            With gstrReportLayout.strHeaders(lintArrInc)
                .lngYPos = gconUnder
                .lngXPos = 700
            End With
        Next lintArrInc
    End Select
    
End Sub
