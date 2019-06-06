Attribute VB_Name = "modDataDeclarations"
Option Explicit

Type CustomerAccount
    lngCustNum          As Long
    strSalutation       As String * 15
    strSurname          As String * 25
    strInitials         As String * 20
    strAdd1             As String * 30
    strAdd2             As String * 30
    strAdd3             As String * 30
    strAdd4             As String * 30
    strAdd5             As String * 30
    strPostcode         As String * 9
    strTelephoneNum     As String * 25
    strDeliverySalutation       As String * 15
    strDeliverySurname          As String * 25
    strDeliveryInitials         As String * 20
    strDeliveryAdd1     As String * 30
    strDeliveryAdd2     As String * 30
    strDeliveryAdd3     As String * 30
    strDeliveryAdd4     As String * 30
    strDeliveryAdd5     As String * 30
    strDeliveryPostcode As String * 9
    strAccountType      As String * 1
    strReceiveMailings  As String * 1
    strAcctInUseByFlag  As String * 10
    strAccountStatus    As String * 10
    strEveTelephoneNum  As String * 25
    strEmail            As String * 50
End Type
Global gstrCustomerAccount As CustomerAccount

Type AdviceNoteOrder
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
    strTelephoneNum     As String * 25
    strDeliverySalutation       As String * 15
    strDeliverySurname          As String * 25
    strDeliveryInitials         As String * 20
    strDeliveryAdd1     As String * 30
    strDeliveryAdd2     As String * 30
    strDeliveryAdd3     As String * 30
    strDeliveryAdd4     As String * 30
    strDeliveryAdd5     As String * 30
    strDeliveryPostcode As String * 9
    strOrderStyle       As String * 10
    strMediaCode        As String * 10
    datDeliveryDate     As Date
    strCourierCode      As String * 10
    strPaymentType1     As String * 10
    strPaymentType2     As String * 10
    strOrderCode        As String * 1
    strCardName         As String * 50
    strCardNumber       As String * 20
    strAuthorisationCode As String * 10
    datExpiryDate       As Date
    strDonation         As String * 8
    strPayment          As String * 8
    strPayment2         As String * 8
    strUnderpayment     As String * 8
    strReconcilliation  As String * 8
    strOOSRefund        As String * 8 ' Not Used
    datBankRepPrintDate As Date
    lngAdviceRemarkNum  As Long
    lngConsignRemarkNum As Long
    strAcctInUseByFlag  As String * 10
    strTotalIncVat      As String * 8
    strVAT              As String * 8
    strPostage          As String * 8
    strOverSeasFlag     As String * 1
    datDespatchDate     As Date
    datCreationDate     As Date
    intNumOfParcels     As Integer
    lngGrossWeight      As Long
    strCardType         As String * 10
    lngIssueNumber      As Long
    lngStockBatchNum    As Long
    datCardStartDate    As Date
    strDenom            As String * 1
End Type
Global gstrAdviceNoteOrder As AdviceNoteOrder

Type Remarks
    lngRemarkNumber     As Long
    strType             As String * 10
    strText             As String * 255
End Type

Type DespQtyTot
    strs As String
End Type

Global gstrConsignmentNote As Remarks
Global gstrInternalNote As Remarks

Global gstrOrderTotal As String
Global gstrVatTotal As String

Global gstrVATRate As String

Global gstrPaymentTypeCode() As String

Type ChequePrint
    lngOrderNum As Long
    lngChequeNum As Long
End Type
Global glngChequeOrderNumPrinted() As ChequePrint

Global glngStockBatchNumber As Long
Global gdatLastStockBatchNumberDate As Date


Global gstrOrderEntryOrderStatus As String * 1
