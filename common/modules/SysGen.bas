Attribute VB_Name = "modSysGen"
Option Explicit
Sub SetDBData(pstrTableAndFields() As TableAndFields)
    ReDim pstrTableAndFields(252)
 
    'AdviceNotes
    With pstrTableAndFields(0): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(1): .strType = "4":     .strName = "OrderNum":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(2): .strType = "10":     .strName = "OrderStatus":    .strSize = "1":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(3): .strType = "10":     .strName = "OrderStyle":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(4): .strType = "10":     .strName = "CallerSalutation":    .strSize = "15":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(5): .strType = "10":     .strName = "CallerSurname":    .strSize = "25":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(6): .strType = "10":     .strName = "CallerInitials":    .strSize = "20":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(7): .strType = "10":     .strName = "AdviceAdd1":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(8): .strType = "10":     .strName = "AdviceAdd2":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(9): .strType = "10":     .strName = "AdviceAdd3":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(10): .strType = "10":     .strName = "AdviceAdd4":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(11): .strType = "10":     .strName = "AdviceAdd5":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(12): .strType = "10":     .strName = "AdvicePostcode":    .strSize = "9":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(13): .strType = "10":     .strName = "TelephoneNum":    .strSize = "25":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(14): .strType = "10":     .strName = "DeliveryAdd1":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(15): .strType = "10":     .strName = "DeliveryAdd2":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(16): .strType = "10":     .strName = "DeliveryAdd3":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(17): .strType = "10":     .strName = "DeliveryAdd4":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(18): .strType = "10":     .strName = "DeliveryAdd5":    .strSize = "30":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(19): .strType = "10":     .strName = "DeliveryPostcode":    .strSize = "9":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(20): .strType = "10":     .strName = "MediaCode":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(21): .strType = "8":     .strName = "DeliveryDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(22): .strType = "10":     .strName = "CourierCode":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(23): .strType = "10":     .strName = "OrderType":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(24): .strType = "10":     .strName = "PaymentType2":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(25): .strType = "10":     .strName = "OrderCode":    .strSize = "1":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(26): .strType = "10":     .strName = "CardNumber":    .strSize = "50":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(27): .strType = "8":     .strName = "ExpiryDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(28): .strType = "5":     .strName = "Donation":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(29): .strType = "5":     .strName = "Payment":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(30): .strType = "5":     .strName = "Payment2":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(31): .strType = "5":     .strName = "Underpayment":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(32): .strType = "5":     .strName = "Reconcilliation":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(33): .strType = "5":     .strName = "Postage":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(34): .strType = "5":     .strName = "Vat":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(35): .strType = "5":     .strName = "TotalIncVat":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(36): .strType = "5":     .strName = "OOSRefund":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(37): .strType = "4":     .strName = "AdviceRemarkNum":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(38): .strType = "4":     .strName = "ConsignRemarkNum":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(39): .strType = "10":     .strName = "OverSeasFlag":    .strSize = "1":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(40): .strType = "10":     .strName = "ProcessedBy":    .strSize = "50":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(41): .strType = "10":     .strName = "LockingFlag":    .strSize = "50":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(42): .strType = "8":     .strName = "CreationDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "Now()":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(43): .strType = "10":     .strName = "AuthorisationCode":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(44): .strType = "8":     .strName = "DespatchDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(45): .strType = "8":     .strName = "BankRepPrintDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(46): .strType = "10":     .strName = "CardName":    .strSize = "50":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(47): .strType = "10":     .strName = "DeliverySalutation":    .strSize = "15":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(48): .strType = "10":     .strName = "DeliverySurname":    .strSize = "25":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(49): .strType = "10":     .strName = "DeliveryInitials":    .strSize = "20":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(50): .strType = "3":     .strName = "NumOfParcels":    .strSize = "2":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(51): .strType = "4":     .strName = "GrossWeight":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(52): .strType = "10":     .strName = "CardType":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(53): .strType = "4":     .strName = "CardIssueNumber":    .strSize = "4":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(54): .strType = "10":     .strName = "StockBatchNum":    .strSize = "10":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(55): .strType = "1":     .strName = "PickPrinted":    .strSize = "1":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(56): .strType = "8":     .strName = "CardStartDate":    .strSize = "8":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(57): .strType = "10":     .strName = "Denom":    .strSize = "1":    .strSourceTable = gtblAdviceNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'CashBook
    With pstrTableAndFields(58): .strType = "10":     .strName = "ChequeNum":    .strSize = "10":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(59): .strType = "8":     .strName = "PrintedDate":    .strSize = "8":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(60): .strType = "5":     .strName = "Amount":    .strSize = "8":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(61): .strType = "10":     .strName = "Name":    .strSize = "60":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(62): .strType = "8":     .strName = "ClearedDate":    .strSize = "8":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(63): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(64): .strType = "4":     .strName = "OrderNum":    .strSize = "4":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(65): .strType = "10":     .strName = "Reason":    .strSize = "10":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(66): .strType = "8":     .strName = "RequestDate":    .strSize = "8":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "Now()":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(67): .strType = "10":     .strName = "PrintedBy":    .strSize = "50":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(68): .strType = "10":     .strName = "Denom":    .strSize = "1":    .strSourceTable = gtblCashBook:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'CustAccounts
    With pstrTableAndFields(69): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(70): .strType = "10":     .strName = "Salutation":    .strSize = "15":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(71): .strType = "10":     .strName = "Surname":    .strSize = "25":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(72): .strType = "10":     .strName = "Initials":    .strSize = "20":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(73): .strType = "10":     .strName = "Add1":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(74): .strType = "10":     .strName = "Add2":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(75): .strType = "10":     .strName = "Add3":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(76): .strType = "10":     .strName = "Add4":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(77): .strType = "10":     .strName = "Add5":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(78): .strType = "10":     .strName = "Postcode":    .strSize = "9":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(79): .strType = "10":     .strName = "TelephoneNum":    .strSize = "25":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(80): .strType = "10":     .strName = "DeliveryAdd1":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(81): .strType = "10":     .strName = "DeliveryAdd2":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(82): .strType = "10":     .strName = "DeliveryAdd3":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(83): .strType = "10":     .strName = "DeliveryAdd4":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(84): .strType = "10":     .strName = "DeliveryAdd5":    .strSize = "30":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(85): .strType = "10":     .strName = "DeliveryPostcode":    .strSize = "9":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(86): .strType = "10":     .strName = "AccountType":    .strSize = "1":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(87): .strType = "10":     .strName = "ReceiveMailings":    .strSize = "1":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(88): .strType = "10":     .strName = "AcctInUseByFlag":    .strSize = "50":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(89): .strType = "4":     .strName = "BPCSCusNum":    .strSize = "4":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(90): .strType = "10":     .strName = "DBIndicator":    .strSize = "1":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = Chr(34) & "N" & Chr(34):     .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(91): .strType = "8":     .strName = "CreationDate":    .strSize = "8":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "Now()":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(92): .strType = "10":     .strName = "AcctStatus":    .strSize = "10":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(93): .strType = "10":     .strName = "DeliverySalutation":    .strSize = "15":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(94): .strType = "10":     .strName = "DeliverySurname":    .strSize = "25":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(95): .strType = "10":     .strName = "DeliveryInitials":    .strSize = "20":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(96): .strType = "10":     .strName = "EMail":    .strSize = "50":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(97): .strType = "10":     .strName = "EveTelephoneNum":    .strSize = "25":    .strSourceTable = gtblCustAccounts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'CustNotes
    With pstrTableAndFields(98): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblCustNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(99): .strType = "10":     .strName = "Notes":    .strSize = "255":    .strSourceTable = gtblCustNotes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'CustomReports
    With pstrTableAndFields(100): .strType = "10":     .strName = "CustRepName":    .strSize = "50":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(101): .strType = "12":     .strName = "ReportSQL":    .strSize = "0":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(102): .strType = "10":     .strName = "Type":    .strSize = "20":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(103): .strType = "10":     .strName = "SysDB":    .strSize = "10":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(104): .strType = "4":     .strName = "SequenceNum":    .strSize = "4":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(105): .strType = "1":     .strName = "InUse":    .strSize = "1":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(106): .strType = "10":     .strName = "Param":    .strSize = "10":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(107): .strType = "10":     .strName = "IntroVer":    .strSize = "50":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(108): .strType = "10":     .strName = "Settings":    .strSize = "20":    .strSourceTable = gtblCustomReports:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'ListDetailsMaster
    With pstrTableAndFields(109): .strType = "4":     .strName = "ListNum":    .strSize = "4":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(110): .strType = "4":     .strName = "SequenceNum":    .strSize = "4":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(111): .strType = "10":     .strName = "ListCode":    .strSize = "10":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(112): .strType = "10":     .strName = "Description":    .strSize = "50":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(113): .strType = "10":     .strName = "UserDef1":    .strSize = "50":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(114): .strType = "10":     .strName = "UserDef2":    .strSize = "50":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(115): .strType = "1":     .strName = "InUse":    .strSize = "1":    .strSourceTable = gtblMasterListDetails:    .strDataUpdatable = "True":    .strDefaultValue = "1":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'ListsMaster
    With pstrTableAndFields(116): .strType = "4":     .strName = "ListNum":    .strSize = "4":    .strSourceTable = gtblMasterLists:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(117): .strType = "10":     .strName = "ListName":    .strSize = "50":    .strSourceTable = gtblMasterLists:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(118): .strType = "1":     .strName = "SysUse":    .strSize = "1":    .strSourceTable = gtblMasterLists:    .strDataUpdatable = "True":    .strDefaultValue = "False":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(119): .strType = "10":     .strName = "Type":    .strSize = "1":    .strSourceTable = gtblMasterLists:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "True":    End With
 
    'OrderLinesMaster
    With pstrTableAndFields(133): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(134): .strType = "4":     .strName = "OrderNum":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(135): .strType = "10":     .strName = "CatNum":    .strSize = "10":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(136): .strType = "10":     .strName = "ItemDescription":    .strSize = "50":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(137): .strType = "10":     .strName = "BinLocation":    .strSize = "50":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(138): .strType = "4":     .strName = "Qty":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(139): .strType = "4":     .strName = "DespQty":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(140): .strType = "5":     .strName = "Price":    .strSize = "8":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(141): .strType = "5":     .strName = "Vat":    .strSize = "8":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(142): .strType = "4":     .strName = "Weight":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(143): .strType = "10":     .strName = "TaxCode":    .strSize = "1":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(144): .strType = "5":     .strName = "TotalPrice":    .strSize = "8":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(145): .strType = "10":     .strName = "TotalWeight":    .strSize = "50":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(146): .strType = "4":     .strName = "Class":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(147): .strType = "4":     .strName = "SalesCode":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(148): .strType = "4":     .strName = "OrderLineNum":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(149): .strType = "4":     .strName = "ParcelNumber":    .strSize = "4":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(150): .strType = "10":     .strName = "Denom":    .strSize = "1":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    If UCase$(App.ProductName) = "LITE" Then
    With pstrTableAndFields(151): .strType = "1":     .strName = "Recalced":    .strSize = "1":    .strSourceTable = gtblMasterOrderLines:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    End If
 
    'PADAvailableMaster
    With pstrTableAndFields(152): .strType = "10":     .strName = "Org_Unit_Code":    .strSize = "10":    .strSourceTable = gtblMasterPADAvailable:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'PADOfficeMaster
    With pstrTableAndFields(153): .strType = "10":     .strName = "Org_Unit_Code":    .strSize = "10":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(154): .strType = "10":     .strName = "FAD_Code":    .strSize = "7":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(155): .strType = "10":     .strName = "Add1":    .strSize = "40":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(156): .strType = "10":     .strName = "Add2":    .strSize = "40":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(157): .strType = "10":     .strName = "Add3":    .strSize = "40":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(158): .strType = "10":     .strName = "Add4":    .strSize = "40":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(159): .strType = "10":     .strName = "Add5":    .strSize = "40":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(160): .strType = "10":     .strName = "P_Code":    .strSize = "8":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(161): .strType = "10":     .strName = "P_Code_S":    .strSize = "4":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(162): .strType = "10":     .strName = "Contract":    .strSize = "7":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(163): .strType = "10":     .strName = "Name":    .strSize = "30":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(164): .strType = "10":     .strName = "Type":    .strSize = "4":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(165): .strType = "10":     .strName = "A_Date":    .strSize = "10":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(166): .strType = "10":     .strName = "Status":    .strSize = "1":    .strSourceTable = gtblMasterPADOffice:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'PADOpening_TimesMaster
    With pstrTableAndFields(167): .strType = "10":     .strName = "Org_Unit_Code":    .strSize = "10":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(168): .strType = "10":     .strName = "Time_Type":    .strSize = "9":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(169): .strType = "10":     .strName = "Weekday":    .strSize = "9":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(170): .strType = "10":     .strName = "From":    .strSize = "5":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(171): .strType = "10":     .strName = "To":    .strSize = "5":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(172): .strType = "10":     .strName = "Lunch_From":    .strSize = "5":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(173): .strType = "10":     .strName = "Lunch_To":    .strSize = "5":    .strSourceTable = gtblMasterPADOpeningTimes:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'PForce
    With pstrTableAndFields(174): .strType = "4":     .strName = "CustNum":    .strSize = "4":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(175): .strType = "4":     .strName = "OrderNum":    .strSize = "4":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(176): .strType = "10":     .strName = "Status":    .strSize = "1":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(177): .strType = "10":     .strName = "ConsignNum":    .strSize = "9":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "True":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(178): .strType = "10":     .strName = "ServiceID":    .strSize = "10":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(179): .strType = "10":     .strName = "BatchNumber":    .strSize = "4":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(180): .strType = "8":     .strName = "DespatchDate":    .strSize = "8":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(181): .strType = "10":     .strName = "DeliverySalutation":    .strSize = "15":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(182): .strType = "10":     .strName = "DeliverySurname":    .strSize = "25":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(183): .strType = "10":     .strName = "DeliveryInitials":    .strSize = "20":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(184): .strType = "10":     .strName = "DeliveryAdd1":    .strSize = "30":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(185): .strType = "10":     .strName = "DeliveryAdd2":    .strSize = "30":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(186): .strType = "10":     .strName = "DeliveryAdd3":    .strSize = "30":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(187): .strType = "10":     .strName = "DeliveryAdd4":    .strSize = "30":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(188): .strType = "10":     .strName = "DeliveryAdd5":    .strSize = "30":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(189): .strType = "10":     .strName = "DeliveryPostcode":    .strSize = "9":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(190): .strType = "3":     .strName = "ParcelItems":    .strSize = "2":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(191): .strType = "4":     .strName = "GrossWeight":    .strSize = "4":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(192): .strType = "10":     .strName = "WeekendHandCode":    .strSize = "10":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(193): .strType = "10":     .strName = "PrepaidInd":    .strSize = "1":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(194): .strType = "10":     .strName = "NotificationCode":    .strSize = "10":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(195): .strType = "10":     .strName = "ConsignRemark":    .strSize = "100":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(196): .strType = "10":     .strName = "SpecialSatDel":    .strSize = "1":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(197): .strType = "10":     .strName = "SpecialBookIn":    .strSize = "1":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(198): .strType = "10":     .strName = "SpecialProof":    .strSize = "1":    .strSourceTable = gtblPForce:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'ProductsMaster
    With pstrTableAndFields(199): .strType = "10":     .strName = "CatNum":    .strSize = "10":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(200): .strType = "10":     .strName = "ItemDescription":    .strSize = "50":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(201): .strType = "10":     .strName = "BinLocation":    .strSize = "50":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(202): .strType = "4":     .strName = "Class":    .strSize = "4":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(203): .strType = "10":     .strName = "ClassItem":    .strSize = "50":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(204): .strType = "10":     .strName = "ClassGroup":    .strSize = "50":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(205): .strType = "5":     .strName = "Price":    .strSize = "8":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(206): .strType = "4":     .strName = "Weight":    .strSize = "4":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(207): .strType = "10":     .strName = "TaxCode":    .strSize = "1":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(208): .strType = "4":     .strName = "NumInStock":    .strSize = "4":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(209): .strType = "4":     .strName = "Qty":    .strSize = "4":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(210): .strType = "10":     .strName = "Comments":    .strSize = "255":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(211): .strType = "8":     .strName = "UserUpdated":    .strSize = "8":    .strSourceTable = gtblMasterProducts:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'Remarks
    With pstrTableAndFields(212): .strType = "4":     .strName = "RemarkNum":    .strSize = "4":    .strSourceTable = gtblRemarks:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(213): .strType = "12":     .strName = "Remark":    .strSize = "0":    .strSourceTable = gtblRemarks:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(214): .strType = "10":     .strName = "Type":    .strSize = "10":    .strSourceTable = gtblRemarks:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(215): .strType = "10":     .strName = "LockingFlag":    .strSize = "50":    .strSourceTable = gtblRemarks:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(216): .strType = "8":     .strName = "CreationDate":    .strSize = "8":    .strSourceTable = gtblRemarks:    .strDataUpdatable = "True":    .strDefaultValue = "Now()":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'Substitutions
    With pstrTableAndFields(225): .strType = "10":     .strName = "SubCatNum":    .strSize = "10":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(226): .strType = "10":     .strName = "CatNum":    .strSize = "10":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(227): .strType = "4":     .strName = "SubItemQty":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(228): .strType = "10":     .strName = "Reason":    .strSize = "50":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(229): .strType = "10":     .strName = "TranType":    .strSize = "10":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(230): .strType = "10":     .strName = "ItemDescription":    .strSize = "50":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(231): .strType = "10":     .strName = "BinLocation":    .strSize = "50":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(232): .strType = "4":     .strName = "Class":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(233): .strType = "10":     .strName = "ClassItem":    .strSize = "50":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(234): .strType = "10":     .strName = "ClassGroup":    .strSize = "50":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(235): .strType = "5":     .strName = "Price":    .strSize = "8":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(236): .strType = "4":     .strName = "Weight":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(237): .strType = "10":     .strName = "TaxCode":    .strSize = "1":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(238): .strType = "4":     .strName = "NumInStock":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(239): .strType = "4":     .strName = "Qty":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "0":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(240): .strType = "10":     .strName = "Comments":    .strSize = "255":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(241): .strType = "8":     .strName = "UserUpdated":    .strSize = "8":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(242): .strType = "4":     .strName = "OrderLineNum":    .strSize = "4":    .strSourceTable = gtblSubstitutions:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'System
    With pstrTableAndFields(243): .strType = "10":     .strName = "Item":    .strSize = "20":    .strSourceTable = gtblSystem:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(244): .strType = "10":     .strName = "Value":    .strSize = "255":    .strSourceTable = gtblSystem:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(245): .strType = "8":     .strName = "DateCreated":    .strSize = "8":    .strSourceTable = gtblSystem:    .strDataUpdatable = "True":    .strDefaultValue = "Now()":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(246): .strType = "8":     .strName = "OtherDate":    .strSize = "8":    .strSourceTable = gtblSystem:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
 
    'Users
    With pstrTableAndFields(247): .strType = "10":     .strName = "UserID":    .strSize = "20":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(248): .strType = "10":     .strName = "UserPassword":    .strSize = "255":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "True":    End With
    With pstrTableAndFields(249): .strType = "10":     .strName = "UserName":    .strSize = "30":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "True":    End With
    With pstrTableAndFields(250): .strType = "4":     .strName = "UserLevel":    .strSize = "4":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(251): .strType = "12":     .strName = "UserNotes":    .strSize = "0":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
    With pstrTableAndFields(252): .strType = "10":     .strName = "Phase":    .strSize = "55":    .strSourceTable = gtblUsers:    .strDataUpdatable = "True":    .strDefaultValue = "":    .strRequired = "False":    .strAllowZeroLength = "False":    End With
End Sub
Sub PopFeatList(pobjList As Object, pAll As Boolean, plngIndex() As Long)
Dim lCtr As Long
Dim lProd As String

    lProd = UCase$(App.ProductName)
    If Left$(App.ProductName, 1) <> "M" Then lProd = "M" & lProd
    lProd = Left$(lProd, 10)
    ReDim plngIndex(70)

    With pobjList
        pobjList.Clear
        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "07/06/02 - NEW: Final PAD changes                  ": _
            plngIndex(lCtr) = 236: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "06/06/02 - NEW: Improved import process            ": _
            plngIndex(lCtr) = 235: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "05/06/02 - NEW: PO Collect, import, select screen  ": _
            plngIndex(lCtr) = 234: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "31/05/02 - NEW: Consign note on thermal, buffer    ": _
            plngIndex(lCtr) = 232: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "30/05/02 - FIX: 'Update Total' routine to grid chng": _
            plngIndex(lCtr) = 230: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "27/05/02 - FIX: Error trap in for old layout files ": _
            plngIndex(lCtr) = 227: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "03/05/02 - NEW: Added status to order screen       ": _
            plngIndex(lCtr) = 218: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "17/04/02 - FIX: VAT, some hard coded occurences    ": _
            plngIndex(lCtr) = 214: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "29/03/02 - NEW: Added Order status, Packing screen ": _
            plngIndex(lCtr) = 213: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "16/03/02 - FIX: Back btns Order Dets & Order screen": _
            plngIndex(lCtr) = 203: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "16/03/02 - NEW: <CR> to next field                 ": _
            plngIndex(lCtr) = 202: lCtr = lCtr + 1

        If lProd = "LOADER" Or pAll Then: _
            .AddItem "16/03/02 - NEW: Change company name in loader.exe  ": _
            plngIndex(lCtr) = 92: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "13/03/02 - NEW: Lmiting Lite, e.g. 50 customers    ": _
            plngIndex(lCtr) = 174: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "13/03/02 - NEW: Default Lists values, machine & Sys": _
            plngIndex(lCtr) = 188: lCtr = lCtr + 1

        If lProd = "MMAINTENAN" Or pAll Then: _
            .AddItem "10/03/02 - NEW: active users list on upgrade screen": _
            plngIndex(lCtr) = 75: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "10/03/02 - NEW: Clear remark flag function         ": _
            plngIndex(lCtr) = 87: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "08/03/02 - FIX: Custom reporting, missing last page": _
            plngIndex(lCtr) = 189: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "08/03/02 - FIX: Put more work into Advice Note     ": _
            plngIndex(lCtr) = 107: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "07/03/02 - NEW: Add hide report date feature       ": _
            plngIndex(lCtr) = 182: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "07/03/02 - NEW: Itelligent manadotory error goto   ": _
            plngIndex(lCtr) = 94: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "06/03/02 - FIX: Various, consign, canel security   ": _
            plngIndex(lCtr) = 185: lCtr = lCtr + 1

        If lProd = "MCLIAPP" Or pAll Then: _
            .AddItem "05/03/02 - NEW: Abort option dialogue              ": _
            plngIndex(lCtr) = 105: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "04/03/02 - FIX: File History Selection Trap        ": _
            plngIndex(lCtr) = 184: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "03/03/02 - NEW: Route out Donation for Standard    ": _
            plngIndex(lCtr) = 106: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "02/03/02 - NEW: A feature to re-calc the Products  ": _
            plngIndex(lCtr) = 119: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "01/03/02 - NEW: Finish Request screen              ": _
            plngIndex(lCtr) = 116: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "26/02/02 - NEW: Table constant names               ": _
            plngIndex(lCtr) = 131: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "23/02/02 - NEW: Need cut down product maint screen ": _
            plngIndex(lCtr) = 115: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "22/02/02 - NEW: Will need an initial data screen.  ": _
            plngIndex(lCtr) = 121: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "20/02/02 - NEW: Added Minder Full to Menu          ": _
            plngIndex(lCtr) = 114: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "19/02/02 - NEW: Finish Slim program                ": _
            plngIndex(lCtr) = 76: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "19/02/02 - FIX: Nag screen in Lite                 ": _
            plngIndex(lCtr) = 70: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "19/02/02 - NEW: Draw icon for Values screen        ": _
            plngIndex(lCtr) = 112: lCtr = lCtr + 1

            .AddItem "18/02/02 - FIX: Refresh scroll bars after Zoom": _
            plngIndex(lCtr) = 95: lCtr = lCtr + 1

            .AddItem "18/02/02 - NEW: New Advice Layout W/Comp Address   ": _
            plngIndex(lCtr) = 78: lCtr = lCtr + 1

            .AddItem "18/02/02 - FIX: Copies on page setup": _
            plngIndex(lCtr) = 96: lCtr = lCtr + 1

            .AddItem "18/02/02 - NEW: Remove client specific reporting   ": _
            plngIndex(lCtr) = 88: lCtr = lCtr + 1

            .AddItem "15/02/02 - NEW: Menus                              ": _
            plngIndex(lCtr) = 100: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "14/02/02 - NEW: Route Savesetting MMOS             ": _
            plngIndex(lCtr) = 103: lCtr = lCtr + 1

        If lProd = "MLITE" Or pAll Then: _
            .AddItem "11/02/02 - NEW: Values Screen LITE version         ": _
            plngIndex(lCtr) = 71: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "30/01/02 - FIX: Acct in use by flag                ": _
            plngIndex(lCtr) = 40: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "30/01/02 - NEW: Added clear acct in use function": _
            plngIndex(lCtr) = 67: lCtr = lCtr + 1

            .AddItem "29/01/02 - NEW: Standard menu All programs": _
            plngIndex(lCtr) = 46: lCtr = lCtr + 1

            .AddItem "29/01/02 - NEW: New features list": _
            plngIndex(lCtr) = 66: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "28/01/02 - FIX: Scrollbar Order Maintenance screen": _
            plngIndex(lCtr) = 44: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "28/01/02 - NEW: Standardised menu options": _
            plngIndex(lCtr) = 45: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "24/01/02 - FIX: Duplicate grid showing Raidzone    ": _
            plngIndex(lCtr) = 39: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "23/01/02 - FIX: Refund button                      ": _
            plngIndex(lCtr) = 38: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "23/01/02 - FIX: Quality Advice Amount": _
            plngIndex(lCtr) = 49: lCtr = lCtr + 1

            .AddItem "23/01/02 - NEW: Report Layout files": _
            plngIndex(lCtr) = 55: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "22/01/02 - NEW: Duplicate Handling": _
            plngIndex(lCtr) = 37: lCtr = lCtr + 1

            .AddItem "21/01/02 - NEW: Extra space product details": _
            plngIndex(lCtr) = 48: lCtr = lCtr + 1

            .AddItem "17/01/02 - NEW: Report layout preparation": _
            plngIndex(lCtr) = 53: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "14/01/02 - FIX: Flexible Card date selection": _
            plngIndex(lCtr) = 25: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "09/01/02 - NEW: Custom report orientation": _
            plngIndex(lCtr) = 56: lCtr = lCtr + 1

            .AddItem "09/01/02 - NEW: Temporary unlock codes": _
            plngIndex(lCtr) = 54: lCtr = lCtr + 1

            .AddItem "08/01/02 - NEW: Hide zoom on print preview": _
            plngIndex(lCtr) = 52: lCtr = lCtr + 1

        If lProd = "MMAINTENAN" Or pAll Then: _
            .AddItem "07/01/02 - NEW: Marketing settings moved": _
            plngIndex(lCtr) = 65: lCtr = lCtr + 1

            .AddItem "07/01/02 - FIX: Unlock code freedom": _
            plngIndex(lCtr) = 51: lCtr = lCtr + 1

            .AddItem "06/01/02 - NEW: New Program provision": _
            plngIndex(lCtr) = 50: lCtr = lCtr + 1

            .AddItem "02/01/02 - NEW: More system colour support": _
            plngIndex(lCtr) = 47: lCtr = lCtr + 1

            .AddItem "30/12/01 - FIX: Static.ini dependency removed": _
            plngIndex(lCtr) = 57: lCtr = lCtr + 1

            .AddItem "30/12/01 - FIX: GPF when closing program": _
            plngIndex(lCtr) = 62: lCtr = lCtr + 1

            .AddItem "27/12/01 - NEW: Advanced Locking": _
            plngIndex(lCtr) = 61: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "18/12/01 - NEW: Extra options on Customer select": _
            plngIndex(lCtr) = 22: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "18/12/01 - NEW: Extra options on Customer select": _
            plngIndex(lCtr) = 63: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "13/12/01 - FIX: Quality reporting problems": _
            plngIndex(lCtr) = 58: lCtr = lCtr + 1

        If lProd = "MMANAGER" Or pAll Then: _
            .AddItem "12/12/01 - NEW: Label printing work started": _
            plngIndex(lCtr) = 59: lCtr = lCtr + 1

            .AddItem "06/12/01 - FIX: Currency by PC": _
            plngIndex(lCtr) = 60: lCtr = lCtr + 1

        If lProd = "MCLIENT" Or pAll Then: _
            .AddItem "06/12/01 - FIX: Grid currency and safety": _
            plngIndex(lCtr) = 64: lCtr = lCtr + 1

        If lProd = "MMOSALL" Or pAll Then: _
            .AddItem "22/10/01 - NEW: Extended Advice memo to 100 chars  ": _
            plngIndex(lCtr) = 11: lCtr = lCtr + 1

    End With

End Sub
Sub FeatureMsg(plngItem As Long)
Dim lstrMsg As String

    Select Case plngItem
    Case 236: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "07/06/02" & vbCrLf & vbCrLf & _
            "NEW: Final PAD changes                  " & vbCrLf & vbCrLf & _
            "Made changes as previously queried with the client, post " & vbCrLf & _
            "office had advised revised process.  This suggested that " & vbCrLf & _
            "the data may not always be as provided at present, put in " & vbCrLf & _
            "extra where clause in selection process to cater for none " & vbCrLf & _
            "reference code records.                                     "
    Case 235: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "06/06/02" & vbCrLf & vbCrLf & _
            "NEW: Improved import process            " & vbCrLf & vbCrLf & _
            "The import process derived from the post office sample " & vbCrLf & _
            "database was extremely slow and generated 200mb of data.  " & vbCrLf & _
            "Analysed the database and found some queries that didn't do " & vbCrLf & _
            "anything.  Client also decided to miss out the 3rd line of " & vbCrLf & _
            "the address to make room the post office name and P/O.      "
    Case 234: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "05/06/02" & vbCrLf & vbCrLf & _
            "NEW: PO Collect, import, select screen  " & vbCrLf & vbCrLf & _
            "Added new Post Office Parcel collection data into the " & vbCrLf & _
            "system.  This involved an import process, adding extra " & vbCrLf & _
            "local and master tables and a selection screen.             "
    Case 232: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "31/05/02" & vbCrLf & vbCrLf & _
            "NEW: Consign note on thermal, buffer    " & vbCrLf & vbCrLf & _
            "Added a new feature to add the consignment note to the " & vbCrLf & _
            "thermal label.  This involved increasing the size of the " & vbCrLf & _
            "buffer and the append query.  Also, a referential integrity " & vbCrLf & _
            "problem arose so a lockup feature was added to stop updates " & vbCrLf & _
            "being made to the consignment note once the order had been " & vbCrLf & _
            "packed.                                                     "
    Case 230: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "30/05/02" & vbCrLf & vbCrLf & _
            "FIX: 'Update Total' routine to grid chng" & vbCrLf & vbCrLf & _
            "With the requested removal of the 'Update Total' buttons, " & vbCrLf & _
            "users are now have to click on a text box after they have " & vbCrLf & _
            "modified a product on the order screen.  A change has been " & vbCrLf & _
            "made, so that when a product is changed this invokes the " & vbCrLf & _
            "new 'Update Total' routine.                                 "
    Case 227: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "27/05/02" & vbCrLf & vbCrLf & _
            "FIX: Error trap in for old layout files " & vbCrLf & vbCrLf & _
            "This displays a basic message, but could do with more work, " & vbCrLf & _
            "as in higher function this causes another error and crashes " & vbCrLf & _
            "the program.  However, this would only occur is someone was " & vbCrLf & _
            "hacking around or hand't got the latest files, but had the " & vbCrLf & _
            "latest program.                                             "
    Case 218: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "03/05/02" & vbCrLf & vbCrLf & _
            "NEW: Added status to order screen       " & vbCrLf & vbCrLf & _
            "Added status change btn to order screen when modify         "
    Case 214: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "17/04/02" & vbCrLf & vbCrLf & _
            "FIX: VAT, some hard coded occurences    " & vbCrLf & vbCrLf & _
            "With the announcement that VAT MAY change on Budget day, " & vbCrLf & _
            "decided to check for hard coded 17.5.  Having found " & vbCrLf & _
            "occurrences in UpdateStock, OrderMasterDespTotal and " & vbCrLf & _
            "CommisionAnalysis report, have made the necessary changes.  " & vbCrLf & _
            "This will be available in the next version.                 "
    Case 213: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "29/03/02" & vbCrLf & vbCrLf & _
            "NEW: Added Order status, Packing screen " & vbCrLf & vbCrLf & _
            "Noticed that order status wasn't on the packing screen.  " & vbCrLf & _
            "This may be useful in the future.                           "
    Case 203: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "16/03/02" & vbCrLf & vbCrLf & _
            "FIX: Back btns Order Dets & Order screen" & vbCrLf & vbCrLf & _
            "The back buttons, weren't passing the route information and " & vbCrLf & _
            "the current loaded form values.  This was only a problem " & vbCrLf & _
            "when returning to the account screen in another order, " & vbCrLf & _
            "after using the back button(s).                             "
    Case 202: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "16/03/02" & vbCrLf & vbCrLf & _
            "NEW: <CR> to next field                 " & vbCrLf & vbCrLf & _
            "Added new feature, which enables the use of carriage " & vbCrLf & _
            "return, to move to the next field.  This applies to the " & vbCrLf & _
            "Account, Order Details and Order screens.  Also, made " & vbCrLf & _
            "credit card details hidden if not selecting payment type " & vbCrLf & _
            "credit card.                                                "
    Case 92: lstrMsg = "PROGRAM: LOADER" & String(3, vbTab) & "16/03/02" & vbCrLf & vbCrLf & _
            "NEW: Change company name in loader.exe  " & vbCrLf & vbCrLf & _
            "Changed company name in loader.exe                          "
    Case 174: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "13/03/02" & vbCrLf & vbCrLf & _
            "NEW: Lmiting Lite, e.g. 50 customers    " & vbCrLf & vbCrLf & _
            "Restrictions have now been put in place to limite the " & vbCrLf & _
            "amount of customers which can be added to the Lite version " & vbCrLf & _
            "to 50.                                                      "
    Case 188: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "13/03/02" & vbCrLf & vbCrLf & _
            "NEW: Default Lists values, machine & Sys" & vbCrLf & vbCrLf & _
            "Added function which will insert default records which if " & vbCrLf & _
            "not present will cause update queries to fail.              "
    Case 75: lstrMsg = "PROGRAM: MMAINTENAN" & String(3, vbTab) & "10/03/02" & vbCrLf & vbCrLf & _
            "NEW: active users list on upgrade screen" & vbCrLf & vbCrLf & _
            "Finished active users list on upgrade screen.               "
    Case 87: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "10/03/02" & vbCrLf & vbCrLf & _
            "NEW: Clear remark flag function         " & vbCrLf & vbCrLf & _
            "Although not necessary, a feature that clears the remark " & vbCrLf & _
            "locking flag has been added.  Similar functions are used " & vbCrLf & _
            "for the advice note and customer account records.  However, " & vbCrLf & _
            "with a recent error which caused the locking flag not to be " & vbCrLf & _
            "unique, the clear function would have prevented it.         "
    Case 189: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "08/03/02" & vbCrLf & vbCrLf & _
            "FIX: Custom reporting, missing last page" & vbCrLf & vbCrLf & _
            "A problem emerged which meant that the last page of a " & vbCrLf & _
            "custom report wasn't being printed.  This was due to a " & vbCrLf & _
            "precision point calculation on the number of pages in a " & vbCrLf & _
            "report.  This would not have been noticeable on one-page " & vbCrLf & _
            "printouts.                                                  "
    Case 107: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "08/03/02" & vbCrLf & vbCrLf & _
            "FIX: Put more work into Advice Note     " & vbCrLf & vbCrLf & _
            "A fix from another problem had stopped the alignment " & vbCrLf & _
            "function from aligning detail lines.  The reason for the " & vbCrLf & _
            "initial fix which 'turned off' the alignment in the first " & vbCrLf & _
            "place is currently unknown.                                 "
    Case 182: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "07/03/02" & vbCrLf & vbCrLf & _
            "NEW: Add hide report date feature       " & vbCrLf & vbCrLf & _
            "Dates are now only shown on custom reports which make use " & vbCrLf & _
            "of them.                                                    "
    Case 94: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "07/03/02" & vbCrLf & vbCrLf & _
            "NEW: Itelligent manadotory error goto   " & vbCrLf & vbCrLf & _
            "When an error is produced due to a user not satisfying a " & vbCrLf & _
            "mandatory field requirement the system now moves focus to " & vbCrLf & _
            "the field which has the problem.  This affects the Account, " & vbCrLf & _
            "Order Detail and Order screen.                              "
    Case 185: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "06/03/02" & vbCrLf & vbCrLf & _
            "FIX: Various, consign, canel security   " & vbCrLf & vbCrLf & _
            "Various small fixes and tightening features were added. " & vbCrLf & _
            "This included:-  Consignment details can now been accessed " & vbCrLf & _
            "in a summary format from the order history screen. Logical " & vbCrLf & _
            "fixes for the Valid From date and telephone number for " & vbCrLf & _
            "ex-directory.  A fix was made for the specific advice " & vbCrLf & _
            "print, which cleared the customer number and made features " & vbCrLf & _
            "like the child cash book not work correctly.  The update "
    Case 105: lstrMsg = "PROGRAM: MCLIAPP" & String(3, vbTab) & "05/03/02" & vbCrLf & vbCrLf & _
            "NEW: Abort option dialogue              " & vbCrLf & vbCrLf & _
            "A new feature which will allow users to save an order " & vbCrLf & _
            "without having to loose what they have entered so far.  The " & vbCrLf & _
            "order status is set to cancelled and must be reset.         "
    Case 184: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "04/03/02" & vbCrLf & vbCrLf & _
            "FIX: File History Selection Trap        " & vbCrLf & vbCrLf & _
            "When selecting a file history option, if the order or " & vbCrLf & _
            "account did not exist this would cause an error.  The error " & vbCrLf & _
            "would normal not appear anyway, except when test history " & vbCrLf & _
            "items were added to the list and use in the live version or " & vbCrLf & _
            "vice versa                                                  "
    Case 106: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "03/03/02" & vbCrLf & vbCrLf & _
            "NEW: Route out Donation for Standard    " & vbCrLf & vbCrLf & _
            "Have now added feature that hides donation from the client " & vbCrLf & _
            "program and the advice note for the standard route version. "
    Case 119: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "02/03/02" & vbCrLf & vbCrLf & _
            "NEW: A feature to re-calc the Products  " & vbCrLf & vbCrLf & _
            "A new relatively simple feature to implement was added to " & vbCrLf & _
            "the Lite version to re-allocate products assigned but not " & vbCrLf & _
            "despatched was added.  This feature could be used in the " & vbCrLf & _
            "full version at a later date.                               "
    Case 116: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "01/03/02" & vbCrLf & vbCrLf & _
            "NEW: Finish Request screen              " & vbCrLf & vbCrLf & _
            "A request screen allowing users of the Lite version to " & vbCrLf & _
            "order the full program and print an order form was " & vbCrLf & _
            "finished.                                                   "
    Case 131: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "26/02/02" & vbCrLf & vbCrLf & _
            "NEW: Table constant names               " & vbCrLf & vbCrLf & _
            "A feature was added to provide database security in the " & vbCrLf & _
            "Lite version database.  This involved substituting all " & vbCrLf & _
            "tables names used in queries and database functions with " & vbCrLf & _
            "constant variables.  This only affected code used in the " & vbCrLf & _
            "Lite version, but most of the code is used in other " & vbCrLf & _
            "programs.  Testing showed no problems.                      "
    Case 115: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "23/02/02" & vbCrLf & vbCrLf & _
            "NEW: Need cut down product maint screen " & vbCrLf & vbCrLf & _
            "Basic stock view screen has been added to the Lite version. " & vbCrLf & _
            " This allows basic grid data access, add, edit and delete.  "
    Case 121: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "22/02/02" & vbCrLf & vbCrLf & _
            "NEW: Will need an initial data screen.  " & vbCrLf & vbCrLf & _
            "Certain data like contact name and phone number are not " & vbCrLf & _
            "required.  Some protection against denomination different " & vbCrLf & _
            "to regional settings has also been setup.                   "
    Case 114: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "20/02/02" & vbCrLf & vbCrLf & _
            "NEW: Added Minder Full to Menu          " & vbCrLf & vbCrLf & _
            "Added Minder Full to Menu, this was possible to a new Force " & vbCrLf & _
            "App shutdown feature developed for the Lite version.  " & vbCrLf & _
            "Minder Full was previously removed from the system, FYI for " & vbCrLf & _
            "new clients, this feature deletes temporary files and runs " & vbCrLf & _
            "scandisk and defrag.                                        "
    Case 76: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "19/02/02" & vbCrLf & vbCrLf & _
            "NEW: Finish Slim program                " & vbCrLf & vbCrLf & _
            "Finished basic (with no anti crack) Check For Updates " & vbCrLf & _
            "screen.  This included writing the Slim.exe program which " & vbCrLf & _
            "will have routing for specific versions in the future.      "
    Case 70: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "19/02/02" & vbCrLf & vbCrLf & _
            "FIX: Nag screen in Lite                 " & vbCrLf & vbCrLf & _
            "Fixed Nag screen in Lite, tab click problem.                "
    Case 112: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "19/02/02" & vbCrLf & vbCrLf & _
            "NEW: Draw icon for Values screen        " & vbCrLf & vbCrLf & _
            "Updated Icon for Reference data and also used this icon in " & vbCrLf & _
            "the Values screen of the Lite version.                      "
    Case 95: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "18/02/02" & vbCrLf & vbCrLf & _
            "FIX: Refresh scroll bars after Zoom" & vbCrLf & vbCrLf & _
            "When changing the zoom setting in the print preview screen " & vbCrLf & _
            "the scroll bars were not being reset.  This has now been " & vbCrLf & _
            "fixed.                                                      "
    Case 78: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "18/02/02" & vbCrLf & vbCrLf & _
            "NEW: New Advice Layout W/Comp Address   " & vbCrLf & vbCrLf & _
            "A new advice note layout has been developed with the " & vbCrLf & _
            "company address on it.  This will require clients to be " & vbCrLf & _
            "issued with new program version and new report layouts.     "
    Case 96: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "18/02/02" & vbCrLf & vbCrLf & _
            "FIX: Copies on page setup" & vbCrLf & vbCrLf & _
            "The copies feature on the print preview screen now works!   "
    Case 88: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "18/02/02" & vbCrLf & vbCrLf & _
            "NEW: Remove client specific reporting   " & vbCrLf & vbCrLf & _
            "Reports that are client specific have been removed from the " & vbCrLf & _
            "standard route version of the program.                      "
    Case 100: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "15/02/02" & vbCrLf & vbCrLf & _
            "NEW: Menus                              " & vbCrLf & vbCrLf & _
            "Added menus for all programs. The most impressive feature " & vbCrLf & _
            "is the file history options, which will undoubtedly save " & vbCrLf & _
            "users a lot of time.                                        "
    Case 103: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "14/02/02" & vbCrLf & vbCrLf & _
            "NEW: Route Savesetting MMOS             " & vbCrLf & vbCrLf & _
            "NEW: Setup all savesetting uses with a variable that is set " & vbCrLf & _
            "at the start of the program.  This has been used to ensure " & vbCrLf & _
            "all standard routes use 'Mindwarp Mail Order System' also " & vbCrLf & _
            "client specific settings and for use in other applications. "
    Case 71: lstrMsg = "PROGRAM: MLITE" & String(3, vbTab) & "11/02/02" & vbCrLf & vbCrLf & _
            "NEW: Values Screen LITE version         " & vbCrLf & vbCrLf & _
            "Added new screen similar to allow access to drop down list " & vbCrLf & _
            "values.  This is a compromise to not having an admin or " & vbCrLf & _
            "config program.                                             "
    Case 40: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "30/01/02" & vbCrLf & vbCrLf & _
            "FIX: Acct in use by flag                " & vbCrLf & vbCrLf & _
            "A problem occurred in relation to the account " & vbCrLf & _
            "in-use-by-flag that occurred when the same user entered " & vbCrLf & _
            "more than one order for the same customer with the program " & vbCrLf & _
            "not being closed down between entries.  This has now been " & vbCrLf & _
            "fixed.                                                      "
    Case 67: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "30/01/02" & vbCrLf & vbCrLf & _
            "NEW: Added clear acct in use function" & vbCrLf & vbCrLf & _
            "To provide extra reassurance a feature to clear the order " & vbCrLf & _
            "account in-use-by-flag was added.  Although this should not " & vbCrLf & _
            "be necessary it does cater for the possibility of corrupt " & vbCrLf & _
            "indexes or a fault in setting the flag.                     "
    Case 46: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "29/01/02" & vbCrLf & vbCrLf & _
            "NEW: Standard menu All programs" & vbCrLf & vbCrLf & _
            "Added the new standard menu feature to all other programs.  " & vbCrLf & _
            "This involved adding a generic function to the Form_Paint " & vbCrLf & _
            "event to allow programmatic form refreshes.  This " & vbCrLf & _
            "functionality could also be used for other purposes in the " & vbCrLf & _
            "future                                                      "
    Case 66: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "29/01/02" & vbCrLf & vbCrLf & _
            "NEW: New features list" & vbCrLf & vbCrLf & _
            "This List. Went back through January and December, and " & vbCrLf & _
            "added some details about new features and fixes.  The code " & vbCrLf & _
            "for this feature is automatically generated from an offline " & vbCrLf & _
            "support database and therefore should not cause any " & vbCrLf & _
            "problems itself.                                            "
    Case 44: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "28/01/02" & vbCrLf & vbCrLf & _
            "FIX: Scrollbar Order Maintenance screen" & vbCrLf & vbCrLf & _
            "The grid on the Order Maintenance screen did not show the " & vbCrLf & _
            "scroll bar when first entering the screen. This has now " & vbCrLf & _
            "been fixed and should show both a vertical or horizontal " & vbCrLf & _
            "scroll bar depending on data.                               "
    Case 45: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "28/01/02" & vbCrLf & vbCrLf & _
            "NEW: Standardised menu options" & vbCrLf & vbCrLf & _
            "Developed new functions to show standard menus.  This will " & vbCrLf & _
            "make it easier to create new menu options and allow screen " & vbCrLf & _
            "specific options.                                           "
    Case 39: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "24/01/02" & vbCrLf & vbCrLf & _
            "FIX: Duplicate grid showing Raidzone    " & vbCrLf & vbCrLf & _
            "On the Duplicate grid an error occurred while in testing, " & vbCrLf & _
            "showing the development server name.  This small fix to " & vbCrLf & _
            "implement.                                                  "
    Case 38: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "23/01/02" & vbCrLf & vbCrLf & _
            "FIX: Refund button                      " & vbCrLf & vbCrLf & _
            "The Refund button was not functioning.  This was caused by " & vbCrLf & _
            "the new safer grid contents on the Order Maintenance " & vbCrLf & _
            "screen.                                                     "
    Case 49: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "23/01/02" & vbCrLf & vbCrLf & _
            "FIX: Quality Advice Amount" & vbCrLf & vbCrLf & _
            "Discovered a potential problem, which would affect quality " & vbCrLf & _
            "advice notes.  The default value of 20 items to print was " & vbCrLf & _
            "over riding any other amount specified.                     "
    Case 55: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "23/01/02" & vbCrLf & vbCrLf & _
            "NEW: Report Layout files" & vbCrLf & vbCrLf & _
            "Work was completed on report layout files.  This included " & vbCrLf & _
            "features to copy and upgrade files when made available.     "
    Case 37: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "22/01/02" & vbCrLf & vbCrLf & _
            "NEW: Duplicate Handling" & vbCrLf & vbCrLf & _
            "Developed new screen to allow duplicates customer accounts " & vbCrLf & _
            "to be merged, so that all of a customers orders would " & vbCrLf & _
            "appear under one account record.  Also made use of the cell " & vbCrLf & _
            "wrap around and a default height of cell for data grid.     "
    Case 48: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "21/01/02" & vbCrLf & vbCrLf & _
            "NEW: Extra space product details" & vbCrLf & vbCrLf & _
            "With the forthcoming report layout files, extra space will " & vbCrLf & _
            "be required to allow the default order detail framework to " & vbCrLf & _
            "be used for other layouts.  This will includes Invoice " & vbCrLf & _
            "layouts. To be more specific certain fields like bin " & vbCrLf & _
            "location were expanded behind the scenes to all them to be " & vbCrLf & _
            "used for other purposes.                                    "
    Case 53: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "17/01/02" & vbCrLf & vbCrLf & _
            "NEW: Report layout preparation" & vbCrLf & vbCrLf & _
            "The report layout structure was modified to include file " & vbCrLf & _
            "support.  Although not complete, provision for extra detail " & vbCrLf & _
            "like page size, layout name etc was added.                  "
    Case 25: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "14/01/02" & vbCrLf & vbCrLf & _
            "FIX: Flexible Card date selection" & vbCrLf & vbCrLf & _
            "More flexibility with card dates entry was provided. This " & vbCrLf & _
            "may have caused a problem in years to come when modifying " & vbCrLf & _
            "this information.  Whether this would ever arise may be " & vbCrLf & _
            "questionable.                                               "
    Case 56: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "09/01/02" & vbCrLf & vbCrLf & _
            "NEW: Custom report orientation" & vbCrLf & vbCrLf & _
            "Added a new feature to allow pre-defined paper orientation " & vbCrLf & _
            "to be used on every custom report.  This will be especially " & vbCrLf & _
            "valuable to newer users how may not realise that a specific " & vbCrLf & _
            "report should be landscape.                                 "
    Case 54: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "09/01/02" & vbCrLf & vbCrLf & _
            "NEW: Temporary unlock codes" & vbCrLf & vbCrLf & _
            "An extra feature to allow clients to be issued with " & vbCrLf & _
            "temporary unlock codes was added.  This will bring up a " & vbCrLf & _
            "reminder when the date approaches.  Then halt the program " & vbCrLf & _
            "when the date has past.                                     "
    Case 52: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "08/01/02" & vbCrLf & vbCrLf & _
            "NEW: Hide zoom on print preview" & vbCrLf & vbCrLf & _
            "A new feature was added to enable the zoom feature to be " & vbCrLf & _
            "hidden from the user.  This was introduced through " & vbCrLf & _
            "scalability problems with label layouts.                    "
    Case 65: lstrMsg = "PROGRAM: MMAINTENAN" & String(3, vbTab) & "07/01/02" & vbCrLf & vbCrLf & _
            "NEW: Marketing settings moved" & vbCrLf & vbCrLf & _
            "The management feature to maintain marketing settings was " & vbCrLf & _
            "moved from the Maintenance program to the Manager program.  " & vbCrLf & _
            "This was an attempt to allow access to this feature by none " & vbCrLf & _
            "maintenance staff.                                          "
    Case 51: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "07/01/02" & vbCrLf & vbCrLf & _
            "FIX: Unlock code freedom" & vbCrLf & vbCrLf & _
            "To bring all clients into compliance with the new Unlock " & vbCrLf & _
            "Code requirement, basic default settings were removed.  " & vbCrLf & _
            "This required clients to be issued with codes.              "
    Case 50: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "06/01/02" & vbCrLf & vbCrLf & _
            "NEW: New Program provision" & vbCrLf & vbCrLf & _
            "With the forthcoming release of the Configuration program " & vbCrLf & _
            "provision was needed to make this work with the loader " & vbCrLf & _
            "program.                                                    "
    Case 47: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "02/01/02" & vbCrLf & vbCrLf & _
            "NEW: More system colour support" & vbCrLf & vbCrLf & _
            "Added some extra colour to the Splash screen. This required " & vbCrLf & _
            "different colours depending on how many colours a PC could " & vbCrLf & _
            "support.                                                    "
    Case 57: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "30/12/01" & vbCrLf & vbCrLf & _
            "FIX: Static.ini dependency removed" & vbCrLf & vbCrLf & _
            "Finally removed all dependency on the Static.ini file.  " & vbCrLf & _
            "This has been replaced, as it was unsafe and prone to abuse " & vbCrLf & _
            "and misuse.                                                 "
    Case 62: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "30/12/01" & vbCrLf & vbCrLf & _
            "FIX: GPF when closing program" & vbCrLf & vbCrLf & _
            "Added Unhook in various places to stop GPFs on system " & vbCrLf & _
            "abnormal shutdowns.                                         "
    Case 61: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "27/12/01" & vbCrLf & vbCrLf & _
            "NEW: Advanced Locking" & vbCrLf & vbCrLf & _
            "A new feature that provided accurate record locking was " & vbCrLf & _
            "introduced.  This was developed by adding locking " & vbCrLf & _
            "information to the users account.  This is then refreshed " & vbCrLf & _
            "every time a login occurs or a lock is requested.  As this " & vbCrLf & _
            "is dependent on Windows, it is presumed to be much safer.   "
    Case 22: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "18/12/01" & vbCrLf & vbCrLf & _
            "NEW: Extra options on Customer select" & vbCrLf & vbCrLf & _
            "Some extra text boxes to allow better agent searching was " & vbCrLf & _
            "added.  Also an index search was added for customer numbers " & vbCrLf & _
            "to provided better performance.                             "
    Case 63: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "18/12/01" & vbCrLf & vbCrLf & _
            "NEW: Extra options on Customer select" & vbCrLf & vbCrLf & _
            "Some extra text boxes to allow better agent searching was " & vbCrLf & _
            "added.  Also an index search was added for customer numbers " & vbCrLf & _
            "to provided better performance.                             "
    Case 58: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "13/12/01" & vbCrLf & vbCrLf & _
            "FIX: Quality reporting problems" & vbCrLf & vbCrLf & _
            "Depending on the type of printer that was being used, when " & vbCrLf & _
            "using quality reporting some of the layout didn't print.  " & vbCrLf & _
            "This was caused by a VB bug.  A safer method was used in " & vbCrLf & _
            "its stead.                                                  "
    Case 59: lstrMsg = "PROGRAM: MMANAGER" & String(3, vbTab) & "12/12/01" & vbCrLf & vbCrLf & _
            "NEW: Label printing work started" & vbCrLf & vbCrLf & _
            "New routes into quality reporting were added to cater for " & vbCrLf & _
            "label layouts.                                              "
    Case 60: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "06/12/01" & vbCrLf & vbCrLf & _
            "FIX: Currency by PC" & vbCrLf & vbCrLf & _
            "Currency settings that were previously determined by a " & vbCrLf & _
            "system wide setting were changed to PC specific.  This may " & vbCrLf & _
            "sound like a problem, but an extra feature was also added " & vbCrLf & _
            "which halts the system if the settings don't match.  This " & vbCrLf & _
            "occurs when the system is first started.                    "
    Case 64: lstrMsg = "PROGRAM: MCLIENT" & String(3, vbTab) & "06/12/01" & vbCrLf & vbCrLf & _
            "FIX: Grid currency and safety" & vbCrLf & vbCrLf & _
            "Grids screen were made safer.  Previously it had been " & vbCrLf & _
            "possible for an agent to modify entries in grids.  When " & vbCrLf & _
            "this involved currency values this was potentially " & vbCrLf & _
            "dangerous.                                                  "
    Case 11: lstrMsg = "PROGRAM: MMOSALL" & String(3, vbTab) & "22/10/01" & vbCrLf & vbCrLf & _
            "NEW: Extended Advice memo to 100 chars  " & vbCrLf & vbCrLf & _
            "Extended Advice note 100 memo to chars & wrap function, by " & vbCrLf & _
            "request.                                                    "
    End Select

    MsgBox lstrMsg, vbInformation, gconstrTitlPrefix & "New Feature Information " & "(" & plngItem & ")"

End Sub

