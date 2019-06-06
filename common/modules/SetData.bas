Attribute VB_Name = "modSetupData"
Option Explicit

Global gconintStaticCompanyName As Integer 
Global gconintStaticCompanyTelNum As Integer 
Global gconintStaticFinanceVAT As Integer 
Global gconintStaticFinanceCSLn2 As Integer 
Global gconintStaticPFAlphaPref As Integer 
Global gconintStaticPFHalconCust As Integer 
'---- from modConfig ---
Type ConfigSettings
    strServerPath As String
    strLocalPath As String
    strCDPath As String
End Type
Global gstrConfigSettings As ConfigSettings
'---- from modConfig ---
'---- from modSetup ----
Type SystemLists
    strListName     As String
    lngSeqNum       As Long
    strListCode     As String
    strExampleDesc  As String
    strTopic        As String
    strDescValue    As String
End Type
Global gstrSystemLists(22) As SystemLists
Global gstrSystemListsTech(27) As SystemLists

Type SystemListsFull
    strListName     As String
    lngSeqNum       As Long
    strListCode     As String
    strExampleDesc  As String
    strTopic        As String
    strDescValue    As String
    strUserDef1     As String
    strUserDef2     As String
    booInUse        As Boolean 
End Type
Global gstrSystemListsGeneralFull(7) As SystemListsFull
Global gstrSystemListsGeneral(52) As SystemListsFull

Type ListDetailChildMulti
    lngSeqNum       As Long
    strListCode     As String
    strDescValue    As String
    strUserDef1     As String
    strUserDef2     As String
End Type
Type ListDetailParentMulti
    strListName     As String
    strTopic        As String
    strDetail(10)   As ListDetailChildMulti
End Type
Global gstrSystemListsMultiFull(25) As ListDetailParentMulti
'---- from modSetup ----

Type ListMasterRecord
    strListName     As String
    booSysUse       As Boolean
End Type
Global mstrListerMasterRecord(27) As ListMasterRecord
Sub FillSystemListsGeneralDefs()
Dim lintCounter As Integer
    
    If UCase$(App.ProductName) <> "LITE" Then
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Consignment Status"
            .lngSeqNum = 0
            .strListCode = "A"
            .strDescValue = "Awaiting"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Consignment Status"
            .lngSeqNum = 1
            .strListCode = "X"
            .strDescValue = "Cancelled"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Consignment Status"
            .lngSeqNum = 2
            .strListCode = "P"
            .strDescValue = "Printed"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Consignment Status"
            .lngSeqNum = 3
            .strListCode = "D"
            .strDescValue = "Downloaded"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Consignment Status"
            .lngSeqNum = 4
            .strListCode = "C"
            .strDescValue = "Courier Picked up"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
    End If
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 0
        .strListCode = "1"
        .strDescValue = "1-January"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 1
        .strListCode = "2"
        .strDescValue = "2-February"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 2
        .strListCode = "3"
        .strDescValue = "3-March"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 3
        .strListCode = "4"
        .strDescValue = "4-April"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 4
        .strListCode = "5"
        .strDescValue = "5-May"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 5
        .strListCode = "6"
        .strDescValue = "6-June"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 6
        .strListCode = "7"
        .strDescValue = "7-July"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 7
        .strListCode = "8"
        .strDescValue = "8-August"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 8
        .strListCode = "9"
        .strDescValue = "9-September"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 9
        .strListCode = "10"
        .strDescValue = "10-October"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 10
        .strListCode = "11"
        .strDescValue = "11-November"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Months"
        .lngSeqNum = 11
        .strListCode = "12"
        .strDescValue = "12-December"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 0
        .strListCode = "C"
        .strDescValue = "Confirmed"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 1
        .strListCode = "X"
        .strDescValue = "Cancelled"
        .strUserDef1 = "CAN CANCEL" 
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 2
        .strListCode = "A"
        .strDescValue = "Awaiting Packing"
        .strUserDef1 = "CAN CANCEL" 
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 3
        .strListCode = "S"
        .strDescValue = "Cancelled with Order OOS"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 4
        .strListCode = "B"
        .strDescValue = "Confirmed Items OOS"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        'These order statuses aren't of use unless you have a Manager or maint prog
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Order Status"
            .lngSeqNum = 5
            .strListCode = "D"
            .strDescValue = "Stock Downloaded"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "Order Status"
            .lngSeqNum = 5
            .strListCode = "E"
            .strDescValue = "Stock Downloaded with OOS"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
    End With
    End If
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 6
        .strListCode = "P"
        .strDescValue = "Printed & Awaiting to be Packed"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Status"
        .lngSeqNum = 7
        .strListCode = "H"
        .strDescValue = "On Hold Awaiting Authorisation"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Type"
        .lngSeqNum = 0
        .strListCode = "1"
        .strDescValue = "Regular Order"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Order Type"
        .lngSeqNum = 1
        .strListCode = "3"
        .strDescValue = "Replacement, No charge"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Payment Method"
        .lngSeqNum = 0
        .strListCode = "C"
        .strDescValue = "Credit Card"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Payment Method"
        .lngSeqNum = 1
        .strListCode = "Q"
        .strDescValue = "Cheque"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Payment Method"
        .lngSeqNum = 2
        .strListCode = "V"
        .strDescValue = "Voucher"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 0
            .strListCode = "IRD1"
            .strDescValue = "Special delivery Instructions, Line 1"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 1
            .strListCode = "IRD2"
            .strDescValue = "Special delivery Instructions, Line 2"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 2
            .strListCode = "IRD3"
            .strDescValue = "Special delivery Instructions, Line 3"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 3
            .strListCode = "IRD4"
            .strDescValue = "Special delivery Instructions, Line 4"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 4
            .strListCode = "JBIR"
            .strDescValue = "Booking in required"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 5
            .strListCode = "JPOD"
            .strDescValue = "Poof of delivery required"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Notification Code"
            .lngSeqNum = 6
            .strListCode = "JPAL"
            .strDescValue = "Pallet service"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Prepaid Indicator"
            .lngSeqNum = 0
            .strListCode = "P"
            .strDescValue = "Prepaid"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Prepaid Indicator"
            .lngSeqNum = 1
            .strListCode = " "
            .strDescValue = "Not Prepaid"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
    End If
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 7
        .strListCode = "SC0"
        .strDescValue = "Courier pack by 10:00 next day"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 8
        .strListCode = "SC2"
        .strDescValue = "Courier pack by 12:00 next day"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 9
        .strListCode = "SCD"
        .strDescValue = "Courier pack next day delivery"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 10
        .strListCode = "SC0P"
        .strDescValue = "Courier pack by 10:00 next day pre-paid"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 11
        .strListCode = "SC2P"
        .strDescValue = "Courier pack by 12:00 next day pre-paid"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 12
        .strListCode = "SCDP"
        .strDescValue = "Courier pack by next day pre-paid"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Weekend Handling Code"
            .lngSeqNum = 0
            .strListCode = "ECSA"
            .strDescValue = "Saturday collection"
            .strUserDef1 = " "
            .strUserDef2 = " "
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Weekend Handling Code"
            .lngSeqNum = 1
            .strListCode = "ECSU"
            .strDescValue = "Sunday collection"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsGeneral(lintCounter)
            .strListName = "PForce Weekend Handling Code"
            .lngSeqNum = 2
            .strListCode = "ESAT"
            .strDescValue = "Saturday delivery"
            .strUserDef1 = " "
            .strUserDef2 = " "
            .booInUse = True 
        End With
    End If
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Y or N"
        .lngSeqNum = 0
        .strListCode = "Y"
        .strDescValue = "Yes"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "Y or N"
        .lngSeqNum = 1
        .strListCode = "N"
        .strDescValue = "No"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = True 
    End With

    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "System Various"
        .lngSeqNum = 1
        .strListCode = "DONAVAIL"
        .strDescValue = "False"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneral(lintCounter)
        .strListName = "System Various"
        .lngSeqNum = 1
        .strListCode = "STCKTHRE"
        .strDescValue = "False"
        .strUserDef1 = " "
        .strUserDef2 = " "
        .booInUse = False 
    End With

End Sub

Sub xFillSystemListsFullDefs()
Dim lintCounter As Integer

    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 0
        .strListCode = "SUP"
        .strDescValue = "Guaranteed 2 day service"
        .strUserDef1 = "48"
        .strUserDef2 = "�3.50" 
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 1
        .strListCode = "SND"
        .strDescValue = "Guaranteed next day delivery"
        .strUserDef1 = "24"
        .strUserDef2 = "�4.00" 
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 2
        .strListCode = "S72"
        .strDescValue = "Guaranteed 3 day service"
        .strUserDef1 = "72"
        .strUserDef2 = "�2.90" 
        .booInUse = True 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 3
        .strListCode = "SMS"
        .strDescValue = "Service master"
        .strUserDef1 = "SM"
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 4
        .strListCode = "S09"
        .strDescValue = "Guaranteed 09:00 next day delivery"
        .strUserDef1 = "09"
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 5
        .strListCode = "S10"
        .strDescValue = "Guaranteed 10:00 next day delivery"
        .strUserDef1 = "10"
        .strUserDef2 = " "
        .booInUse = False 
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsGeneralFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .lngSeqNum = 6
        .strListCode = "S12"
        .strDescValue = "Guaranteed 12:00 next day delivery"
        .strUserDef1 = "12"
        .strUserDef2 = " "
        .booInUse = False 
    End With

End Sub
Sub FillSystemListsTechDefs()
Dim lintCounter As Integer
    
    If UCase$(App.ProductName) <> "LITE" Then
        With gstrSystemListsTech(lintCounter)
            .strListName = "DB"
            .lngSeqNum = 0
            .strListCode = "Central"
            .strDescValue = "Central.mdb"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "DB"
            .lngSeqNum = 0
            .strListCode = "Local"
            .strDescValue = "Local.mdb"
        End With
    End If
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "DB"
            .lngSeqNum = 1
            .strListCode = "LocalTest"
            .strDescValue = "LocalTest.mdb"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "DB"
            .lngSeqNum = 3
            .strListCode = "CentraTest"
            .strDescValue = "CentralTest.mdb"
        End With
    End If
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 0
            .strListCode = "ProgCount"
            .strDescValue = "4"
        End With
        
        'Prog1
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 1
            .strListCode = "Prog1Desc"
            .strDescValue = "Mindwarp MOS Client"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 1
            .strListCode = "Prog1"
            .strDescValue = "mmos.exe"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 1
            .strListCode = "Prog1Param"
            .strDescValue = "X"
        End With
        
        'Prog2
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 2
            .strListCode = "Prog2"
            .strDescValue = "madmin.exe"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 2
            .strListCode = "Prog2Param"
            .strDescValue = "ADMIN"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 2
            .strListCode = "Prog2Desc"
            .strDescValue = "Mindwarp MOS Admin"
        End With
        
        'Prog 3
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 3
            .strListCode = "Prog3"
            .strDescValue = "mreps.exe"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 3
            .strListCode = "Prog3Param"
            .strDescValue = "REPORT"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 3
            .strListCode = "Prog3Desc"
            .strDescValue = "Mindwarp MOS Reporting"
        End With
        
        'Prog 4
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 4
            .strListCode = "Prog4"
            .strDescValue = "mconf.exe"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 4
            .strListCode = "Prog4Param"
            .strDescValue = "CONFIG"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "Programs"
            .lngSeqNum = 4
            .strListCode = "Prog4Desc"
            .strDescValue = "Mindwarp MOS Configuration"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 0
            .strListCode = "AppPath"
            .strDescValue = gstrConfigSettings.strLocalPath
        End With
            
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 1
            .strListCode = "ServerPath"
            .strDescValue = gstrConfigSettings.strServerPath
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 3
            .strListCode = "SrvTestPth"
            .strDescValue = gstrConfigSettings.strServerPath & "TestNew\"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 4
            .strListCode = "SuppPath"
            .strDescValue = gstrConfigSettings.strServerPath & "Setup\Support\"
        End With
        
        lintCounter = lintCounter + 1
        With gstrSystemListsTech(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 5
            .strListCode = "SupTestPth"
            .strDescValue = gstrConfigSettings.strServerPath & "TestNew\Setup\Support\"
        End With
    End If
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 0
        .strListCode = "10"
        .strDescValue = "Distribution"
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 1
        .strListCode = "20"
        .strDescValue = "Order Entry"
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 2
        .strListCode = "30"
        .strDescValue = "Sales"
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 3
        .strListCode = "40"
        .strDescValue = "Accounts"
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 4
        .strListCode = "50"
        .strDescValue = "General Managers"
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsTech(lintCounter)
        .strListName = "User Levels"
        .lngSeqNum = 5
        .strListCode = "99"
        .strDescValue = "Information Systems"
    End With
    
End Sub
Sub UpdateListValues(pbooDBOpen As Boolean, Optional pobjStatusbar As Variant, Optional pintStatusSegment As Integer)
Dim lintArrInc As Integer
Dim lstrSQL As String

    If pbooDBOpen = False Then
        Set gdatCentralDatabase = OpenDatabase(gstrConfigSettings.strServerPath & "Central.mdb", , False)
    End If
    
    'Create Lists Master Records
    For lintArrInc = 0 To UBound(mstrListerMasterRecord)
        With mstrListerMasterRecord(lintArrInc)

            If Trim$(.strListName) <> "" Then
                lstrSQL = "INSERT INTO " & gtblMasterLists & " " & _
                    "( ListName, SysUse ) Values ( '" & Trim$(.strListName) & _
                    "', " & .booSysUse & ");"
                gdatCentralDatabase.Execute lstrSQL
            End If
        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
    'Create List Detial records for frmStaticX(s) screens
    For lintArrInc = 0 To UBound(gstrSystemLists)
        With gstrSystemLists(lintArrInc)
            lstrSQL = "INSERT INTO " & gtblMasterListDetails & " ( SequenceNum, " & _
                "ListCode, Description, InUse, ListNum ) SELECT " & .lngSeqNum & " AS Expr1, '" & .strListCode & _
                "' AS Expr2, '" & .strDescValue & "' AS Expr3, True AS Expr4, " & gtblMasterLists & ".ListNum as lm " & _
                "FROM " & gtblMasterLists & " WHERE (((" & gtblMasterLists & ".ListName)='" & .strListName & "'));"
            gdatCentralDatabase.Execute lstrSQL
        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
    'Create Derived reList Details (file location) also none derived, user levels etc
    For lintArrInc = 0 To UBound(gstrSystemListsTech)
        With gstrSystemListsTech(lintArrInc)
            If Trim$(.strListName) <> "" Then
                lstrSQL = "INSERT INTO " & gtblMasterListDetails & " ( SequenceNum, " & _
                    "ListCode, Description, InUse, ListNum ) SELECT " & .lngSeqNum & " AS Expr1, '" & .strListCode & _
                    "' AS Expr2, '" & .strDescValue & "' AS Expr3, True AS Expr4, " & gtblMasterLists & ".ListNum as lm " & _
                    "FROM " & gtblMasterLists & " WHERE (((" & gtblMasterLists & ".ListName)='" & .strListName & "'));"
                gdatCentralDatabase.Execute lstrSQL
            End If
        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
    'Create Other stuff like Parcel Force statics & order Status
    For lintArrInc = 0 To UBound(gstrSystemListsGeneral)
        With gstrSystemListsGeneral(lintArrInc)

            lstrSQL = "INSERT INTO " & gtblMasterListDetails & " ( SequenceNum, " & _
                "ListCode, Description, InUse, ListNum, UserDef1, UserDef2 ) SELECT " & _
                .lngSeqNum & " AS Expr1, '" & .strListCode & _
                "' AS Expr2, '" & .strDescValue & "' AS Expr3, " & .booInUse & " AS Expr4, " & _
                "" & gtblMasterLists & ".ListNum as lm, '" & .strUserDef1 & "' as ud1, '" & _
                .strUserDef2 & "' as ud2 " & _
                "FROM " & gtblMasterLists & " WHERE (((" & gtblMasterLists & ".ListName)='" & .strListName & "'));"
            gdatCentralDatabase.Execute lstrSQL

        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
    For lintArrInc = 0 To UBound(gstrSystemListsGeneralFull)
        With gstrSystemListsGeneralFull(lintArrInc)

            lstrSQL = "INSERT INTO " & gtblMasterListDetails & " ( SequenceNum, " & _
                "ListCode, Description, InUse, ListNum, UserDef1, UserDef2 ) SELECT " & _
                .lngSeqNum & " AS Expr1, '" & .strListCode & _
                "' AS Expr2, '" & .strDescValue & "' AS Expr3, " & .booInUse & " AS Expr4, " & _
                "" & gtblMasterLists & ".ListNum as lm, '" & .strUserDef1 & "' as ud1, '" & _
                .strUserDef2 & "' as ud2 " & _
                "FROM " & gtblMasterLists & " WHERE (((" & gtblMasterLists & ".ListName)='" & .strListName & "'));"
            gdatCentralDatabase.Execute lstrSQL
        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
    Dim lintArrInc2 As Integer
    
    For lintArrInc = 0 To UBound(gstrSystemListsMultiFull)
        With gstrSystemListsMultiFull(lintArrInc)
            For lintArrInc2 = 0 To 10
                If .strDetail(lintArrInc2).strListCode <> "" Then
                    lstrSQL = "INSERT INTO " & gtblMasterListDetails & " ( SequenceNum, " & _
                        "ListCode, Description, InUse, ListNum, UserDef1, UserDef2 ) SELECT " & _
                        .strDetail(lintArrInc2).lngSeqNum & " AS Expr1, '" & .strDetail(lintArrInc2).strListCode & _
                        "' AS Expr2, '" & .strDetail(lintArrInc2).strDescValue & "' AS Expr3, True AS Expr4, " & _
                        "" & gtblMasterLists & ".ListNum as lm, '" & .strDetail(lintArrInc2).strUserDef1 & "' as ud1, '" & _
                        .strDetail(lintArrInc2).strUserDef2 & "' as ud2 " & _
                        "FROM " & gtblMasterLists & " WHERE (((" & gtblMasterLists & ".ListName)='" & .strListName & "'));"
                    gdatCentralDatabase.Execute lstrSQL
                End If
            Next lintArrInc2
        End With
    Next lintArrInc
    If Not IsMissing(pobjStatusbar) Then
        pobjStatusbar.Value = pobjStatusbar.Value + pintStatusSegment
    End If
    
End Sub
Sub FillListMasterSystemDefs()
Dim lintCounter As Integer

    With mstrListerMasterRecord(lintCounter)
        .strListName = "Account Status"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Account Type"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Courier, Postage & Handling"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Credit Card Type"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Media Codes"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Order Code"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Product Classes"
        .booSysUse = False
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "User Levels"
        .booSysUse = True
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Amount Levels"
        .booSysUse = True
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "DB"
            .booSysUse = True
        End With
        
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "SysFileInfo"
            .booSysUse = True
        End With
    
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "PForce Consignment Range"
            .booSysUse = True
        End With
        
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "PForce Contract Details"
            .booSysUse = True
        End With
        
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "Programs"
            .booSysUse = True
        End With
    End If
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Company Address"
        .booSysUse = True
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "Card Serv Header"
        End With
        
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "Consignment Status"
            .booSysUse = True
        End With
    End If
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Months"
        .booSysUse = True
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Order Status"
        .booSysUse = True
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Order Type"
        .booSysUse = True
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Payment Method"
        .booSysUse = True
    End With
    
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "PForce Notification Code"
            .booSysUse = True
        End With
        
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "PForce Prepaid Indicator"
            .booSysUse = True
        End With
    End If
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "PForce Service Indicator"
        .booSysUse = False
    End With
        
    If UCase$(App.ProductName) <> "LITE" Then
        lintCounter = lintCounter + 1
        With mstrListerMasterRecord(lintCounter)
            .strListName = "PForce Weekend Handling Code"
            .booSysUse = True
        End With
    End If
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "Y or N"
        .booSysUse = True
    End With
    
    lintCounter = lintCounter + 1
    With mstrListerMasterRecord(lintCounter)
        .strListName = "System Various"
        .booSysUse = False
    End With
    
    
End Sub
Sub FillSystemListsDefs()
Dim lintCounter As Integer
'From modSetup

    With gstrSystemLists(lintCounter)
        .strListName = "Amount Levels"
        .lngSeqNum = 1
        .strListCode = "VAT"
        .strExampleDesc = "17.5"
        .strTopic = "Vat"
        gconintStaticFinanceVAT = lintCounter 
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Amount Levels"
        .lngSeqNum = 0
        .strListCode = "DENOM"
        .strExampleDesc = "�"
        .strTopic = "Denomination"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Amount Levels"
        .lngSeqNum = 3
        .strListCode = "POWAIVER"
        .strExampleDesc = "�75"
        .strTopic = "Postage Waiver / Card Authorisation level"
    End With: lintCounter = lintCounter + 1
    
    If UCase$(App.ProductName) <> "LITE" Then
        With gstrSystemLists(lintCounter)
            .strListName = "Card Serv Header"
            .lngSeqNum = 1
            .strListCode = "CSHEAD1A"
            .strExampleDesc = "MIDLAND CARD SERV CHARGE-CARD CLAIMS My Company "
            .strTopic = "Card Server Header" & vbCrLf & "line 1 (first half)"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "Card Serv Header"
            .lngSeqNum = 2
            .strListCode = "CSHEAD1B"
            .strExampleDesc = "Plc. - Merchant No. 1234"
            .strTopic = "Card Server Header" & vbCrLf & "line 1 (second half)"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "Card Serv Header"
            .lngSeqNum = 3
            .strListCode = "CSHEAD2A"
            .strExampleDesc = "CONTACT - Joe Bloggs  Telephone 0123 4567890"
            .strTopic = "Card Server Header" & vbCrLf & "line 2"
            gconintStaticFinanceCSLn2 = lintCounter 
        End With: lintCounter = lintCounter + 1
    End If
    
    'frmCompany
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 1
        .strListCode = "CONAME"
        .strExampleDesc = "My Company"
        .strTopic = "Company Name"
        gconintStaticCompanyName = lintCounter 
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 1
        .strListCode = "COCONTA"
        .strExampleDesc = "Joe Bloggs"
        .strTopic = "Company Contact"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 2
        .strListCode = "COADLI1"
        .strExampleDesc = "Kingsway"
        .strTopic = "Address Line 1"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 3
        .strListCode = "COADLI2"
        .strExampleDesc = "My Town"
        .strTopic = "Address Line 2"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 4
        .strListCode = "COADLI3"
        .strExampleDesc = "Tyne && Wear"
        .strTopic = "Address Line 3"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 5
        .strListCode = "COADLI4"
        .strExampleDesc = "AB11 0DE"
        .strTopic = "Address Line 4"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 6
        .strListCode = "COADLI5"
        .strExampleDesc = "UK"
        .strTopic = "Address Line 5"
    End With: lintCounter = lintCounter + 1
    
    With gstrSystemLists(lintCounter)
        .strListName = "Company Address"
        .lngSeqNum = 6
        .strListCode = "COTELEP"
        .strExampleDesc = "0123 4567890"
        .strTopic = "Company Telephone"
        gconintStaticCompanyTelNum = lintCounter 
    End With: lintCounter = lintCounter + 1
    
    If UCase$(App.ProductName) <> "LITE" Then
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Consignment Range"
            .lngSeqNum = 0
            .strListCode = "ALPHAPREF"
            .strExampleDesc = "jk"
            .strTopic = "Range Alpha Prefix"
            gconintStaticPFAlphaPref = lintCounter 
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Consignment Range"
            .lngSeqNum = 1
            .strListCode = "START"
            .strExampleDesc = "7900009"
            .strTopic = "Range Start Number"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Consignment Range"
            .lngSeqNum = 2
            .strListCode = "END"
            .strExampleDesc = "8199993"
            .strTopic = "Range Last Number"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Contract Details"
            .lngSeqNum = 0
            .strListCode = "ContractNo"
            .strExampleDesc = "H677329"
            .strTopic = "Contract Number"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Contract Details"
            .lngSeqNum = 1
            .strListCode = "AcctNo"
            .strExampleDesc = "TRA3543"
            .strTopic = "Account Number"
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Contract Details"
            .lngSeqNum = 2
            .strListCode = "HalconCoID"
            .strExampleDesc = "1"
            .strTopic = "Halcon Co ID"
            gconintStaticPFHalconCust = lintCounter 
        End With: lintCounter = lintCounter + 1
        
        With gstrSystemLists(lintCounter)
            .strListName = "PForce Contract Details"
            .lngSeqNum = 3
            .strListCode = "HCustCode"
            .strExampleDesc = "TRA3543"
            .strTopic = "Halcon Customer Code"
        End With: lintCounter = lintCounter + 1
    
        With gstrSystemLists(lintCounter)
            .strListName = "SysFileInfo"
            .lngSeqNum = 0
            .strListCode = "PFEFile"
            .strExampleDesc = "c:\Pforce.txt"
        End With
    End If
    
End Sub
Sub FillListsMultiDef()
Dim lintCounter As Integer
'from modSetup

    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Account Status"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "DECEASED"
        .strDetail(0).strDescValue = "Deceased"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "GONE"
        .strDetail(1).strDescValue = "Gone away and no New Address"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
        .strDetail(2).lngSeqNum = 4
        .strDetail(2).strListCode = "NONE"
        .strDetail(2).strDescValue = "No Mailings"
        .strDetail(2).strUserDef1 = " "
        .strDetail(2).strUserDef2 = " "
    End With
        
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Account Type"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "T"
        .strDetail(0).strDescValue = "Trade"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "G"
        .strDetail(1).strDescValue = "General Public"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Credit Card Type"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "VISA"
        .strDetail(0).strDescValue = "Visa"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "MASTER"
        .strDetail(1).strDescValue = "Master Card"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
        .strDetail(2).lngSeqNum = 2
        .strDetail(2).strListCode = "SWITCH"
        .strDetail(2).strDescValue = "Switch"
        .strDetail(2).strUserDef1 = " "
        .strDetail(2).strUserDef2 = " "
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Order Code"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "E"
        .strDetail(0).strDescValue = "Email"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "P"
        .strDetail(1).strDescValue = "Phone"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
        .strDetail(2).lngSeqNum = 2
        .strDetail(2).strListCode = "F"
        .strDetail(2).strDescValue = "Fax"
        .strDetail(2).strUserDef1 = " "
        .strDetail(2).strUserDef2 = " "
        .strDetail(3).lngSeqNum = 3
        .strDetail(3).strListCode = "O"
        .strDetail(3).strDescValue = "Post"
        .strDetail(3).strUserDef1 = " "
        .strDetail(3).strUserDef2 = " "
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Media Codes"
        .strDetail(0).lngSeqNum = 1
        .strDetail(0).strListCode = "LOPA"
        .strDetail(0).strDescValue = "Local Paper"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 2
        .strDetail(1).strListCode = "YELLO"
        .strDetail(1).strDescValue = "Yellow Pages"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
    End With
        
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Product Classes"
        .strDetail(0).lngSeqNum = 42
        .strDetail(0).strListCode = "95"
        .strDetail(0).strDescValue = "Samples"
        .strDetail(0).strUserDef1 = "50630"
        .strDetail(0).strUserDef2 = "OTH"
    End With

    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "Courier, Postage & Handling"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "ANC"
        .strDetail(0).strDescValue = "ANC"
        .strDetail(0).strUserDef1 = " "
        .strDetail(0).strUserDef2 = " "
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "PF24"
        .strDetail(1).strDescValue = "Parcel Force 24 hrs"
        .strDetail(1).strUserDef1 = "24"
        .strDetail(1).strUserDef1 = " "
        .strDetail(1).strUserDef2 = " "
        .strDetail(2).lngSeqNum = 2
        .strDetail(2).strListCode = "PF48"
        .strDetail(2).strDescValue = "Parcel Force 48 hrs"
        .strDetail(2).strUserDef1 = "48"
        .strDetail(2).strUserDef1 = " "
        .strDetail(2).strUserDef2 = " "
    End With
    
    lintCounter = lintCounter + 1
    With gstrSystemListsMultiFull(lintCounter)
        .strListName = "PForce Service Indicator"
        .strDetail(0).lngSeqNum = 0
        .strDetail(0).strListCode = "SUP"
        .strDetail(0).strDescValue = "Guaranteed 2 day service"
        .strDetail(0).strUserDef1 = "48"
        .strDetail(0).strUserDef2 = "�3.50"
        
        .strDetail(1).lngSeqNum = 1
        .strDetail(1).strListCode = "SND"
        .strDetail(1).strDescValue = "Guaranteed next day delivery"
        .strDetail(1).strUserDef1 = "24"
        .strDetail(1).strUserDef2 = "�4.00"
    
        .strDetail(2).lngSeqNum = 2
        .strDetail(2).strListCode = "S72"
        .strDetail(2).strDescValue = "Guaranteed 3 day service"
        .strDetail(2).strUserDef1 = "72"
        .strDetail(2).strUserDef2 = "�2.90"
    
        .strDetail(3).lngSeqNum = 3
        .strDetail(3).strListCode = "SMS"
        .strDetail(3).strDescValue = "Service master"
        .strDetail(3).strUserDef1 = "SM"
        .strDetail(3).strUserDef2 = " "
    
        .strDetail(4).lngSeqNum = 4
        .strDetail(4).strListCode = "S09"
        .strDetail(4).strDescValue = "Guaranteed 09:00 next day delivery"
        .strDetail(4).strUserDef1 = "09"
        .strDetail(4).strUserDef2 = " "
    
        .strDetail(5).lngSeqNum = 5
        .strDetail(5).strListCode = "S10"
        .strDetail(5).strDescValue = "Guaranteed 10:00 next day delivery"
        .strDetail(5).strUserDef1 = "10"
        .strDetail(5).strUserDef2 = " "
    
        .strDetail(6).lngSeqNum = 6
        .strDetail(6).strListCode = "S12"
        .strDetail(6).strDescValue = "Guaranteed 12:00 next day delivery"
        .strDetail(6).strUserDef1 = "12"
        .strDetail(6).strUserDef2 = " "
    End With
    
End Sub
Sub LiteListsSetup()

    If EmptyLists = False Then
        Exit Sub
    End If
    
    DoEvents
    
    FillSystemListsDefs
    FillSystemListsTechDefs
    FillListMasterSystemDefs
    FillSystemListsGeneralDefs
    FillListsMultiDef
    
    CopyExampToDescs
    
    UpdateListValues True
        
    AddDefaultRecords
    
End Sub
Function EmptyLists() As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
    
    On Error GoTo ErrHandler
    
    EmptyLists = False
    
    lstrSQL = "SELECT Count(" & gtblMasterLists & ".ListNum) AS CountOfListNum, " & _
        "Count(" & gtblMasterListDetails & ".ListNum) AS CountOfListNum1 " & _
        "FROM " & gtblMasterListDetails & " INNER JOIN " & gtblMasterLists & " ON " & _
        "" & gtblMasterListDetails & ".ListNum = " & gtblMasterLists & ".ListNum;"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            If Val(.Fields("CountOfListNum") & "") = 0 And _
                Val(.Fields("CountOfListNum1") & "") = 0 Then
                EmptyLists = True
            End If
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "EmptyLists", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function
Sub CopyExampToDescs()
Dim lintArrInc As Integer

    
    For lintArrInc = 0 To UBound(gstrSystemLists)
        With gstrSystemLists(lintArrInc)
            If .strDescValue = "" Then
                .strDescValue = .strExampleDesc
            End If
        End With
    Next lintArrInc
    
    For lintArrInc = 0 To UBound(gstrSystemListsTech)
        With gstrSystemListsTech(lintArrInc)
            If .strDescValue = "" Then
                .strDescValue = .strExampleDesc
            End If
        End With
    Next lintArrInc
    
    For lintArrInc = 0 To UBound(gstrSystemListsGeneral)
        With gstrSystemListsGeneral(lintArrInc)
            If .strDescValue = "" Then
                .strDescValue = .strExampleDesc
            End If
        End With
    Next lintArrInc
    
    For lintArrInc = 0 To UBound(gstrSystemListsGeneralFull)
        With gstrSystemListsGeneralFull(lintArrInc)
            If .strDescValue = "" Then
                .strDescValue = .strExampleDesc
            End If
        End With
    Next lintArrInc
    
    
End Sub
Sub AddDefaultRecords()
Dim lstrSQL As String

    lstrSQL = "INSERT INTO " & gtblSystem & " ( [Value], Item, OtherDate ) VALUES('05/03/02 09:14:38','StockUploaded',now());"
    gdatCentralDatabase.Execute lstrSQL
    
    lstrSQL = "INSERT INTO " & gtblMachine & " ( [Value], Item ) VALUES('05/02/02 09:14:38','StockDownloaded');"
    gdatLocalDatabase.Execute lstrSQL
    
    lstrSQL = "INSERT INTO " & gtblSystem & " ( [Value], Item ) VALUES(0,'BATCHINCR');"
    gdatCentralDatabase.Execute lstrSQL
    
End Sub
