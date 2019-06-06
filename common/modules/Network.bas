Attribute VB_Name = "modNetwork"
Option Explicit

Global gwrkODBC As Workspace
Global gwrkJet As Workspace
    
Dim mstrRefreshQuery As String

Enum ReturnPhaseItems
    retPhaseItemUser
    retPhaseItemSysStart
    retPhaseItemMach
End Enum

Global Const gtblAdviceNotes = "AdviceNotes"
Global Const gtblCashBook = "CashBook"
Global Const gtblCustAccounts = "CustAccounts"
Global Const gtblCustNotes = "CustNotes"
Global Const gtblCustomReports = "CustomReports"
Global Const gtblListDetails = "ListDetails"
Global Const gtblLists = "Lists"
Global Const gtblOrderLines = "Orderlines"
Global Const gtblPForce = "PForce"
Global Const gtblProducts = "Products"
Global Const gtblMachine = "Machine"

Global Const gtblRemarks = "Remarks"
Global Const gtblSubstitutions = "Substitutions"
Global Const gtblSystem = "System"
Global Const gtblUsage = "Usage"
Global Const gtblUsers = "Users"

Global Const gtblPADAvailable = "PADAvailable"
Global Const gtblPADOffice = "PADOffice"
Global Const gtblPADOpeningTimes = "PADOpening_Times"

Global Const gtblMasterListDetails = "ListDetailsMaster"
Global Const gtblMasterOrderLines = "OrderLinesMaster"
Global Const gtblMasterProducts = "ProductsMaster"
Global Const gtblMasterLists = "ListsMaster"

Global Const gtblMasterPADAvailable = "PADAvailableMaster"
Global Const gtblMasterPADOffice = "PADOfficeMaster"
Global Const gtblMasterPADOpeningTimes = "PADOpening_TimesMaster"

Sub InitDb(Optional pstrSpecificStaticLdr As String = "")
Dim lstrCentralDBInput As String
Dim lstrLocalDBInput As String
Dim lstrCentralTrainingDBInput As String
Dim lvarErrorStage As Variant
    
    On Error GoTo ErrorHandler
    
    Dim lstrStaticLdr As String: lstrStaticLdr = Trim$(App.Path) & "\" & gconstrStaticLdr
    If pstrSpecificStaticLdr <> "" Then
        lstrStaticLdr = pstrSpecificStaticLdr
    End If
    
    lvarErrorStage = 140

    If UCase(Dir(lstrStaticLdr)) = UCase(gconstrStaticLdr) Then
        CheckStaticCipher pstrSpecificStaticLdr
        gstrStatic.strUnlockCode = ""
    Else
        lvarErrorStage = 150
        MsgBox "Please check your access to the network!, " & _
        "A required file is missing!", , gconstrTitlPrefix & "Init DB"
        Unhook
        End
    End If
    
    With gstrStatic
        If InStr(UCase(Command$), "/TEST") = 0 Then
            If Trim$(Dir(.strCentralDBFile)) = "" Or Trim$(.strCentralDBFile) = "" Then
                lvarErrorStage = 10
                MsgBox "Please check your access to the network!, " & _
                "if in difficulty consult your technical support office!" & vbCrLf & vbCrLf & _
                "System closing!", , gconstrTitlPrefix & "Init DB"
                Unhook
                End
            End If
            If Trim$(Dir(.strLocalDBFile)) = "" Or Trim$(.strLocalDBFile) = "" Then
                lvarErrorStage = 20
                MsgBox "Please check your access to the network!, " & _
                "if in difficulty consult your technical support office!" & vbCrLf & vbCrLf & _
                "System closing!", , gconstrTitlPrefix & "Init DB"
                End
            End If
        End If
            
        If InStr(UCase(Command$), "/TEST") > 0 Then
            If Trim$(Dir(.strCentralTestingDBFile)) = "" Or Trim$(.strCentralTestingDBFile) = "" Then
                lvarErrorStage = 30
                MsgBox "Please check your access to the network!, " & _
                "if in difficulty consult your technical support office!" & vbCrLf & vbCrLf & _
                "System closing!", , gconstrTitlPrefix & "Init DB"
                End
            End If
            If Trim$(Dir(.strLocalTestingDBFile)) = "" Or Trim$(.strLocalTestingDBFile) = "" Then
                lvarErrorStage = 40
                MsgBox "Please check your access to the network!, " & _
                "if in difficulty consult your technical support office!" & vbCrLf & vbCrLf & _
                "System closing!", , gconstrTitlPrefix & "Init DB"
                End
            End If
        End If
        
        'Connect to DB's
        If InStr(UCase(Command$), "/TEST") > 0 Then
            gstrUserMode = gconstrTestingMode
            lvarErrorStage = 55
            'Refresh User table against locking file LDB
            mstrRefreshQuery = RefreshUserPhase(.strCentralTestingDBFile)
            
            If gbooSQLServerInUse = True Then
                Set gwrkJet = CreateWorkspace("", "admin", "", dbUseJet)
                Set gwrkODBC = CreateWorkspace("NewODBCWorkspace", "admin", "", dbUseODBC)
                Set gdatCentralDatabase = gwrkODBC.OpenConnection("Connection3", _
                    dbDriverCompleteRequired, True, _
                        "ODBC;DATABASE=Mmos;DSN=Mmos;")
                Set gdatLocalDatabase = OpenDatabase(.strLocalTestingDBFile, , False)
            Else
                If gstrSystemRoute = srCompanyRoute Then
                    lvarErrorStage = 60
                    Set gdatLocalDatabase = OpenDatabase(.strLocalTestingDBFile, , False)
                        lvarErrorStage = 70
                    Set gdatCentralDatabase = OpenDatabase(.strCentralTestingDBFile, , False)
                Else
                    lvarErrorStage = 80
                    Set gdatLocalDatabase = OpenDatabase(.strLocalTestingDBFile, _
                        dbDriverComplete, False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
                    lvarErrorStage = 90
                    Set gdatCentralDatabase = OpenDatabase(.strCentralTestingDBFile, _
                        dbDriverComplete, False, Trim$(gstrDBPasswords.strCentralDBPasswordString))
                End If
            End If
            MsgBox vbTab & "You are now in the Testing Environment." & vbCrLf & vbCrLf & _
                "Consequently you have been switched to a copy " & _
                "of the main Database, " & vbCrLf & _
                "so your testing will not be done in the Live Database.", , gconstrTitlPrefix & "Init DB"
        Else
            gstrUserMode = gconstrLiveMode
            lvarErrorStage = 95
            'Refresh User table against locking file LDB
            mstrRefreshQuery = RefreshUserPhase(.strCentralDBFile)
            
            If gstrSystemRoute = srCompanyRoute Then
                lvarErrorStage = 100
                Set gdatLocalDatabase = OpenDatabase(.strLocalDBFile, , False)
                lvarErrorStage = 110
                Set gdatCentralDatabase = OpenDatabase(.strCentralDBFile, , False)
            Else
                lvarErrorStage = 120
                Set gdatLocalDatabase = OpenDatabase(.strLocalDBFile, _
                    dbDriverComplete, False, Trim$(gstrDBPasswords.strLocalDBPasswordString))
                lvarErrorStage = 130
                Set gdatCentralDatabase = OpenDatabase(.strCentralDBFile, _
                    dbDriverComplete, False, Trim$(gstrDBPasswords.strCentralDBPasswordString))
            End If
        End If
    End With
    
Exit Sub
ErrorHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "InitDB", "", True, lvarErrorStage)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case gconIntErrHandEndProgram
        Unload frmButtons
        Unload frmSplash
        gintForceAppClose = fcCompleteClose
        Unload mdiMain
    Case Else
        Resume Next
    End Select
    
End Sub
Sub ConcurrencyTest()
Dim lstrMachName As String
Dim lstrPhaseString As String

    lstrMachName = CurrentMachineName
    
    gdatCentralDatabase.Execute mstrRefreshQuery
    
    
    lstrPhaseString = LockingPhaseGen(False)
    
    If CheckUserPhaseExists(Trim$(gstrGenSysInfo.strUserName)) Then
        'if current user has  none blank phase
        If gstrSystemRoute = srCompanyDebugRoute Or gstrSystemRoute = srCompanyRoute Then
            MsgBox "Warning! User name " & Trim$(UCase$(gstrGenSysInfo.strUserName)) & " is currently in use within the central database!" & vbCrLf & vbCrLf & _
                "Please ensure you are not logged in elsewhere or have it open in MS Access." & vbCrLf & _
                "Alternatively, this user name may be being used by another user!" & vbCrLf & _
                vbCrLf & "Access Denied! " & vbCrLf & vbCrLf & _
                "Access Reinstated - The feature has been temporarily disabled for My Company! ", , gconstrTitlPrefix & "Init DB"
                
            gdatCentralDatabase.Execute "UPDATE Users SET Users.Phase = '" & lstrPhaseString & _
                "' WHERE ((ucase(Users.UserID)='" & UCase$(Trim$(gstrGenSysInfo.strUserName)) & "'));"
        Else
            MsgBox "Warning! User name " & Trim$(UCase$(gstrGenSysInfo.strUserName)) & " is currently in use!" & vbCrLf & vbCrLf & "Access Denied!", , gconstrTitlPrefix & "Init DB"
            
            Busy True
            gdatCentralDatabase.Close
            gdatLocalDatabase.Close
            Set gdatLocalDatabase = Nothing
            Set gdatCentralDatabase = Nothing
            UpdateLoader
            Busy False
            Unhook
            End
        End If
            
    Else
        If gbooSQLServerInUse = False Then
            gdatCentralDatabase.Execute "UPDATE Users SET Users.Phase = '" & lstrPhaseString & _
                "' WHERE ((ucase(Users.UserID)='" & UCase$(Trim$(gstrGenSysInfo.strUserName)) & "'));"
        Else
            
            gdatCentralDatabase.Execute "UPDATE Users SET Users.Phase = '" & lstrPhaseString & _
                "' WHERE ((UPPER(Users.UserID)='" & UCase$(Trim$(gstrGenSysInfo.strUserName)) & "'));"
        End If
    End If
    
End Sub
Function RefreshUserPhase(pstrDatabasename As String) As String
Dim lintArrInc As Integer
Dim lstrUsers() As String
Dim lstrSQL As String

    ListLoggedUsers pstrDatabasename, lstrUsers()
    
    lstrSQL = ""
    
    If lstrUsers(0) <> "" Then
    
        For lintArrInc = 0 To UBound(lstrUsers)
            If lstrSQL <> "" Then lstrSQL = lstrSQL & " and "
            lstrSQL = lstrSQL & "Right(Phase,len('" & Trim$(lstrUsers(lintArrInc)) & "')) <> '" & Trim$(lstrUsers(lintArrInc)) & "' "
        Next lintArrInc
        RefreshUserPhase = "Update Users set Phase = ' ' where " & lstrSQL & ";"
    Else
        RefreshUserPhase = "Update Users set Phase = ' ';"
    End If
    
End Function
Function ListLoggedUsers(pstrDatabase As String, pstrList() As String) As String
ReDim msString(1) As String
Dim miLoop As Integer
Dim lintRetVal As Integer
Dim lstrUsers As String
Dim lstrDBLockingFile As String
Dim lintCounter As Integer

    ReDim pstrList(0)

    lstrDBLockingFile = Left(pstrDatabase, Len(pstrDatabase) - 3) & "ldb"
    
    If Dir(lstrDBLockingFile) <> "" Then
        lintRetVal = LDBUser_GetUsers(msString, lstrDBLockingFile, &H2)
        
        For miLoop = LBound(msString) To UBound(msString)
            If Len(msString(miLoop)) = 0 Then
                Exit For
            End If
            If lintCounter = 0 Then
                pstrList(0) = Trim$(msString(miLoop))
            Else
                ReDim Preserve pstrList(UBound(pstrList) + 1)
                pstrList(UBound(pstrList)) = Trim$(msString(miLoop))
            End If
            lstrUsers = lstrUsers & vbTab & msString(miLoop) & vbCrLf
            lintCounter = lintCounter + 1
        Next miLoop
    End If
    
    ListLoggedUsers = lstrUsers
    
End Function
Function CheckUserPhaseExists(pstrUserName As String) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
    
    On Error GoTo ErrHandler
    
    If gbooSQLServerInUse = False Then
        lstrSQL = "SELECT Users.UserID, Users.Phase From Users " & _
            "WHERE (((UCASE(Users.UserID))='" & UCase$(Trim$(pstrUserName)) & "'));"
    Else
        
        lstrSQL = "SELECT Users.UserID, Users.Phase From Users " & _
            "WHERE (((UPPER(Users.UserID))='" & UCase(Trim$(pstrUserName)) & "'));"
    End If
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        Do Until .EOF
            llngRecCount = llngRecCount + 1
            If IsNull(.Fields("Phase")) Or Trim$(.Fields("Phase") & "") = "" Then
                CheckUserPhaseExists = False
            Else
                CheckUserPhaseExists = True
            End If
            .MoveNext
        Loop
    End With
    
    If llngRecCount = 0 Then
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckUserPhaseExists", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Function CountLoggedPhaseUsers() As Integer
Dim lsnaLists As Recordset
Dim lstrSQL As String
    
    On Error GoTo ErrHandler
    
    CountLoggedPhaseUsers = 0
    
    lstrSQL = "SELECT Sum(1) AS [Counter] From Users WHERE (((Users.Phase)<>''));"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            CountLoggedPhaseUsers = Val(.Fields("Counter") & "")
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CountLoggedPhaseUsers", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select

End Function
Function CountUserAccounts() As Integer
Dim lsnaLists As Recordset
Dim lstrSQL As String
    
    On Error GoTo ErrHandler
    
    CountUserAccounts = 0
    
    lstrSQL = "SELECT Count(Users.UserID) AS Counter FROM Users;"
        
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            CountUserAccounts = Val(.Fields("Counter") & "")
        End If
    End With
        
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CountUserAccounts", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function
Function CheckKeyForConcurrent() As Boolean
Dim lstrRefreshQueryFunc As String
Dim lintUserNumAllowance As Integer
Dim lintUserNum As Integer

    CheckKeyForConcurrent = True
    
    CheckStaticCipher
    
    With gstrReferenceInfo
        gstrKey.strUnlockKey = gstrStatic.strUnlockCode
        
        If Decode(.strCompanyName, .strCompanyTelephone, .strCompanyContact) <> "21" Then
            CheckKeyForConcurrent = False
        End If
        
        lintUserNumAllowance = gstrKey.strUserNum
        gstrKey.strUserNum = ""
        
        If InStr(UCase(Command$), "/TEST") > 0 Then
            lstrRefreshQueryFunc = RefreshUserPhase(gstrStatic.strCentralTestingDBFile)
        Else
            lstrRefreshQueryFunc = RefreshUserPhase(gstrStatic.strCentralDBFile)
        End If
        
        gdatCentralDatabase.Execute lstrRefreshQueryFunc
    
        lintUserNum = CountLoggedPhaseUsers
    
        If lintUserNum > lintUserNumAllowance Then
            CheckKeyForConcurrent = False
        End If
        
        ClearDecode
        
        CheckStaticCipher
        gstrStatic.strUnlockCode = ""
    End With
    
End Function
Function CheckKeyForEverything() As Integer
Dim lstrRefreshQueryFunc As String
Dim lintUserNumAllowance As Integer
Dim lintUserNum As Integer
Dim lintUserAcctsNum As Integer
Dim ldatCoverExpire As Date
'Check number of logins against user allowance in key
'Should always read Static file freshly so, whats in memory can't be interfered with.

    CheckKeyForEverything = 0
    
    CheckStaticCipher
    
    With gstrReferenceInfo
        gstrKey.strUnlockKey = gstrStatic.strUnlockCode
        
        If Decode(.strCompanyName, .strCompanyTelephone, .strCompanyContact) <> "21" Then
            CheckKeyForEverything = 99
            GoTo NormalExit
            Exit Function
        End If
        
        lintUserNumAllowance = gstrKey.strUserNum
        gstrKey.strUserNum = ""
        
        If InStr(UCase(Command$), "/TEST") > 0 Then
            lstrRefreshQueryFunc = RefreshUserPhase(gstrStatic.strCentralTestingDBFile)
        Else
            lstrRefreshQueryFunc = RefreshUserPhase(gstrStatic.strCentralDBFile)
        End If
        
        gdatCentralDatabase.Execute lstrRefreshQueryFunc
    
        lintUserNum = CountLoggedPhaseUsers
        If lintUserNum > lintUserNumAllowance Then
            CheckKeyForEverything = 98
            GoTo NormalExit
            Exit Function
        End If
        
        lintUserAcctsNum = CountUserAccounts
        If lintUserAcctsNum > lintUserNumAllowance And lintUserNumAllowance < 30 Then
            CheckKeyForEverything = 97
            GoTo NormalExit
            Exit Function
        End If
                
        ldatCoverExpire = ConvertJulian(CLng(gstrKey.strCoverDate))
                
        If gstrKey.strCover = "Wazz" Then
            If ldatCoverExpire < Date Then
                CheckKeyForEverything = 95
                GoTo NormalExit
                Exit Function
            ElseIf ldatCoverExpire < DateAdd("d", 7, Date) Then
                gstrTempKeyFail = ldatCoverExpire
            End If
        End If
        
        If ldatCoverExpire < Date Then
            CheckKeyForEverything = 96
            GoTo NormalExit
            Exit Function
        End If
        

                
NormalExit:
        ClearDecode
        CheckStaticCipher
        gstrStatic.strUnlockCode = ""
    End With
    
End Function

Function CheckKeyForUserAccts() As Boolean
Dim lintUserNumAllowance As Integer
Dim lintUserAcctsNum As Integer
'Check number of records in user table against allowance in key
'Should always read Static file freshly so, whats in memory can't be interfered with.

    CheckKeyForUserAccts = True
    
    CheckStaticCipher
    
    With gstrReferenceInfo
        gstrKey.strUnlockKey = gstrStatic.strUnlockCode
        
        If Decode(.strCompanyName, .strCompanyTelephone, .strCompanyContact) <> "21" Then
            CheckKeyForUserAccts = False
        End If
        
        lintUserNumAllowance = gstrKey.strUserNum
        gstrKey.strUserNum = ""
            
        lintUserAcctsNum = CountUserAccounts
    
        If lintUserAcctsNum > lintUserNumAllowance And lintUserNumAllowance < 30 Then
            CheckKeyForUserAccts = False
        End If
        
        ClearDecode
        
        CheckStaticCipher
        gstrStatic.strUnlockCode = ""
    End With
    
End Function
Function CheckKeyForMatches() As Boolean
'Check key matches entry in static, registry and file
'Should always read Static file freshly so, whats in memory can't be interfered with.
End Function
Function CheckKeyForCoverExpire() As Date
'Check key expiry of cover
'Should always read Static file freshly so, whats in memory can't be interfered with.
    CheckKeyForCoverExpire = CDate(0)
    
    CheckStaticCipher
    
    With gstrReferenceInfo
        gstrKey.strUnlockKey = gstrStatic.strUnlockCode
        
        If Decode(.strCompanyName, .strCompanyTelephone, .strCompanyContact) <> "21" Then
            CheckKeyForCoverExpire = CDate(0)
            Exit Function
        End If
        
        CheckKeyForCoverExpire = ConvertJulian(CLng(gstrKey.strCoverDate))
        
        ClearDecode
        
        CheckStaticCipher
        gstrStatic.strUnlockCode = ""
    End With
End Function

Function CheckAcctInUseAvail(plngCustomerNum As Long) As Boolean
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim lstrName As String

    CheckAcctInUseAvail = True

    If InStr(UCase(Command$), "/TEST") > 0 Then
        mstrRefreshQuery = RefreshUserPhase(gstrStatic.strCentralTestingDBFile)
    Else
        mstrRefreshQuery = RefreshUserPhase(gstrStatic.strCentralDBFile)
    End If
    
    'Now you know EXACTLY whose logged in
    gdatCentralDatabase.Execute mstrRefreshQuery
        
    lstrSQL = "SELECT * FROM CustAccounts INNER JOIN Users ON CustAccounts.AcctInUseByFlag" & _
        "= Users.Phase WHERE (((CustAccounts.CustNum)=" & plngCustomerNum & _
        ") AND ((Trim([AcctInUseByFlag]))<>'' And (Trim([AcctInUseByFlag])) Is Not Null " & _
        "And (Trim([AcctInUseByFlag]))=[Phase]) AND ((Users.Phase)<>'' And (Users.Phase) Is " & _
        "Not Null And (Users.Phase)=[AcctInUseByFlag]));"
    
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With lsnaLists
        If Not .EOF Then
            CheckAcctInUseAvail = False
            lstrName = Trim$(Trim$(.Fields("Salutation") & "") & " " & _
                Trim$(.Fields("Initials") & "") & " " & Trim$(.Fields("Surname") & ""))
                
            MsgBox "The Account for Customer " & Chr(34) & lstrName & Chr(34) & _
                " is in use by " & .Fields("UserID") & _
                " (" & .Fields("UserName") & ")" & vbCrLf & vbCrLf & _
                "Please Wait until this user has finished with this account and try again later!", _
                vbExclamation, gconstrTitlPrefix & "Account Selection"

        End If
    End With
    
    lsnaLists.Close

Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "CheckAcctInUseAvail", "CENTRAL")
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function
Function LockingPhaseGen(pbooWithTimeStamp As Boolean) As String
Dim lstrMachName As String

    lstrMachName = CurrentMachineName
    
    
    If pbooWithTimeStamp = False Then
        LockingPhaseGen = UCase$(Trim$(gstrGenSysInfo.strUserName)) & "@" & _
            Format(gdatSystemStartTime, "DDMMYYHHMMSS") & "#" & lstrMachName
    Else
        'machine name not record in record locking
        LockingPhaseGen = UCase$(Trim$(gstrGenSysInfo.strUserName)) & "@" & _
            Format(gdatSystemStartTime, "DDMMYYHHMMSS") & "&" & Format(Now(), "DDMMYYHHMMSS")
    
    End If
End Function
Function ReturnPhaseItem(pstrPhaseString As String, pstrItem As ReturnPhaseItems) As String
    
    MsgBox "function ReturnPhaseItem need finishing"
    
    Select Case pstrItem
    Case retPhaseItemUser
    Case retPhaseItemSysStart
    Case retPhaseItemMach
    End Select

End Function
Sub UpdateLoader()
Dim lstrSourcePath As String
Dim lstrDestinationPath As String
Dim lstrSourceFile As String
Dim lstrDestinationFile As String
Dim lbooCopyDone As Boolean
Dim lretval As Variant
Dim llngHwnd As Long
Dim fSuccess As Boolean
Dim lstrRepProg As String
Dim lstrRepParam As String

    ShowStatus 82
    
    If DebugVersion Then
        Exit Sub
    End If
    
    lstrSourcePath = gstrStatic.strServerPath
    lstrDestinationPath = AppPath
    
    lstrSourceFile = lstrSourcePath & "Loader.exe"
    lstrDestinationFile = lstrDestinationPath & "Loader.exe"
    
    lbooCopyDone = FileCopyIfNewer(lstrSourceFile, lstrDestinationFile)
    
    'Also copy Static.Ldr
    lbooCopyDone = FileCopyIfNewer(lstrSourcePath & gconstrStaticLdr, lstrDestinationPath & gconstrStaticLdr)
    
    'Get Latest Clean and Run, thus adding shortcut to startup
    lstrSourceFile = lstrSourcePath & "Minder.exe"
    lstrDestinationFile = lstrDestinationPath & "Minder.exe"
    lbooCopyDone = FileCopyIfNewer(lstrSourceFile, lstrDestinationFile)
    
    fSuccess = OSfCreateShellLink("StartUp", "Minder", _
        lstrDestinationPath & "minder.exe", _
        "", True, "$(Programs)")

    If fSuccess = True Then
        On Error Resume Next
        Kill "C:\WINDOWS\Start Menu\Programs\StartUp\clean.lnk"
    End If
    
    On Error Resume Next
    RunNWait lstrDestinationFile

End Sub
Sub UserLicenceCheck()

    Select Case 0 'DEV NOTE 2019 - removed licensing CheckKeyForEverything
    Case 0
        'OK
    Case 99
        'There is a problem with your user license!
        frmMakeIt.Show vbModal
        Unhook
        End
    Case 97, 98 '98
        MsgBox "You have exceeded the limit of your user license!", , gconstrTitlPrefix & "Startup"
        frmMakeIt.Show vbModal
        Unhook
        End
    Case 96
        gdatCoverDate = CheckKeyForCoverExpire
    Case 95
        MsgBox "Your Temporary License has now expired!" & vbCrLf & vbCrLf & _
            "Temporary licenses are normally issued in certain circumstances at our discretion." & vbCrLf & _
            "Typically, this allows clients to acquire the software quickly, while they arrange " & vbCrLf & _
            "finance through their internal company structure." & vbCrLf & vbCrLf & _
            "Please Contact Mindwarp Consultancy Ltd.", , gconstrTitlPrefix & "Startup"
        frmMakeIt.Show vbModal
        Unhook
        End
    Case Else
        If DebugVersion Then
            MsgBox "User license limit ok!"
        End If
    End Select
    
End Sub
