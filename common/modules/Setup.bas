Attribute VB_Name = "modSetup"
Option Explicit
Sub TableAttach(pstrTableName As String, pdabDatabase As Database, _
    Optional pstrDatabase As Variant, Optional pstrSourceTableName As Variant)
Dim ltdfMaster As TableDef
Dim lstrDatabase As String

    lstrDatabase = pstrDatabase
    
    If IsMissing(pstrSourceTableName) Then
        pstrSourceTableName = pstrTableName
    End If
    
    Set ltdfMaster = pdabDatabase.CreateTableDef(pstrTableName)
    
    With ltdfMaster
    
        .Connect = ";DATABASE=" & lstrDatabase
        .SourceTableName = pstrSourceTableName
        
        On Error GoTo AttachError
        
        pdabDatabase.TableDefs.Append ltdfMaster
        
    End With
    
NormalEnd:
    Exit Sub
    
AttachError:
    If Err.Number = 3012 Then 'already attached!
        Resume Next
    Else
        MsgBox "Attach Error (" & Err.Number & ") " & Err.Description, vbExclamation, , , gconstrTitlPrefix & "TableAttach"
    End If
    Resume Next

End Sub
Sub TableDetach(pstrTableName As String, pdabDatabase As Database)

    On Error Resume Next
    With pdabDatabase.TableDefs(pstrTableName)
    
        If .SourceTableName = "" Then
            'MsgBox "Can't detach base tables!", vbExclamation
        Else
            pdabDatabase.TableDefs.Delete pstrTableName
        End If
    
    End With

End Sub
Function DeployStaticInfo(Optional pbooConfig As Boolean)
Dim lsnaLists As Recordset
Dim lstrSQL As String
Dim llngRecCount As Long
'Const lconstrStaticIni = "Static.ini"
    
    If IsMissing(pbooConfig) Then
        pbooConfig = False
    End If
    
    If pbooConfig = False Then
        ShowStatus 41
    End If
    
    On Error GoTo ErrHandler
    
    lstrSQL = "SELECT ListsMaster.ListName, ListDetailsMaster.ListCode, " & _
        "ListDetailsMaster.Description FROM ListsMaster INNER JOIN " & _
        "ListDetailsMaster ON ListsMaster.ListNum = ListDetailsMaster.ListNum " & _
        "WHERE (((ListsMaster.SysUse)=True));"
    
            
    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    
    With gstrStatic
    
        llngRecCount = 0
        
        Do Until lsnaLists.EOF
            'SetPrivateINI gstrStatic.strServerPath & gconstrStaticIni, lsnaLists.Fields("Listname"), _
                lsnaLists.Fields("ListCode"), lsnaLists.Fields("Description")
            
            If pbooConfig = False Then
                
               
                'SetPrivateINI gstrStatic.strTrueLiveServerPath & gconstrStaticIni, lsnaLists.Fields("Listname"), _
                    lsnaLists.Fields("ListCode"), lsnaLists.Fields("Description")
            End If
            
            Select Case lsnaLists.Fields("ListName")
            'Case "Amount Levels"
            '    If lsnaLists.Fields("ListCode") = "DENOM" Then .strDenomination = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "POSTAGE" Then .strPostageCost = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "POWAIVER" Then .strPostageWaiveratio = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "VAT" Then .strVATRate175 = lsnaLists.Fields("Description")
            Case "DB"
                If lsnaLists.Fields("ListCode") = "Central" Then .strCentralDBFile = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "CentraTest" Then .strCentralTestingDBFile = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Local" Then .strLocalDBFile = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "LocalTest" Then .strLocalTestingDBFile = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Reps" Then .strReportsDBFile = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "RepsTest" Then .strReportsTestingDBFile = lsnaLists.Fields("Description")
            Case "SysFileInfo"
                If lsnaLists.Fields("ListCode") = "ServerPath" Then .strServerPath = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "SrvTestPth" Then .strServerTestNewPath = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "SuppPath" Then .strSupportPath = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "SupTestPth" Then .strSupportTestPath = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "PFEFile" Then .strPFElecFile = lsnaLists.Fields("Description")
                
            Case "Programs"
                If lsnaLists.Fields("ListCode") = "Prog1" Then .strPrograms(0).strProgram = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog1Desc" Then .strPrograms(0).strDesc = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog1Param" Then .strPrograms(0).strParam = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog2" Then .strPrograms(1).strProgram = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog2Desc" Then .strPrograms(1).strDesc = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog2Param" Then .strPrograms(1).strParam = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog3" Then .strPrograms(2).strProgram = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog3Desc" Then .strPrograms(2).strDesc = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog3Param" Then .strPrograms(2).strParam = lsnaLists.Fields("Description")
               
                If lsnaLists.Fields("ListCode") = "Prog4" Then .strPrograms(3).strProgram = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog4Desc" Then .strPrograms(3).strDesc = lsnaLists.Fields("Description")
                If lsnaLists.Fields("ListCode") = "Prog4Param" Then .strPrograms(3).strParam = lsnaLists.Fields("Description")
            'Case "Company Address"
            '    If lsnaLists.Fields("ListCode") = "CONAME" Then .strCompanyName = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "COADLI1" Then .strCompanyAddLine1 = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "COADLI2" Then .strCompanyAddLine2 = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "COADLI3" Then .strCompanyAddLine3 = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "COADLI4" Then .strCompanyAddLine4 = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "COADLI5" Then .strCompanyAddLine5 = lsnaLists.Fields("Description")
            'Case "Card Serv Header"
            '    If lsnaLists.Fields("ListCode") = "CSHEAD1A" Then .strCreditCardClaimsHead1A = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "CSHEAD1B" Then .strCreditCardClaimsHead1B = lsnaLists.Fields("Description")
            '    If lsnaLists.Fields("ListCode") = "CSHEAD2A" Then .strCreditCardClaimsHead2A = lsnaLists.Fields("Description")
            End Select
            lsnaLists.MoveNext
        Loop
        'pobjList.AddItem ""
    End With
    
    If pbooConfig = False Then
        If llngRecCount <> 0 Then
            MsgBox "Deployment completed OK!", , gconstrTitlPrefix & "Static Deployment"
        End If
    End If
    
    lsnaLists.Close
    
Exit Function
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "DeployStaticInfo", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Function
    Case Else
        Resume Next
    End Select
End Function

Sub UpdateObjects(pintIndex As Integer, ByVal pobjTextBox As TextBox)

    With gstrSystemLists(pintIndex)
        .strDescValue = pobjTextBox
    End With
    
End Sub
Sub UpdateListDetailWithObject(pintIndex As Integer, ByVal pobjTextBox As TextBox)
Dim lstrSQL As String

    On Error GoTo ErrHandler
    
    With gstrSystemLists(pintIndex)
        
        lstrSQL = "UPDATE ListsMaster INNER JOIN ListDetailsMaster ON " & _
            "ListsMaster.ListNum = ListDetailsMaster.ListNum SET " & _
            "ListDetailsMaster.Description = '" & OneSpace(pobjTextBox.Text) & "' WHERE " & _
            "(((ListsMaster.ListName)='" & .strListName & _
            "') AND ((ListDetailsMaster.ListCode)='" & .strListCode & "'));"

        gdatCentralDatabase.Execute lstrSQL
        
    End With
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "UpdateListDetailWithObject", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Function CheckObjects(pintIndex As Integer, ByVal pobjTextBox As TextBox) As Boolean
    
    CheckObjects = False
    
    With gstrSystemLists(pintIndex)
        If Trim$(pobjTextBox) = "" Then
            MsgBox "You must enter a value for each field!", , gconstrTitlPrefix & "Mandatory Field"
            Exit Function
        End If
    End With
    
    CheckObjects = True
    
End Function

Sub PopulateObjects(pintIndex As Integer, ByVal pobjLabel As Label, _
    ByVal pobjTextBox As TextBox, ByVal pobjLabelExample As Object)

    With gstrSystemLists(pintIndex)
        If .strTopic = "" Then
            pobjLabel = Trim$(.strListCode) & " :"
        Else
            pobjLabel = Trim$(.strTopic) & " :"
        End If
        pobjTextBox.MaxLength = 50
        
        'remove after testing
        If .strDescValue = "" Then
            pobjTextBox = .strExampleDesc
        Else
            pobjTextBox = .strDescValue
        End If
        
        pobjLabelExample = "e.g. " & .strExampleDesc
    End With
    
End Sub

Sub ReStackArray(pintIndex As Integer, ByVal pobjCodeList As ListBox, _
    ByVal pobjDescList As Object, ByVal pobjSeqList As ListBox)
Dim lintArrInc As Integer

    With gstrSystemListsMultiFull(pintIndex)
        'Clear Array first
        For lintArrInc = 0 To 10
            .strDetail(lintArrInc).strListCode = ""
            .strDetail(lintArrInc).strDescValue = ""
            .strDetail(lintArrInc).lngSeqNum = 0
        Next lintArrInc
        
        For lintArrInc = 0 To pobjCodeList.ListCount
            'pobjCodeList.ListIndex = lintArrInc
            'pobjDescList.ListIndex = lintArrInc
            'pobjSeqList.ListIndex = lintArrInc
            .strDetail(lintArrInc).strListCode = pobjCodeList.List(lintArrInc)
            .strDetail(lintArrInc).strDescValue = pobjDescList.List(lintArrInc)
            .strDetail(lintArrInc).lngSeqNum = Val(pobjSeqList.List(lintArrInc))
        Next lintArrInc
    End With

End Sub

Sub PopulateMultiObjects(pintIndex As Integer, ByVal pobjCodeList As ListBox, _
    ByVal pobjDescList As Object, ByVal pobjSeqList As ListBox, ByVal pobjLabel As Label)
Dim lintArrInc As Integer

    pobjCodeList.Clear
    pobjDescList.Clear
    pobjSeqList.Clear
    
    With gstrSystemListsMultiFull(pintIndex)
        pobjLabel = .strListName
        pobjLabel.Tag = pintIndex
        For lintArrInc = 0 To 10
            If .strDetail(lintArrInc).strListCode <> "" Then
                pobjCodeList.AddItem .strDetail(lintArrInc).strListCode
                pobjDescList.AddItem .strDetail(lintArrInc).strDescValue
                pobjSeqList.AddItem .strDetail(lintArrInc).lngSeqNum
            End If
        Next lintArrInc
    End With
    
    If pobjCodeList.ListCount > 0 Then
        pobjCodeList.ListIndex = 0
    End If
    
End Sub
Sub FillObjectWithListValue(pintIndex As Integer, ByVal pobjLabel As Label, _
    ByVal pobjTextBox As TextBox)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    On Error GoTo ErrHandler
    
    With gstrSystemLists(pintIndex)
        lstrSQL = "SELECT ListsMaster.ListName, ListDetailsMaster.ListCode, " & _
            "ListDetailsMaster.Description, ListDetailsMaster.UserDef1, " & _
            "ListDetailsMaster.UserDef2, ListDetailsMaster.InUse, " & _
            "ListDetailsMaster.SequenceNum FROM ListsMaster INNER JOIN " & _
            "ListDetailsMaster ON ListsMaster.ListNum = ListDetailsMaster.ListNum " & _
            "Where (((ListsMaster.ListName) = '" & Trim$(.strListName) & _
            "' AND (ListDetailsMaster.ListCode)='" & .strListCode & "'));"
            
        Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
        If Not lsnaLists.EOF Then
        
            pobjTextBox = lsnaLists.Fields("Description")
            pobjLabel = Trim$(.strTopic) & " :"
            pobjTextBox.MaxLength = 50
            pobjTextBox.Visible = True
            pobjLabel.Visible = True
        End If
    End With
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillObjectWithListValue", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Sub FillMultiObjectWithListValues(pintIndex As Integer, ByVal pobjCodeList As ListBox, _
    ByVal pobjDescList As Object, ByVal pobjSeqList As ListBox, ByVal pobjLabel As Label)
Dim lstrSQL As String
Dim lsnaLists As Recordset

    On Error GoTo ErrHandler
    
    pobjCodeList.Clear
    pobjDescList.Clear
    pobjSeqList.Clear
    
    With gstrSystemListsMultiFull(pintIndex)
        lstrSQL = "SELECT ListsMaster.ListName, ListDetailsMaster.ListCode, " & _
            "ListDetailsMaster.Description, ListDetailsMaster.UserDef1, " & _
            "ListDetailsMaster.UserDef2, ListDetailsMaster.InUse, " & _
            "ListDetailsMaster.SequenceNum FROM ListsMaster INNER JOIN " & _
            "ListDetailsMaster ON ListsMaster.ListNum = ListDetailsMaster.ListNum " & _
            "Where (((ListsMaster.ListName) = '" & Trim$(.strListName) & "')) ORDER BY " & _
            "ListDetailsMaster.SequenceNum;"
    End With

    Set lsnaLists = gdatCentralDatabase.OpenRecordset(lstrSQL, dbOpenSnapshot)
    Do Until lsnaLists.EOF
        pobjLabel = lsnaLists.Fields("ListName")
        pobjCodeList.AddItem lsnaLists.Fields("ListCode")
        pobjDescList.AddItem lsnaLists.Fields("Description")
        pobjSeqList.AddItem lsnaLists.Fields("SequenceNum")
        lsnaLists.MoveNext
    Loop
    lsnaLists.Close
    
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "FillMultiObjectWithListValues", "Central", True)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    Case Else
        Resume Next
    End Select
End Sub
Sub DeployLoaderFile()
Dim lintRetVal As Integer
Dim lvarErrorStage

    lintRetVal = MsgBox("This process will update the Server and File location settings used by all users." & _
        vbCrLf & vbCrLf & "The information used to make this change is stored only in the Live database.  Therefore" & _
        vbCrLf & "this process is independent of the testing environment." & vbCrLf & vbCrLf & _
        "If you wish to proceed click YES!", vbCritical + vbYesNo, gconstrTitlPrefix & "System Lists")

    If lintRetVal = vbYes Then
    
        'Used to get unlock code 28/10/01
        Decrypt gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, gconEncryptStatic
        
        gstrUserMode = gconstrLiveMode
        'ModeChange Me, gstrUserMode
        With gstrStatic
            If gstrSystemRoute = srCompanyRoute Then
                lvarErrorStage = 110
'                Set gdatCentralDatabase = OpenDatabase(.strCentralDBFile, , False)
                Set gdatCentralDatabase = OpenDatabase(.strTrueLiveServerPath & .strShortCentralDBFile, , False)
            Else
                lvarErrorStage = 130
 '               Set gdatCentralDatabase = OpenDatabase(.strCentralDBFile, _
                    dbDriverComplete, False, Trim$(gstrDBPasswords.strCentralDBPasswordString))
                Set gdatCentralDatabase = OpenDatabase(.strTrueLiveServerPath & .strShortCentralDBFile, _
                    dbDriverComplete, False, Trim$(gstrDBPasswords.strCentralDBPasswordString))
            End If
        
            
'            If .strUnlockCode = "" And gstrSystemRoute = srCompanyRoute Then
            ''If gstrSystemRoute = srCompanyRoute Then
            ''    .strUnlockCode = "XVRBHH370YRQORUFPWST"
            ''End If
        
            
            ''If gstrSystemRoute = srCompanyDebugRoute Then
            ''    .strUnlockCode = "XVRBHH370YRQORUFPWST"
            ''End If
            
            DeployStaticInfo
            
            
            'Encrypt gstrStatic.strServerPath & gconstrStaticLdr
            Encrypt gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, gconEncryptStatic
            MsgBox "Static has been deployed!" & vbCrLf & vbCrLf & "You are about to logged out of the system!", , gconstrTitlPrefix & "System Lists"
            
            'UpdateLoader
            'Unload Me
           
            Unhook
            End
        End With
    End If
End Sub
Sub Migrate(pobjForm As Form)
Dim llngBrowseForNum As Long
Dim lstrOldServerPath As String
Dim lintRetVal

    llngBrowseForNum = 18

    'Will need password protection
    lintRetVal = MsgBox(gstrWorkAround & "Please Ensure that ALL users are logged out of the system " & _
        "before proceeding", vbYesNo, gconstrTitlPrefix & "Config")
    If lintRetVal <> vbYes Then Exit Sub
    
    If UCase(Dir(Trim$(App.Path) & "\" & gconstrStaticLdr, vbNormal)) = UCase(gconstrStaticLdr) Then
        'get existing loader info, this will be used unless overwritten below
        CheckStaticCipher
        With gstrStatic
            lintRetVal = MsgBox("INSTRUCTIONS:" & vbCrLf & vbCrLf & _
                "Before continuing with this process, first copy all files and subdirectories," & _
                vbCrLf & "From :-" & vbCrLf & vbTab & .strServerPath & vbCrLf & _
                "To :- " & vbCrLf & vbTab & "your new location / path. " & vbCrLf & vbCrLf & _
                "Then press YES (below) and select your new path." & vbCrLf & vbCrLf & _
                "After this process is complete your old path will be requied to allow your existing users to " & vbCrLf & _
                "seamlessly migrate to your new location.  Only when you are sure that all your users are using" & vbCrLf & _
                "the new location, may you withdraw it. ", vbYesNo, gconstrTitlPrefix & "Config")
            If lintRetVal <> vbYes Then Exit Sub
            
            'Make a copy of the Old loader file just in case
            lstrOldServerPath = .strServerPath
            WinCopyDlg pobjForm, lstrOldServerPath & gconstrStaticLdr, lstrOldServerPath & "staticbak.ldr"
            
            'Get a new server path from the user
            .strServerPath = GetNetDir(pobjForm, llngBrowseForNum)
            If .strServerPath = "" Then
                Exit Sub
            End If
            
            'Update other paths using Server path
            .strServerTestNewPath = .strServerPath & "TestNew\"
            .strSupportPath = .strServerPath & "Setup\Support\"
            .strSupportTestPath = .strServerPath & "TestNew\Setup\Support\"
            '.strCentralDBFile = .strServerPath & .strShortCentralDBFile
            '.strCentralTestingDBFile = .strServerPath & .strShortCentralTestingDBFile
                                   
            .strLocalDBFile = .strShortLocalDBFile
            .strLocalTestingDBFile = .strShortLocalTestingDBFile
            .strCentralDBFile = .strShortCentralDBFile
            .strCentralTestingDBFile = .strShortCentralTestingDBFile
            .strReportsDBFile = .strShortReportsDBFile
            .strReportsTestingDBFile = .strShortReportsTestingDBFile
            
            'Write new loader file to new server path
            Encrypt .strServerPath & gconstrStaticLdr, gconEncryptStatic
            
            'Copy the New loader file to the old server, so users will be reidirected
            WinCopyDlg pobjForm, .strServerPath & gconstrStaticLdr, lstrOldServerPath & gconstrStaticLdr
            
            MsgBox "Process Completed", , gconstrTitlPrefix & "Config"
            
            MsgBox gstrWorkAround & "You must now Re-Link all attached tables!", gconstrTitlPrefix & "Config"
        End With
    Else
        MsgBox "Existing Loader file not found!", , gconstrTitlPrefix & "Config"
    End If
        
End Sub
Sub SetupShortcuts(pobjTextBox As TextBox, plngUserLevel As Long, pstrLocalPath As String)
Dim fSuccess As Boolean
Const lconMCLGroup = "Mindwarp Mail Order Suite"


    fSuccess = OSfCreateShellLink("StartUp", "Minder", _
        pstrLocalPath & "minder.exe", _
        "", True, "$(Programs)")
    If fSuccess = False Then
        pobjTextBox.LinkTopic = "ProgMan|Progman"
        pobjTextBox.LinkMode = 2               ' Establish manual link.
        pobjTextBox.LinkExecute "[CreateGroup(Startup)]"
        pobjTextBox.LinkExecute "[AddItem(" & pstrLocalPath & "minder.exe, Minder)]"

    End If
    
    pobjTextBox.LinkTopic = "ProgMan|Progman"
    pobjTextBox.LinkMode = 2               ' Establish manual link.
    pobjTextBox.LinkExecute "[CreateGroup(" & lconMCLGroup & ")]"
    
    fSuccess = OSfCreateShellLink(lconMCLGroup, "Mindwarp Mail Order System", _
        pstrLocalPath & "loader.exe", _
        "/X", True, "$(Programs)")
    If fSuccess = False Then
        pobjTextBox.LinkExecute "[AddItem(" & pstrLocalPath & "loader.exe /X, Mindwarp MOS)]"
    End If
    
    If plngUserLevel >= 30 Then
        fSuccess = OSfCreateShellLink(lconMCLGroup, "Mindwarp MOS Manager", _
            pstrLocalPath & "loader.exe", _
            "/REPORT", True, "$(Programs)")
        If fSuccess = False Then
            pobjTextBox.LinkExecute "[AddItem(" & pstrLocalPath & "loader.exe /REPORT, Mindwarp MOS Manager)]"
        End If
    End If
    
    If plngUserLevel >= 50 Then
        fSuccess = OSfCreateShellLink(lconMCLGroup, "Mindwarp MOS Maintenance", _
            pstrLocalPath & "loader.exe", _
            "/ADMIN", True, "$(Programs)")
        If fSuccess = False Then
            pobjTextBox.LinkExecute "[AddItem(" & pstrLocalPath & "loader.exe /ADMIN, Mindwarp MOS Maintenance)]"
        End If
    End If
    
    If plngUserLevel >= 99 Then
        fSuccess = OSfCreateShellLink(lconMCLGroup, "Mindwarp MOS Configuration", _
            pstrLocalPath & "loader.exe", _
            "/CONFIG", True, "$(Programs)")
        If fSuccess = False Then
            pobjTextBox.LinkExecute "[AddItem(" & pstrLocalPath & "loader.exe /CONFIG, Mindwarp MOS Configuration)]"
        End If
    End If
    
End Sub
Sub SetAttrAll(pstrPath As String, pstrWildcard As String)
Dim lstrFileName As String

    lstrFileName = Dir(pstrPath & "\" & pstrWildcard, vbNormal)
    
    Do While lstrFileName <> ""
        SetAttr pstrPath & "\" & lstrFileName, vbNormal
        lstrFileName = Dir
    Loop
        
End Sub
Sub AddNewUser(pstrUserName As String, pstrFullName As String, plngUserLevel As Long, _
    pstrPassword As String, pstrUserNotes As String, Optional pbooNoStatus As Variant)
Dim lstrSQL As String

    If IsMissing(pbooNoStatus) Then
        pbooNoStatus = False
    End If
    
    If pbooNoStatus = False Then
        ShowStatus 80
    End If
    
    On Error GoTo ErrHandler
    
    pstrPassword = Hash(pstrPassword)
    
    lstrSQL = "INSERT INTO Users ( UserID, UserPassword, Username," & _
        "Userlevel, UserNotes ) select '" & _
        Trim$(pstrUserName) & "', '" & pstrPassword & "', '" & pstrFullName & _
        "', '" & plngUserLevel & "', '" & pstrUserNotes & "';"
    
    gdatCentralDatabase.Execute lstrSQL

Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "AddNewUser", "Central")
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
Function CopyData(from_db As Database, to_db As Database, _
   from_nm As String, to_nm As String) As Integer
Dim dbsource As Database
Dim dbdest As Database
Dim rsSource As Recordset
Dim rsDest As Recordset
Dim lintArrInc As Integer
'Added

    On Error GoTo CopyErr
     
    Set rsSource = from_db.OpenRecordset(from_nm)
    Set rsDest = to_db.OpenRecordset(to_nm, dbOpenTable)
    
    While rsSource.EOF = False
       rsDest.AddNew
       
       For lintArrInc = 0 To rsSource.Fields.Count - 1
          rsDest(lintArrInc) = rsSource(lintArrInc)
       Next
       
       rsDest.Update
       rsSource.MoveNext
    Wend
    
    CopyData = True
    GoTo CopyEnd
   
CopyErr:
    CopyData = False
    Resume CopyEnd

CopyEnd:
End Function

