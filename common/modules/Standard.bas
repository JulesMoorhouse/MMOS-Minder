Attribute VB_Name = "modStandard"
Option Explicit

'Moved from modReports
Type GeneralTrackNUpdate
    lngOrderNum As Long
End Type
Global glngTrackNUpdate() As GeneralTrackNUpdate

Global Const gstrOurCompany = "Mindwarp Consultancy Ltd"
Global Const gstrOurAdd1 = "My Street"
Global Const gstrOurAdd2 = "My Town"
Global Const gstrOurAdd3 = "My County"
Global Const gstrOurAdd4 = "AB 9ZZ"
Global Const gstrOurTel = "(019) 285077 (Answering Service)"
Global Const gstrOurCompanyWeb = "https://www.JulesMoorhouse.com"

Global Const gstrOurContactWeb = "https://www.JulesMoorhouse.com/contactme"

Global gstrIniAppName As String

Type ReferenceInfo
    strCompanyName              As String
    strCompanyAddLine1          As String
    strCompanyAddLine2          As String
    strCompanyAddLine3          As String
    strCompanyAddLine4          As String
    strCompanyAddLine5          As String
    strCreditCardClaimsHead1A   As String
    strCreditCardClaimsHead1B   As String
    strCreditCardClaimsHead2A   As String
    strVATRate175               As String
    strPostageCost              As String
    strPostageWaiveratio        As String
    strDenomination             As String * 1
    
    strCompanyContact           As String
    strCompanyTelephone         As String
    booDonationAvail            As Boolean
    booStockThreashold          As Boolean
End Type
Global gstrReferenceInfo As ReferenceInfo

Type GeneralSystemInfo
    strUserName                 As String * 20
    strUserFullName             As String * 30
    strUserPassword             As String * 20
    lngUserLevel                As Long
    strUserNotes                As String * 255
End Type
Global gstrGenSysInfo           As GeneralSystemInfo

Global Const gstrWorkAround = "This message has been shown to make you aware that this function is not 100% finished." & vbCrLf & _
    "The following details will aid you in a workaround." & vbCrLf & vbCrLf

Global Const gstrComingSoon = "This feature is not yet available." & vbCrLf & vbCrLf & _
    "We will undoubtedly contact you when this feature becomes available." & vbCrLf & _
    "Alternatively, please check our web site for updates and news on new features" & vbCrLf & _
    "http://www.A-Website-to-be-made-available-soon.com" & vbCrLf & vbCrLf & _
    "Mindwarp Consultancy Ltd"

Type ListVars
    strListName                 As String * 50
    strListCode                 As String * 10
    strDescription              As String * 50
    strUserDef1                 As String * 50
    strUserDef2                 As String * 50
End Type

Global gdatCoverDate As Date

'Button Bar route
Global mstrRoute As String
Global gstrButtonRoute As String
Global gstrCurrentLoadedForm As Form

Const gconScreenIni = "d:\desktopnt\Mscreen.ini"
Global Const gconlongButtonTop = 925

Global Const gconstrCSVImport = "CSV import (*.csv)"
Global Const gconstrTabbedImport = "Tabbed Delimoted import (*.tab)"

Public Enum SystemRoutes
    srCompanyRoute = 100
    srCompanyDebugRoute = 300
    srStandardRoute = 200
End Enum
Global gstrSystemRoute As SystemRoutes

Global gbooSQLServerInUse As Boolean
Type DBPasswords
    strLocalDBPasswordString    As String
    strCentralDBPasswordString  As String
End Type
Global gstrDBPasswords As DBPasswords

Global gconstrTitlPrefix As String
Global gconstrProductFullName   As String
Global gconstrProductShortName As String

Type ProgramStaticInfo
    strProgram                  As String
    strParam                    As String
    strDesc                     As String
End Type
Type StaticInfo
    strLocalDBFile              As String
    strLocalTestingDBFile       As String
    strCentralDBFile            As String
    strCentralTestingDBFile     As String
    
    strReportsDBFile            As String
    strReportsTestingDBFile     As String

    strServerPath               As String
    'strAppPath                  As String
    strServerTestNewPath        As String
    strSupportPath              As String
    strSupportTestPath          As String
    
    strPFElecFile               As String
    strVerLogBStatus            As String
    
    strPrograms(3)              As ProgramStaticInfo
    
    'Derived vars, not included in Static
    strShortLocalDBFile              As String
    strShortLocalTestingDBFile       As String
    strShortCentralDBFile            As String
    strShortCentralTestingDBFile     As String
    strShortReportsDBFile            As String
    strShortReportsTestingDBFile     As String
    
    'A variable which will hold the true live server path
    strTrueLiveServerPath           As String
    strUnlockCode                   As String
End Type
Global gstrStatic As StaticInfo
Global Const gconstrStaticIni = "Static.ini"
Global Const gconstrStaticLdr = "Static.ldr"
Sub SetSystemNames()

    gconstrTitlPrefix = "MMOS - "
    gconstrProductFullName = "Mindwarp Mail Order System"
    gconstrProductShortName = "MMOS"
    gstrIniAppName = "Mindwarp Mail Order System"

    Select Case gstrSystemRoute
    Case srCompanyRoute, srCompanyDebugRoute
        If gstrSystemRoute = srCompanyDebugRoute Then
            If InStr(UCase(Command$), "/TEST") > 0 Then
                'DEVNOTE: 2019 - removed
                'gstrDBPasswords.strCentralDBPasswordString = ";Uid=Admin;pwd=password123" 'TODO: Change for production
                'gstrDBPasswords.strLocalDBPasswordString = ";Uid=Admin;pwd=password123"
            Else
                gstrSystemRoute = srCompanyRoute
            End If
        End If
    Case srStandardRoute
        If UCase$(App.ProductName) <> "LITE" Then
            gstrIniAppName = "Mindwarp Mail Order System"
        Else
            gstrIniAppName = "MMOS Lite"
        End If
        'DEVNOTE: 2019 - removed
        'gstrDBPasswords.strCentralDBPasswordString = ";Uid=Admin;pwd=password123"
        'gstrDBPasswords.strLocalDBPasswordString = ";Uid=Admin;pwd=password123"
        
    End Select

End Sub
Function AppPath()

    Dim Ret As String
    Ret = App.Path
    If Right(Ret, 1) <> "\" Then
        Ret = Ret & "\"
    End If

    AppPath = Ret

End Function
Sub WriteBuffer(pstring As String)
    
    With gstrStatic
        .strLocalDBFile = ReturnNthStr(pstring, 1, Chr(182))
        .strLocalTestingDBFile = ReturnNthStr(pstring, 2, Chr(182))
        .strCentralDBFile = ReturnNthStr(pstring, 3, Chr(182))
        .strCentralTestingDBFile = ReturnNthStr(pstring, 4, Chr(182))
        
        .strReportsDBFile = ReturnNthStr(pstring, 6, Chr(182))
        .strReportsTestingDBFile = ReturnNthStr(pstring, 7, Chr(182))
    
        .strPrograms(3).strProgram = ReturnNthStr(pstring, 8, Chr(182))
        .strPrograms(3).strParam = ReturnNthStr(pstring, 9, Chr(182))
        .strPrograms(3).strDesc = ReturnNthStr(pstring, 10, Chr(182))
        
        .strServerPath = ReturnNthStr(pstring, 21, Chr(182))
        .strServerTestNewPath = ReturnNthStr(pstring, 22, Chr(182))
        .strSupportPath = ReturnNthStr(pstring, 23, Chr(182))
        .strSupportTestPath = ReturnNthStr(pstring, 24, Chr(182))
        
        .strPFElecFile = ReturnNthStr(pstring, 25, Chr(182))
        .strVerLogBStatus = ReturnNthStr(pstring, 26, Chr(182))
        
        .strPrograms(0).strProgram = ReturnNthStr(pstring, 27, Chr(182))
        .strPrograms(0).strParam = ReturnNthStr(pstring, 28, Chr(182))
        .strPrograms(0).strDesc = ReturnNthStr(pstring, 29, Chr(182))
        
        .strPrograms(1).strProgram = ReturnNthStr(pstring, 30, Chr(182))
        .strPrograms(1).strParam = ReturnNthStr(pstring, 31, Chr(182))
        .strPrograms(1).strDesc = ReturnNthStr(pstring, 32, Chr(182))
        
        .strPrograms(2).strProgram = ReturnNthStr(pstring, 33, Chr(182))
        .strPrograms(2).strParam = ReturnNthStr(pstring, 34, Chr(182))
        .strPrograms(2).strDesc = ReturnNthStr(pstring, 35, Chr(182))
        
        .strUnlockCode = ReturnNthStr(pstring, 37, Chr(182))
    End With
   
End Sub

Function ReadBuffer() As String
Const lstrIntlyBlank = "Blank"

    With gstrStatic
        ReadBuffer = .strLocalDBFile & Chr(182) & .strLocalTestingDBFile & _
        Chr(182) & .strCentralDBFile & Chr(182) & .strCentralTestingDBFile & _
        Chr(182) & lstrIntlyBlank & Chr(182) & .strReportsDBFile & _
        Chr(182) & .strReportsTestingDBFile & _
        Chr(182) & .strPrograms(3).strProgram & Chr(182) & .strPrograms(3).strParam & Chr(182) & .strPrograms(3).strDesc & _
        Chr(182) & lstrIntlyBlank & Chr(182) & lstrIntlyBlank & _
        Chr(182) & lstrIntlyBlank & Chr(182) & lstrIntlyBlank & _
        Chr(182) & lstrIntlyBlank & Chr(182) & lstrIntlyBlank & _
        Chr(182) & lstrIntlyBlank & Chr(182) & lstrIntlyBlank & _
        Chr(182) & lstrIntlyBlank & Chr(182) & lstrIntlyBlank & _
        Chr(182) & .strServerPath & _
        Chr(182) & .strServerTestNewPath & Chr(182) & .strSupportPath & _
        Chr(182) & .strSupportTestPath & Chr(182) & .strPFElecFile & _
        Chr(182) & .strVerLogBStatus & _
        Chr(182) & .strPrograms(0).strProgram & Chr(182) & .strPrograms(0).strParam & Chr(182) & .strPrograms(0).strDesc & _
        Chr(182) & .strPrograms(1).strProgram & Chr(182) & .strPrograms(1).strParam & Chr(182) & .strPrograms(1).strDesc & _
        Chr(182) & .strPrograms(2).strProgram & Chr(182) & .strPrograms(2).strParam & Chr(182) & .strPrograms(2).strDesc & Chr(182) & .strUnlockCode & Chr(182)
        
   End With

End Function
Function ReturnNthStr(pstrString As String, pintPosition As Integer, pstrChar As String) As String
Dim lintArrInc As Integer
Dim llngPos As Long
Dim llngLastPos As Long
Dim lintErrPos As Integer

    lintErrPos = 0
    On Error GoTo EndFunc
    lintErrPos = 1
    Do Until lintArrInc = pintPosition
        lintErrPos = 2
        lintArrInc = lintArrInc + 1
        lintErrPos = 3
        llngPos = InStr(llngLastPos + 1, pstrString, pstrChar)
        lintErrPos = 4
        ReturnNthStr = Mid(pstrString, llngLastPos + 1, (llngPos - llngLastPos) - 1)
        lintErrPos = 5
        llngLastPos = llngPos
        lintErrPos = 6
    Loop
    lintErrPos = 7
    Exit Function
    
EndFunc:
    Select Case Err.Number
    Case 5
        If lintErrPos = 4 Then
            ReturnNthStr = ""
            Exit Function
        End If
    End Select
    Exit Function
End Function
Sub XCheckStatic()
Dim lstrVAT As String
Dim lstrDenom As String
Dim lstrPostage As String
Dim lstrPOWaiver As String
    
    With gstrStatic
        .strServerPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "ServerPath")
        .strServerTestNewPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SrvTestPth")
        .strSupportPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SuppPath")
        .strSupportTestPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SupTestPth")
                
        .strTrueLiveServerPath = .strServerPath
        If InStr(UCase(Command$), "/TEST") > 0 Then
            .strServerPath = .strServerTestNewPath
        End If
    
        .strLocalDBFile = AppPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Local")
        .strLocalTestingDBFile = AppPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "LocalTest")
        .strCentralDBFile = .strServerPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Central")
        .strCentralTestingDBFile = .strServerPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "CentraTest")
        
        .strReportsDBFile = AppPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Reps")
        .strReportsTestingDBFile = AppPath & GetPrivateINI(AppPath & gconstrStaticIni, "DB", "RepsTest")
    
        'Used in Loader program
        .strShortLocalDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Local")
        .strShortLocalTestingDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "LocalTest")
        .strShortCentralDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Central")
        .strShortCentralTestingDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "CentraTest")
        
        .strShortReportsDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Reps")
        .strShortReportsTestingDBFile = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "RepsTest")
    
        .strVerLogBStatus = Trim$(GetPrivateINI(AppPath & gconstrStaticIni, "Verbose Logging", "BSTAT"))
        .strPFElecFile = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "PFEFile")
    
   End With
    
End Sub

Sub DebugFormControlSizes(pobjForm As Form)
Dim lintArrInc As Integer
Dim lstrIndex As String

    If pobjForm.Height = 8445 And pobjForm.Width = 10605 Then
    
        On Error Resume Next
        For lintArrInc = 0 To pobjForm.Controls.Count - 1   ' Use the Controls collection
            If Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "tim" And _
                  Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "tab" And _
                  Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "ctl" Then
                With pobjForm.Controls(lintArrInc)
                    
                    lstrIndex = ""
                    If .Index <> "" Then
                        lstrIndex = "(" & .Index & ")"
                    End If
                   
                    SetPrivateINI gconScreenIni, pobjForm.Name, .Name & lstrIndex, _
                    .Left & "," & .Top & "," & .Width & "," & .Height

                End With
            End If
        Next
    End If
    
End Sub
Sub CheckStaticCipher(Optional pstrSpecificStaticLdr As String = "")
Dim lstrVAT As String
Dim lstrDenom As String
Dim lstrPostage As String
Dim lstrPOWaiver As String
 
    Dim lstrStaticLdr As String: lstrStaticLdr = Trim$(App.Path) & "\" & gconstrStaticLdr
    If pstrSpecificStaticLdr <> "" Then
        lstrStaticLdr = pstrSpecificStaticLdr
    End If
     'to be used in the future
    Decrypt lstrStaticLdr, gconEncryptStatic
    
    With gstrStatic
        .strTrueLiveServerPath = .strServerPath
    
        If InStr(UCase(Command$), "/TEST") > 0 Then
            .strServerPath = .strServerTestNewPath
        End If
    
        'Used in Loader program
        If .strShortLocalDBFile = "" Then
            .strShortLocalDBFile = .strLocalDBFile
            .strShortLocalTestingDBFile = .strLocalTestingDBFile
            .strShortCentralDBFile = .strCentralDBFile
            .strShortCentralTestingDBFile = .strCentralTestingDBFile
            .strShortReportsDBFile = .strReportsDBFile
            .strShortReportsTestingDBFile = .strReportsTestingDBFile
        End If
        
        If (InStr(1, .strLocalDBFile, AppPath) = 0) Then
            .strLocalDBFile = AppPath & .strLocalDBFile
        End If
        
        If (InStr(1, .strLocalTestingDBFile, AppPath) = 0) Then
            .strLocalTestingDBFile = AppPath & .strLocalTestingDBFile
        End If
        
        If (InStr(1, .strCentralDBFile, AppPath) = 0) Then
            .strCentralDBFile = .strServerPath & .strCentralDBFile
        End If
        
        If (InStr(1, .strCentralTestingDBFile, AppPath) = 0) Then
            .strCentralTestingDBFile = .strServerPath & .strCentralTestingDBFile
        End If
        
        If (InStr(1, .strReportsDBFile, AppPath) = 0) Then
            .strReportsDBFile = AppPath & .strReportsDBFile
        End If
        
        If (InStr(1, .strReportsTestingDBFile, AppPath) = 0) Then
            .strReportsTestingDBFile = AppPath & .strReportsTestingDBFile
        End If
    
   End With
End Sub

Function OneSpace(pstrString As String) As String

    If Len(Trim$(pstrString)) = 0 Or IsBlank(pstrString) Then
        OneSpace = " "
    Else
        OneSpace = (Trim$(pstrString))
    End If

End Function
Function IsBlank(pstrString As Variant) As Boolean
Dim lstrString As String
Dim llngIndex As Long
Dim lbooBlank As Boolean

    If IsNull(pstrString) Then
        IsBlank = True
        pstrString = ""
        Exit Function
    End If
    
    lstrString = pstrString
    
    If Len(lstrString) = 0 Or Trim$(lstrString) = "" Or Asc(Left$(lstrString, 1) & " ") = 0 Then
        IsBlank = True
        pstrString = ""
        Exit Function
    End If

    Exit Function
    
    lbooBlank = True
    For llngIndex = 1 To Len(lstrString)
        If Mid$(lstrString, llngIndex, 1) <> " " Then
            lbooBlank = False
            Exit For
        End If
    Next
    
    If lbooBlank = True Then
        IsBlank = True
        pstrString = ""
        Exit Function
    End If

End Function
Function ReplaceStr(TextIn, ByVal SearchStr As String, ByVal Replacement As String, ByVal CompMode As Integer)
Dim WorkText As String, Pointer As Integer

    If IsNull(TextIn) Then
        ReplaceStr = Null
    Else
        WorkText = TextIn
        Pointer = InStr(1, WorkText, SearchStr, CompMode)
        Do While Pointer > 0
            WorkText = Left(WorkText, Pointer - 1) & Replacement & _
                Mid(WorkText, Pointer + Len(SearchStr))
            Pointer = InStr(Pointer + Len(Replacement), WorkText, _
                SearchStr, CompMode)
        Loop
        ReplaceStr = WorkText
    End If
    
End Function
Function ProperCase(pstrText As String) As String

Dim lintIndex As Integer
Dim lstrPrevious As String * 1
Dim lstrCurrent As String * 1
Dim lblnUpshift As Boolean
Dim lstrOriginal As String
    
    lstrOriginal = Trim$(LCase$(pstrText))
    
    lstrPrevious = " "
    For lintIndex = 1 To Len(lstrOriginal)
        lblnUpshift = False

        If Right$(UCase$(lstrPrevious), 1) < "A" Or _
           Right$(UCase$(lstrPrevious), 1) > "Z" Then
            lblnUpshift = True
        End If
        If Right$(lstrPrevious, 1) = "'" Then
            If lintIndex + 1 < Len(lstrOriginal) Then
                If UCase$(Mid$(lstrOriginal, lintIndex + 1, 1)) >= "A" And _
                   UCase$(Mid$(lstrOriginal, lintIndex + 1, 1)) <= "Z" Then
                    lblnUpshift = True
                Else
                    lblnUpshift = False
                End If
            End If
        End If
        If lintIndex > 2 And lblnUpshift = False Then
            If UCase$(Mid$(lstrOriginal, lintIndex - 2, 2)) = "MC" Then
                If lintIndex > 3 Then
                    If UCase$(Mid$(lstrOriginal, lintIndex - 3, 1)) = " " Then
                        lblnUpshift = True
                    End If
                Else
                    lblnUpshift = True
                End If
            End If
        End If
        lstrCurrent = Mid$(lstrOriginal, lintIndex, 1)
        If lblnUpshift Then
            Mid$(lstrOriginal, lintIndex, 1) = UCase$(lstrCurrent)
        End If
        lstrPrevious = lstrCurrent
    Next
    
    ProperCase = lstrOriginal

End Function
Function Pad(pobjForm As Form, plngLength As Long, pstrString As String) As String
'Moved from modSetup for use in Lite version
Dim llngTextWidth As Long
Dim lstrPadding As String
    
    Do While llngTextWidth <= pobjForm.TextWidth(String(plngLength, "p")) / 2
        lstrPadding = lstrPadding & " "
        llngTextWidth = pobjForm.TextWidth(pstrString & lstrPadding)
    Loop
    
    Pad = pstrString & lstrPadding & vbTab
    
End Function
Function CSVNthStr(pstrString As String, pintPosition As Integer) As String
Dim lintArrInc As Integer
Dim llngPos As Long
Dim llngLastPos As Long
Dim lintErrPos As Integer
Dim llngSubComma As Long
Dim llngSubQuote As Long
Dim llngKeptPos As Long
Dim pstrChar As String

    pstrChar = ","
    lintErrPos = 0
    On Error GoTo EndFunc
    lintErrPos = 1
    Do Until lintArrInc = pintPosition
        lintErrPos = 2
        lintArrInc = lintArrInc + 1
        lintErrPos = 3
        llngPos = InStr(llngLastPos + 1, pstrString, pstrChar)
        lintErrPos = 9
        If Mid$(pstrString, llngLastPos + 1, 1) = Chr(34) Then
            lintErrPos = 8
            llngSubComma = InStr(llngPos - 1, pstrString, ",")
            lintErrPos = 10
            llngSubQuote = InStr(llngPos - 1, pstrString, Chr(34))
            If llngSubComma < llngSubQuote Then
                lintErrPos = 4
                llngPos = InStr(llngPos, pstrString, Chr(34)) + 1
            End If
        End If
        lintErrPos = 11
        CSVNthStr = Mid(pstrString, llngLastPos + 1, (llngPos - llngLastPos) - 1)
        lintErrPos = 5
        llngLastPos = llngPos
        lintErrPos = 6
    Loop
    lintErrPos = 7
    Exit Function
    
EndFunc:
    Select Case Err.Number
    Case 5
        If lintErrPos = 4 Then
            CSVNthStr = ""
            Exit Function
        End If
    End Select
    Exit Function
End Function
