Attribute VB_Name = "modAPI"
Option Explicit

Global gstrSourcePath As String
Global gbooErrorFound As Boolean

Global gintLastError As String
Global gdatSystemStartTime As Date

Global Const gconstrPleaseLeaveError = "" & _
"Please contact you Technical support office and inform them.  Please leave this error message on the screen"

Global Const gconIntErrHandRetry = 20
Global Const gconIntErrHandExitFunction = 30
Global Const gconIntErrHandEndProgram = 40

Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" _
(ByVal lstrBuffer As String, ByVal llngsize As Long) As Long

Public Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function CloseWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpclassname As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function DestroyWindow Lib "User32" (ByVal hwnd As Long) As Long

Public Declare Function OSfCreateShellLink Lib "VB6STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const MAX_COMPUTERNAME_LENGTH As Long = 15&


Public Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Public Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function CreateProcessA Lib "KERNEL32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long


Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Public Declare Function apiFindWindow Lib "User32" Alias "FindWindowA" (ByVal lpclassname As Any, ByVal lpCaption As Any)
   
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&

Declare Function LDBUser_GetUsers Lib "MSLDBUSR.DLL" (lpszUserBuffer() As String, ByVal lpszFilename As String, ByVal nOptions As Long) As Integer
Declare Function LDBUser_GetError Lib "MSLDBUSR.DLL" (ByVal nErrorNo As Long) As String
Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function apiCopyFile Lib "KERNEL32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'GetNetDir calls Start
Public Const CSIDL_NETWORK As Long = &H12
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
    "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long

Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
                              (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                              pIdl As ITEMIDLIST) As Long
                              
Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As String
   'pszDisplayName As Long
   'lpszTitle As Long
   lpszTitle As String
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type

Private Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type

Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Public Const NOERROR = 0
'GetNetDir Calls End

'Windows Copy Dlg start
Private Const FO_COPY = &H2&
Private Const FO_DELETE = &H3&
Private Const FO_MOVE = &H1&
Private Const FO_RENAME = &H4&
Private Const FOF_ALLOWUNDO = &H40&   'Preserve Undo information.
Private Const FOF_CONFIRMMOUSE = &H2& 'Not currently implemented.
Private Const FOF_CREATEPROGRESSDLG = &H0&
Private Const FOF_FILESONLY = &H80&
Private Const FOF_MULTIDESTFILES = &H1&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FOF_RENAMEONCOLLISION = &H8&
Private Const FOF_SILENT = &H4&
Private Const FOF_SIMPLEPROGRESS = &H100&
Private Const FOF_WANTMAPPINGHANDLE = &H20&

Private Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" _
      (lpFileOp As Any) As Long
      
'---Stop Close Button and close from Control box Start---
'Menu item constants.
Public Const SC_CLOSE       As Long = &HF060&
'SetMenuItemInfo fMask constants.
Public Const MIIM_STATE     As Long = &H1&
Public Const MIIM_ID        As Long = &H2&
'SetMenuItemInfo fState constants.
Public Const MFS_GRAYED     As Long = &H3&
Public Const MFS_CHECKED    As Long = &H8&
'SendMessage constants.
Public Const WM_NCACTIVATE  As Long = &H86
'User-defined Types.
Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type
'Declarations.
Public Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemInfo Lib "User32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "User32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Application-specific constants and variables.
Public Const xSC_CLOSE  As Long = -10
Public Const SwapID     As Long = 1
Public Const ResetID    As Long = 2
Public hMenu  As Long
Public MII    As MENUITEMINFO
'---Stop Close Button and close from Control box End---
'Windows Copy Dlg End

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
    lpRootPathName As String, ByVal lpVolumeNameBuffer As _
    String, ByVal nVolumeNameSize As Integer, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength _
    As Long, lpFileSystemFlags As Long, ByVal _
    lpFileSystemNameBuffer As String, ByVal _
    nFileSystemNameSize As Long) As Long
    
Public Declare Function GetDriveType Lib "KERNEL32" Alias "GetDriveTypeA" (ByVal sDrive As String) As Long

'Hook
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const WM_TCARD = &H52&

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_CLOSE_ALL = &H12

Public Const HH_TP_HELP_WM_HELP = &H11

Public Declare Function HTMLHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hwnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long
    
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public glngPrevWndProc As Long
Global gHW As Long

Public Declare Function DefWindowProc Lib "User32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)

'Dynamic menu
Type dhDoubleWordByWord
    LowWord As Integer
    HighWord As Integer
End Type

Type dhDoubleWordLong
    DoubleWord As Long
End Type

Type dhDoubleWordByByte
    LowWordLowByte As Byte
    LowWordHighByte As Byte
    HighWordLowByte As Byte
    HighWordHighByte As Byte
End Type
Public Const WM_MENUSELECT = &H11F
Public Const WM_COMMAND = &H111
Public Const WM_ACTIVATE As Long = &H6
Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_SEPARATOR = &H800&
Public Const MF_BITMAP = &H4&
Public Const MF_CHECKED = &H8&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_UNCHECKED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_UNGRAYED = &H0
Public Const MF_BYCOMMAND = &H0&

Public Declare Function CreateMenu Lib "User32" () As Long
Public Declare Function CreatePopupMenu Lib "User32" () As Long
Public Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function SetMenu Lib "User32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyMenu Lib "User32" (ByVal hMenu As Long) As Long
Global glngUIRetMenu As Long
'----Dynamic Menu stuff

'App Instance stuff
Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetTopWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Const GW_HWNDNEXT = 2

'reporting and screen
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const HORZRES = 8
Public Const VERTRES = 10

Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const PHYSICALWIDTH = 110
Public Const PHYSICALHEIGHT = 111
Public Const PHYSICALOFFSETX = 112
Public Const PHYSICALOFFSETY = 113

Private Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long

Public Function DriveType(sDrive As String) As String
Dim sDriveName As String
Const DRIVE_TYPE_UNDTERMINED = 0
Const DRIVE_ROOT_NOT_EXIST = 1
Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6

    sDriveName = GetDriveType(sDrive)
    
    Select Case sDriveName
    'Case DRIVE_TYPE_UNDTERMINED
    '    DriveType = "has not been recognized"
    'Case DRIVE_ROOT_NOT_EXIST
    '    DriveType = "specified doesn't exist"
    'Case DRIVE_CDROM
    '    DriveType = "is a CD-ROM drive."
    'Case DRIVE_FIXED
    '    DriveType = "cannot be removed I.E. Hard Disk"
    'Case DRIVE_RAMDISK
    '    DriveType = "is a RAM disk."
    'Case DRIVE_REMOTE
    '    DriveType = "is a remote I.E Network drive."
    'Case DRIVE_REMOVABLE
    '    DriveType = "can be removed I.E. Floppy Disk."
    Case DRIVE_FIXED
        DriveType = "FIXED"
    Case Else
        DriveType = "NOT FIXED"
    End Select
    
End Function
      
Public Function GetPrivateINI(ByVal PrivateINI As String, ByVal AppName As String, ByVal keyword As String) As String
'return a string from the Private.INI File
    Dim result As String * 128, StrLen As Integer
    
    StrLen = GetPrivateProfileString(AppName, keyword, "", result, Len(result), PrivateINI)
    GetPrivateINI = Left$(result, StrLen)
End Function
Public Function SetPrivateINI(ByVal PrivateINI As String, ByVal AppName As String, ByVal keyword As String, ByVal keyval As String) As Integer
'write to the private.ini file, returns true or false to indicate success
    SetPrivateINI = WritePrivateProfileString(AppName, keyword, keyval, PrivateINI)
End Function
Function GetWindowsDir() As String
Dim lstrWindowsPath As String
Dim llngReturnValue As Long
Dim llngPathBufferSize As Long
Dim lstrPathBuff As String

    llngPathBufferSize = 255
    lstrPathBuff = Space$(llngPathBufferSize)
    
    llngReturnValue = GetWindowsDirectory(lstrPathBuff, llngPathBufferSize)
       
    If llngReturnValue = 0 Then
        Exit Function
    End If
    
    GetWindowsDir = Left$(lstrPathBuff, (Len(Trim$(lstrPathBuff)) - 1))

End Function
Public Function GetTempDir() As String
Dim lstrTempDir As String

    lstrTempDir = Environ("TEMP")

    If Dir(lstrTempDir, vbDirectory) = "." Then
        lstrTempDir = Environ("TMP")
        If Dir(lstrTempDir, vbDirectory) = "." Then
            lstrTempDir = GetWindowsDir() & "\TEMP"
            If Dir(lstrTempDir, vbDirectory) = "." Then
                On Error Resume Next
                MkDir Left$(GetWindowsDir, 3) & "TEMP"
                lstrTempDir = Left$(GetWindowsDir, 3) & "TEMP"
            End If
        End If
    End If
    GetTempDir = lstrTempDir & "\"
    
End Function

Function ModifyTimeStamp(pstrFilename As String)
    Dim lstrFile As String
    Dim lintAnyThing As Integer
    Dim lintFreeFile As Integer
    
    lstrFile = Dir$(pstrFilename)
    If lstrFile = "" Then GoTo NoSuchFile
    
    On Error GoTo FileError
        lintFreeFile = FreeFile
        Open pstrFilename For Binary As lintFreeFile
        Get lintFreeFile, 1, lintAnyThing
        Put lintFreeFile, 1, lintAnyThing
        Close lintFreeFile
    Exit Function
FileError:
    MsgBox "Unable to time-stamp file", 16, "Error", , gconstrTitlPrefix & "ModifyTimeStamp"
    Exit Function
NoSuchFile:
    MsgBox "That file does not exist!", 16, "Error", , gconstrTitlPrefix & "ModifyTimeStamp"
End Function
Sub RunNWait(pstrProgName As String)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ReturnValue As Integer

    On Error GoTo ErrHandler
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    ReturnValue = CreateProcessA(0&, pstrProgName, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish:
    Do
       ReturnValue = WaitForSingleObject(proc.hProcess, 0)
       DoEvents
       Loop Until ReturnValue <> 258

    ReturnValue = CloseHandle(proc.hProcess)
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RunNWait", "", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    'Case gconIntErrHandEndProgram
    '    LastChanceCafe
    Case Else
        Resume Next
    End Select



End Sub
Sub RunNDontWait(pstrProgName As String)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ReturnValue As Integer

    On Error GoTo ErrHandler
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    ReturnValue = CreateProcessA(0&, pstrProgName, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish:
   ' Do
       ReturnValue = WaitForSingleObject(proc.hProcess, 0)
       DoEvents
   '    Loop Until ReturnValue <> 258

    ReturnValue = CloseHandle(proc.hProcess)
Exit Sub
ErrHandler:
    
    Select Case GlobalErrorHandler(Err.Number, "RunNDontWait", "", False)
    Case gconIntErrHandRetry
        Resume
    Case gconIntErrHandExitFunction
        Exit Sub
    'Case gconIntErrHandEndProgram
    '    LastChanceCafe
    Case Else
        Resume Next
    End Select



End Sub
Function GlobalErrorHandler(pintErrNum As Integer, pstrFunctionName As String, _
    Optional pstrUserDef As Variant, Optional pbooVital As Variant, Optional pvarErrorStageNum As Variant) As Integer
Dim lintMsgRetVal As Integer
Dim lstrStageMsg As String
Dim lstrErrorDescription As String
Dim lerrLoop As Error
Dim lstrExtraErrInfo As String
Dim lstrErrHelpFile As String
Dim llngErrHelpContext As Long

    If IsMissing(pvarErrorStageNum) Then
        lstrStageMsg = vbCrLf & vbCrLf & "TIME UP: " & DateDiff("n", gdatSystemStartTime, Now())
    Else
        lstrStageMsg = vbCrLf & vbCrLf & "STAGE: " & pvarErrorStageNum & vbTab & "TIME UP: " & DateDiff("n", gdatSystemStartTime, Now())
    End If
    
    lstrErrorDescription = GetErrorDescription(pintErrNum)
    
    If lstrErrorDescription = "Application-defined or object-defined error" Then
        lstrExtraErrInfo = vbCrLf & vbCrLf & "An internal component has passed an error to the program!" & vbCrLf & vbCrLf & vbTab
        For Each lerrLoop In Errors
            With lerrLoop
                lstrExtraErrInfo = lstrExtraErrInfo & "COMPONENT: " & .Source & "  Error Number " & _
                    .Number & vbCrLf & .Description & vbCrLf & vbCrLf
                lstrErrHelpFile = .HelpFile
                llngErrHelpContext = .HelpContext
            End With
    Next
        
    End If
    If DebugVersion Then ' Will only happen when in VB Development
        MsgBox pintErrNum & " " & lstrErrorDescription & lstrExtraErrInfo & lstrStageMsg, _
            vbMsgBoxHelpButton, gconstrTitlPrefix & "Debug Error frm :-" & pstrFunctionName, _
            lstrErrHelpFile, llngErrHelpContext
        GlobalErrorHandler = gconIntErrHandRetry
        Debug.Assert (False)
        'Now press F8 to leave this function and find the error.
        Exit Function
    End If
    
    If IsMissing(pbooVital) Then
        pbooVital = False
    End If

    'Busy False, Screen.ActiveForm
     
    Screen.MousePointer = vbNormal
    Screen.ActiveForm.Enabled = True
        
    Select Case pintErrNum
    Case 52
        MsgBox "A Network access error has occured!" & _
        vbCrLf & vbCrLf & "TECHNICAL DETAILS:" & vbCrLf & _
        "Error Number " & pintErrNum & " " & lstrErrorDescription & vbCrLf & vbCrLf & _
        gconstrPleaseLeaveError & lstrStageMsg, vbCritical, gconstrTitlPrefix & "Error From - " & pstrFunctionName
    'case 3022 ' duplicate
    Case 68
        lintMsgRetVal = MsgBox("The network is not available!" & vbCrLf & _
            "Please contact you Technical support office and inform them." & _
            "" & lstrStageMsg, vbRetryCancel, gconstrTitlPrefix & "Error From - " & pstrFunctionName)
        If lintMsgRetVal = vbRetry Then
            GlobalErrorHandler = gconIntErrHandRetry
        End If
    Case 3186
        lintMsgRetVal = MsgBox("Someone else is updating the database, please try again" & _
            vbCrLf & vbCrLf & "You must Wait and Rety, to allow information to be saved." & _
            "" & lstrStageMsg, vbRetryCancel + vbCritical, gconstrTitlPrefix & "Error From - " & pstrFunctionName)
        If lintMsgRetVal = vbRetry Then
            GlobalErrorHandler = gconIntErrHandRetry
        ElseIf lintMsgRetVal = vbCancel And pbooVital = True Then
            lintMsgRetVal = MsgBox("You are about to loose information!" & vbCrLf & vbCrLf & _
                "You MUST! contact your techincal support office!" & vbCrLf & _
                "By using CANCEL will Close the program!" & _
                "" & lstrStageMsg, vbApplicationModal + vbCritical + vbRetryCancel, gconstrTitlPrefix & "Error From - " & _
                pstrFunctionName)
            If lintMsgRetVal = vbRetry Then
                GlobalErrorHandler = gconIntErrHandRetry
            Else
                GlobalErrorHandler = gconIntErrHandEndProgram
            End If
        End If
    
    Case 3261
        lintMsgRetVal = MsgBox("Someone has the " & pstrUserDef & _
            " database open and locked!" & vbCrLf & _
            "Please contact you Technical support office and inform them." & _
            vbCrLf & vbCrLf & "You must Wait and Rety, to allow information to be saved." & _
            "" & lstrStageMsg, vbRetryCancel + vbCritical, gconstrTitlPrefix & "Error From - " & pstrFunctionName)
        If lintMsgRetVal = vbRetry Then
            GlobalErrorHandler = gconIntErrHandRetry
        ElseIf lintMsgRetVal = vbCancel And pbooVital = True Then
            lintMsgRetVal = MsgBox("You are about to loose information!" & vbCrLf & vbCrLf & _
                "You MUST! contact your techincal support office!" & vbCrLf & _
                "By using CANCEL will Close the program!" & _
                "" & lstrStageMsg, vbApplicationModal + vbCritical + vbRetryCancel, gconstrTitlPrefix & "Error From - " & _
                pstrFunctionName)
            If lintMsgRetVal = vbRetry Then
                GlobalErrorHandler = gconIntErrHandRetry
            Else
                GlobalErrorHandler = gconIntErrHandEndProgram
            End If
        End If
    Case 3045
        lintMsgRetVal = MsgBox("Someone has the " & pstrUserDef & _
            " database open and locked!" & vbCrLf & _
            "Please contact you Technical support office and inform them." & _
            vbCrLf & vbCrLf & "You must Wait and Rety, to allow information to be saved." & _
            "" & lstrStageMsg, vbRetryCancel + vbCritical, gconstrTitlPrefix & "Error From - " & pstrFunctionName)
        If lintMsgRetVal = vbRetry Then
            GlobalErrorHandler = gconIntErrHandRetry
        ElseIf lintMsgRetVal = vbCancel And pbooVital = True Then
            'lintMsgRetVal = MsgBox("You are about to loose information!" & vbCrLf & vbCrLf & _
                "You MUST! contact your techincal support office!" & vbCrLf & _
                "By using CANCEL will Close the program!" & _
                "" & lstrStageMsg, vbApplicationModal + vbCritical + vbRetryCancel, gconstrTitlPrefix & "Error From - " & _
                pstrFunctionName)
            'If lintMsgRetVal = vbRetry Then
            '    GlobalErrorHandler = gconIntErrHandRetry
            'Else
                GlobalErrorHandler = gconIntErrHandEndProgram
            'End If
        End If
    Case 3044
        MsgBox "Your computer has not been able to access the required files on the network." & vbCrLf & vbCrLf & _
            "TECHNICAL DETAILS:" & vbCrLf & _
            "Possible reasons for this error :- " & vbCrLf & _
            vbTab & "1. A linked table is pointing at an incorrect path." & vbCrLf & _
            vbTab & "2. Your network card or PC is faulty. " & vbCrLf & _
            vbTab & "3. One of the databases required has been moved. " & vbCrLf & vbCrLf & _
            "Please contact you Technical support office and inform them." & lstrStageMsg, _
            vbApplicationModal + vbCritical, gconstrTitlPrefix & "Error From - " & _
            pstrFunctionName

        GlobalErrorHandler = gconIntErrHandEndProgram
    Case 91
        If gintLastError = pstrFunctionName & " " & pintErrNum Then
            GlobalErrorHandler = gconIntErrHandExitFunction
        End If
    
    Case Else
        MsgBox "You have encountered an unusual error!, " & vbCrLf & _
        "please make notes of when and how it occured and report it!" & _
        vbCrLf & vbCrLf & "TECHNICAL DETAILS:" & vbCrLf & _
        "Error Number " & pintErrNum & " " & lstrErrorDescription & vbCrLf & vbCrLf & lstrExtraErrInfo & _
        gconstrPleaseLeaveError & lstrStageMsg, vbCritical, gconstrTitlPrefix & "Error From - " & pstrFunctionName
        
    End Select
    
    gintLastError = pstrFunctionName & " " & pintErrNum

End Function
Function DebugVersion() As Boolean

    DebugVersion = False
    On Error GoTo DebugError
    
    'leave this debug.print as it is used to check for debug mode
    Debug.Print 1 / 0
    
NormalEnd:
    Exit Function
    
DebugError:
    DebugVersion = True
    Resume NormalEnd

End Function

Function GetErrorDescription(pintErrNumber As Integer) As String
On Error Resume Next
    Err.Raise pintErrNumber

    GetErrorDescription = Err.Description

End Function
Function FileCopyIfNewer(pstrSourceFile As String, pstrDestinationFile As String) As Boolean
Dim lbooDocopy As Boolean

    lbooDocopy = False
    
    On Error GoTo CopyError
    
    If Trim$(Dir(pstrDestinationFile)) <> "" Then
        If FileDateTime(pstrSourceFile) > FileDateTime(pstrDestinationFile) Then
            lbooDocopy = True
        End If
    Else
        FileCopy pstrSourceFile, pstrDestinationFile
    End If
    
NormalExit:
    FileCopyIfNewer = lbooDocopy
    On Error GoTo 0
    Exit Function
    
CopyError:
    If Err = 68 Then
        MsgBox "The network is not available...", , gconstrTitlPrefix & "Accessing Network"
    End If
    Resume NormalExit

End Function
Public Function CurrentMachineName() As String
Dim lSize As Long
Dim sBuffer As String

    sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    lSize = Len(sBuffer)

   If GetComputerName(sBuffer, lSize) Then
       CurrentMachineName = Left$(sBuffer, lSize)
   End If

End Function
Sub CopyFile(SourceFile As String, DestFile As String)
Dim result As Long

    'If Dir(SourceFile) = "" Then
    '   MsgBox Chr(34) & SourceFile & Chr(34) & _
          " is not valid file name."
    'Else
       result = apiCopyFile(SourceFile, DestFile, False)
    'End If
   
End Sub
Function GetNetDir(pobjObject As Form, plngWindowType As Long) As String
Dim BI As BrowseInfo
Dim IDL As ITEMIDLIST
Dim pIdl As Long
Dim sPath As String
Dim SHFI As SHFILEINFO
Dim lstrDisplayName As String
'plngWindowType = 18

    With BI
        .hwndOwner = pobjObject.hwnd
        If SHGetSpecialFolderLocation(ByVal pobjObject.hwnd, ByVal plngWindowType, IDL) = NOERROR Then
            .pIDLRoot = IDL.mkid.cb
        End If
        .pszDisplayName = String$(MAX_PATH, 0)
        .lpszTitle = "Please select a Server Folder!"
        .ulFlags = BIF_RETURNONLYFSDIRS 'GetReturnType()
    End With
    
    ' Show the Browse dialog
    pIdl = SHBrowseForFolder(BI)
    
    If pIdl = 0 Then Exit Function
    
    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList ByVal pIdl, ByVal sPath
    
    GetNetDir = Left(sPath, InStr(sPath, vbNullChar) - 1)
    
    lstrDisplayName = Left$(BI.pszDisplayName, InStr(BI.pszDisplayName, vbNullChar) - 1)
    
    If GetNetDir = "" Then
        GetNetDir = "@\\" & lstrDisplayName
    End If
                                 
    If Right(GetNetDir, 1) <> "\" Then
        GetNetDir = GetNetDir & "\"
    End If
    
    If Right(GetNetDir, 2) = ":\" Then
        GetNetDir = "@" & GetNetDir
    End If
    CoTaskMemFree pIdl
  
End Function

Sub WinCopyDlg(pobjObject As Object, pstrSource As String, pstrDest As String, Optional pvarCustomText As Variant)
Dim result As Long
Dim lenFileop As Long
Dim foBuf() As Byte
Dim fileop As SHFILEOPSTRUCT
    
    lenFileop = LenB(fileop)    ' double word alignment increase
    ReDim foBuf(1 To lenFileop) ' the size of the structure.
    
    With fileop
        .hwnd = pobjObject.hwnd
        .wFunc = FO_COPY
        If Not IsMissing(pvarCustomText) Then
            .lpszProgressTitle = pvarCustomText
            .fFlags = FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_SIMPLEPROGRESS Or FOF_FILESONLY
        Else
            .fFlags = FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_FILESONLY
        End If
        .pFrom = pstrSource & vbNullChar & vbNullChar
        .pTo = pstrDest & vbNullChar & vbNullChar
    End With
    
    ' Now we need to copy the structure into a byte array
    Call CopyMemory(foBuf(1), fileop, lenFileop)
    
    ' Next we move the last 12 bytes by 2 to byte align the data
    Call CopyMemory(foBuf(19), foBuf(21), 12)
    result = SHFileOperation(foBuf(1))
    
    If result <> 0 Then  ' Operation failed
       MsgBox Err.LastDllError 'Show the error returned from
                               'the API.
       Else
       If fileop.fAnyOperationsAborted <> 0 Then
          MsgBox "Operation Failed"
       End If
    End If

End Sub
Sub DisableCloseButton(pobjForm As Form)
Dim Ret As Long

    'Initailise Start
    hMenu = GetSystemMenu(pobjForm.hwnd, 0)
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = SC_CLOSE
    Ret = GetMenuItemInfo(hMenu, MII.wID, False, MII)
    'Initailise End

    Ret = SetIdCloseButton(SwapID)
    If Ret <> 0 Then
    
        If MII.fState = (MII.fState Or MFS_GRAYED) Then
            MII.fState = MII.fState - MFS_GRAYED
        Else
            MII.fState = (MII.fState Or MFS_GRAYED)
        End If
    
        MII.fMask = MIIM_STATE
        Ret = SetMenuItemInfo(hMenu, MII.wID, False, MII)
        If Ret = 0 Then
            Ret = SetIdCloseButton(ResetID)
        End If
    
        Ret = SendMessage(pobjForm.hwnd, WM_NCACTIVATE, True, 0)
    End If
    
End Sub
Private Function SetIdCloseButton(Action As Long) As Long
Dim MenuID As Long
Dim Ret As Long

    MenuID = MII.wID
    If MII.fState = (MII.fState Or MFS_GRAYED) Then
        If Action = SwapID Then
            MII.wID = SC_CLOSE
        Else
            MII.wID = xSC_CLOSE
        End If
    Else
        If Action = SwapID Then
            MII.wID = xSC_CLOSE
        Else
            MII.wID = SC_CLOSE
        End If
    End If

    MII.fMask = MIIM_ID
    Ret = SetMenuItemInfo(hMenu, MenuID, False, MII)
    If Ret = 0 Then
        MII.wID = MenuID
    End If
    SetIdCloseButton = Ret
    
End Function

Function ListLoggedUsersOld(pstrDatabase As String) As String
ReDim msString(1) As String
Dim miLoop As Integer
Dim lintRetVal As Integer
Dim lstrUsers As String
Dim lstrDBLockingFile As String

    lstrDBLockingFile = Left(pstrDatabase, Len(pstrDatabase) - 3) & "ldb"
    
    lintRetVal = LDBUser_GetUsers(msString, lstrDBLockingFile, &H2)
    
    For miLoop = LBound(msString) To UBound(msString)
        If Len(msString(miLoop)) = 0 Then
            Exit For
        End If
        
        lstrUsers = lstrUsers & vbTab & msString(miLoop) & vbCrLf
    Next miLoop
    
    ListLoggedUsersOld = lstrUsers
    
End Function
Function GetHDDSerialNumber(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String

    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    Res = GetVolumeInformation(strDrive, Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetHDDSerialNumber = SerialNum
    
End Function
Function CheckForOtherMMosprog(plngAppHandle As Long) As Boolean
Dim llngDummyVariable As Long
Dim lintlLenTitle As Integer
Dim lstrWinTitle As String * 256

    'initialize the function return as False
    CheckForOtherMMosprog = False
    
    lintlLenTitle = Len(gconstrProductFullName)

    'Get the handle of the first child of the desktop window
    plngAppHandle = GetTopWindow(0)

    'Loop through all top-level windows and search for the sub-string
    'in the Window title
    Do Until plngAppHandle = 0
        llngDummyVariable = GetWindowText(plngAppHandle, lstrWinTitle, 255)
        If Left(lstrWinTitle, lintlLenTitle) = gconstrProductFullName Then
            CheckForOtherMMosprog = True
            Exit Function
        Else
            plngAppHandle = GetWindow(plngAppHandle, GW_HWNDNEXT)
        End If
    Loop
End Function
Function dhHackByte(lngIn As Long, bytByte As Byte) As Byte
Dim dwb As dhDoubleWordByByte
Dim dwl As dhDoubleWordLong

    dwl.DoubleWord = lngIn
    LSet dwb = dwl
    Select Case bytByte
        Case 1
            dhHackByte = dwb.LowWordLowByte
        Case 2
            dhHackByte = dwb.LowWordHighByte
        Case 3
            dhHackByte = dwb.HighWordLowByte
        Case 4
            dhHackByte = dwb.HighWordHighByte
    End Select
    
End Function
Function dhHackWord(lngIn As Long, bytWord As Byte) As Integer
Dim dww As dhDoubleWordByWord
Dim dwl As dhDoubleWordLong

    dwl.DoubleWord = lngIn
    LSet dww = dwl
    Select Case bytWord
        Case 1
            dhHackWord = dww.LowWord
        Case 2
            dhHackWord = dww.HighWord
    End Select
    
End Function
Function WindowLoaded(pstrWindowCaption As String) As Boolean
Dim llngDummyVariable As Long
Dim lintlLenTitle As Integer
Dim lstrWinTitle As String * 256
Dim plngAppHandle As Long

    'initialize the function return as False
    WindowLoaded = False
    
    lintlLenTitle = Len(pstrWindowCaption)

    'Get the handle of the first child of the desktop window
    plngAppHandle = GetTopWindow(0)

    'Loop through all top-level windows and search for the sub-string
    'in the Window title
    Do Until plngAppHandle = 0
        llngDummyVariable = GetWindowText(plngAppHandle, lstrWinTitle, 255)
        If Left(lstrWinTitle, lintlLenTitle) = pstrWindowCaption Then
            WindowLoaded = True
            Exit Function
        Else
            plngAppHandle = GetWindow(plngAppHandle, GW_HWNDNEXT)
        End If
    Loop
End Function
Sub PosPtrOnCtl(pobjControl As Object, Optional plngXOffset As Variant, Optional plngYOffset As Variant)
Dim llngXPos As Integer
Dim llngYPos As Integer

    If IsMissing(plngXOffset) Then
        plngXOffset = 0
    End If
    
    If IsMissing(plngYOffset) Then
        plngYOffset = 0
    End If
    
    llngXPos = (pobjControl.Parent.Left + pobjControl.Left + plngXOffset + _
        (pobjControl.Width / 2) + 60) / Screen.TwipsPerPixelX
    llngYPos = (pobjControl.Parent.Top + pobjControl.Top + plngYOffset + _
        (pobjControl.Height / 2) + 360) / Screen.TwipsPerPixelY
    
    SetCursorPos llngXPos, llngYPos
    
End Sub



