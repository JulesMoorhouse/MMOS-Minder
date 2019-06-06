Attribute VB_Name = "ModMinder"
Option Explicit

Dim mstrAppPath As String

'the following declares are for the free resources check
#If Win32 Then
      Private Type STARTUPINFO
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

      Private Type PROCESS_INFORMATION
         hProcess As Long
         hThread As Long
         dwProcessID As Long
         dwThreadID As Long
      End Type

      Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal _
         hHandle As Long, ByVal dwMilliseconds As Long) As Long

      Private Declare Function CreateProcessA Lib "KERNEL32" (ByVal _
         lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
         lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
         lpStartupInfo As STARTUPINFO, lpProcessInformation As _
         PROCESS_INFORMATION) As Long

      Private Declare Function CloseHandle Lib "KERNEL32" (ByVal _
         hObject As Long) As Long

      Private Const NORMAL_PRIORITY_CLASS = &H20&
      Private Const INFINITE = -1&

    Type SYSTEM_INFO
            dwOemID As Long
            dwPageSize As Long
            lpMinimumApplicationAddress As Long
            lpMaximumApplicationAddress As Long
            dwActiveProcessorMask As Long
            dwNumberOrfProcessors As Long
            dwProcessorType As Long
            dwAllocationGranularity As Long
            dwReserved As Long
    End Type
    Type OSVERSIONINFO
            dwOSVersionInfoSize As Long
            dwMajorVersion As Long
            dwMinorVersion As Long
            dwBuildNumber As Long
            dwPlatformId As Long
            szCSDVersion As String * 128
    End Type
    Type MEMORYSTATUS
            dwLength As Long
            dwMemoryLoad As Long
            dwTotalPhys As Long
            dwAvailPhys As Long
            dwTotalPageFile As Long
            dwAvailPageFile As Long
            dwTotalVirtual As Long
            dwAvailVirtual As Long
    End Type
    Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" _
       (lpVersionInformation As OSVERSIONINFO) As Long
    Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As _
       MEMORYSTATUS)
    Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As _
       SYSTEM_INFO)
    Public Const PROCESSOR_INTEL_386 = 386
    Public Const PROCESSOR_INTEL_486 = 486
    Public Const PROCESSOR_INTEL_PENTIUM = 586
    Public Const PROCESSOR_MIPS_R4000 = 4000
    Public Const PROCESSOR_ALPHA_21064 = 21064
#Else
   Global Const WF_CPU286 = &H2
   Global Const WF_CPU386 = &H4
   Global Const WF_CPU486 = &H8
   Global Const WF_80x87 = &H400
   Global Const WF_STANDARD = &H10
   Global Const WF_ENHANCED = &H20
   Global Const WF_WINNT = &H4000
   Type SYSHEAPINFO
      dwSize As Long
      wUserFreePercent As Integer
      wGDIFreePercent As Integer
      hUserSegment As Integer
      hGDISegment As Integer
   End Type
   Declare Function GetVersion Lib "kernel" () As Long
   Declare Function GetWinFlags Lib "kernel" () As Long
   Declare Function GetFreeSpace Lib "kernel" (ByVal wFlags As Integer) _
      As Long
   Declare Function GlobalCompact Lib "kernel" _
      (ByVal dwMinFree As Long) As Long
   Declare Function SystemHeapInfo Lib "toolhelp.dll" _
      (shi As SYSHEAPINFO) As Integer
#End If

Global gbooDoneTimeEvent As Boolean

Function CheckWindowsVersion() As String

Dim llngVer As Long
Dim lstrVerMajor As String
Dim lstrVerMinor As String
Dim llngStatus As Long
Dim lstrString As String
Dim lintReturn As Integer
Dim lstrVerBuild As String

    #If Win32 Then
        ' Get operating system and version.
        Dim verinfo As OSVERSIONINFO
        verinfo.dwOSVersionInfoSize = Len(verinfo)
        lintReturn = GetVersionEx(verinfo)
        If lintReturn = 0 Then
            MsgBox "Error Getting Version Information"
            Exit Function
        End If
        Select Case verinfo.dwPlatformId
            Case 0
                lstrString = lstrString & "Windows 32s "
            Case 1
                lstrString = lstrString & "Windows 95 "
            Case 2
                lstrString = lstrString & "Windows NT "
        End Select
    
        lstrVerMajor = verinfo.dwMajorVersion
        lstrVerMinor = verinfo.dwMinorVersion
        lstrVerBuild = verinfo.dwBuildNumber
        lstrString = lstrString & lstrVerMajor & "." & lstrVerMinor
        lstrString = lstrString & " (Build " & lstrVerBuild & ")" & vbCrLf & vbCrLf
    
    #Else
        ' Get operating system and version.
        llngVer = GetVersion()
        llngStatus = GetWinFlags()
    
        If llngStatus And WF_WINNT Then
            lstrString = lstrString & "Microsoft Windows NT "
        Else
            lstrString = lstrString & "Microsoft Windows "
        End If
        lstrVerMajor = Format$(llngVer And &HFF)
        lstrVerMinor = Format$(llngVer \ &H100, "00")
        lstrString = lstrString & lstrVerMajor & "." & lstrVerMinor & vbCrLf
    
    #End If
    
    CheckWindowsVersion = lstrString

End Function

Sub DelWinTmp()

    On Error Resume Next
    
    Select Case Left$(CheckWindowsVersion, 20)
    Case "Microsoft Windows NT"
        'no defrag
    Case Else
        Kill "c:\*.bak"
        Kill "c:\*.old"
        Kill "c:\*.log"
        Kill "c:\*.tmp"
        Kill "c:\Windows\*.bak"
        Kill "c:\Windows\*.old"
        Kill "c:\Windows\*.log"
        Kill "c:\Windows\*.tmp"
    End Select

End Sub

Sub xFileScan()
Const lstrDrive = "C:\"
Dim lstrSearchPath As String
Dim lstrSearchFile As String
Dim lintIndex As Integer
Dim lstrdepthfirstsearch(1 To 100) As String

lintIndex = 1
lstrdepthfirstsearch(lintIndex) = lstrDrive
Do
    lstrSearchPath = lstrdepthfirstsearch(lintIndex)
    lstrdepthfirstsearch(lintIndex) = ""
    lintIndex = lintIndex - 1
    GoSub ExpandNode

Loop While lstrdepthfirstsearch(1) <> ""
    
Exit Sub

ExpandNode:
lstrSearchFile = Dir(lstrSearchPath, vbDirectory)
Do While lstrSearchFile <> ""
    If lstrSearchFile <> "." And lstrSearchFile <> ".." Then
        If (GetAttr(lstrSearchPath & lstrSearchFile) And vbDirectory) = vbDirectory Then
            lintIndex = lintIndex + 1
            lstrdepthfirstsearch(lintIndex) = lstrSearchPath & lstrSearchFile & "\"

        End If
    End If
    lstrSearchFile = Dir
Loop

Return

ErrorHandler:
    Resume Next

End Sub

Sub PerformDDE(frm As Form, ByVal lstrGroup As String, ByVal lstrCmd As String, ByVal lstrTitle As String, ByVal intDDE As Integer, ByVal fLog As Boolean)
Const lstrCOMMA$ = ","
Const lstrRESTORE$ = ", 1)]"
Const lstrENDCMD$ = ")]"
Const lstrSHOWGRP$ = "[ShowGroup("
Const lstrADDGRP$ = "[CreateGroup("
Const lstrREPLITEM$ = "[ReplaceItem("
Const lstrADDITEM$ = "[AddItem("
Dim intIdx As Integer
Dim intRetry As Integer

    For intRetry = 1 To 20
        On Error Resume Next
        frm.lblDDE.LinkTopic = "PROGMAN|PROGMAN"
        If Err = 0 Then
            Exit For
        End If
        DoEvents
    Next intRetry
        
    frm.lblDDE.LinkMode = 2
    For intIdx = 1 To 10
      DoEvents
    Next
    frm.lblDDE.LinkTimeout = 100
    
    On Error Resume Next
    
    If Err = 0 Then
        Select Case intDDE
        Case 1 'Add Icon / Shortcut
            frm.lblDDE.LinkExecute lstrREPLITEM & lstrTitle & lstrENDCMD
            Err = 0
            frm.lblDDE.LinkExecute lstrADDITEM & lstrCmd & lstrCOMMA & lstrTitle & String$(3, lstrCOMMA) & lstrENDCMD
        Case 2 ' Add Program Group
            #If Win16 Then
                frm.lblDDE.LinkExecute lstrADDGRP & lstrGroup & lstrCOMMA & lstrCmd & lstrENDCMD
            #Else
                frm.lblDDE.LinkExecute lstrADDGRP & lstrGroup & lstrENDCMD
            #End If
    '       frm.lblDDE.LinkExecute lstrSHOWGRP & lstrGroup & lstrRESTORE
        End Select
    End If
    
    frm.lblDDE.LinkMode = 0
    frm.lblDDE.LinkTopic = ""
    
    Err = 0
    
End Sub
Sub DelUnauthorisedFiles()
Dim lstrSystemFile As String

    On Error GoTo DelSystem_ErrHand
    
    lstrSystemFile = Dir(mstrAppPath, vbDirectory)
    Do While lstrSystemFile <> ""
        If lstrSystemFile <> "." And lstrSystemFile <> ".." Then
            If GetAttr(mstrAppPath & lstrSystemFile) <> 16 And UCase(lstrSystemFile) <> _
               "MINDER.EXE" And UCase(lstrSystemFile) <> "LOADER.EXE" And _
               UCase(lstrSystemFile) <> UCase(gconstrStaticLdr) Then
                Kill mstrAppPath & lstrSystemFile
            End If
        End If
        lstrSystemFile = Dir
    Loop
    
    Exit Sub
    
DelSystem_ErrHand:
    Select Case Err
    Case 76 'System Path not found
        frmMinder.BackColor = vbRed
        Exit Sub
    Case Else
        Resume Next
    End Select

End Sub

Sub DelTempFiles()
Dim lstrTempPath As String
Dim lstrTempFile As String

    lstrTempPath = Environ("TEMP")
    
    On Error Resume Next
    
    If lstrTempPath <> "" Then
    
        lstrTempFile = Dir(lstrTempPath & "\" & "*.tmp")
        Do Until lstrTempFile = ""
    
            If Format(FileDateTime(lstrTempPath & "\" & lstrTempFile), "DD/MM/YY") _
               < Format(Date, "DD/MM/YY") Then ' if yesterday or before
                SetAttr lstrTempPath & "\" & lstrTempFile, vbNormal
                Kill lstrTempPath & "\" & lstrTempFile
                    
            End If
            
            lstrTempFile = Dir
        Loop
        
        lstrTempFile = Dir(lstrTempPath & "\" & "~*.*")
        
        Do Until lstrTempFile = ""
            If Format(FileDateTime(lstrTempPath & "\" & lstrTempFile), "DD/MM/YY") _
               < Format(Date, "DD/MM/YY") Then ' if yesterday or before
                SetAttr lstrTempPath & "\" & lstrTempFile, vbNormal
                Kill lstrTempPath & "\" & lstrTempFile
            End If
            lstrTempFile = Dir
        Loop
        
        Kill lstrTempPath & "\" & "*.tmp"
        Kill lstrTempPath & "\" & "~*.*"
    End If

End Sub

Sub Main()
Dim lstrDestinationPath As String
Dim lstrSourceFile As String
Dim lstrDestinationFile As String
Dim lbooCopyDone As Boolean

    On Error Resume Next
    gbooDoneTimeEvent = False
    
    If gbooDoneTimeEvent = True Then End
    
    Load frmMinder
    frmMinder.Show
    DoEvents
    Screen.MousePointer = vbHourglass
    
    If UCase(Dir(Trim$(App.Path) & "\" & gconstrStaticLdr, vbNormal)) = UCase(gconstrStaticLdr) Then
        CheckStaticCipher
    End If
    
    mstrAppPath = AppPath
    'Always delete temporary files
    DelTempFiles
    DelWinTmp
    'Add Clean into Startup Group
    If GetSetting("Minder", "Windows", "Startup") <> "Yes" Then
        PerformDDE frmMinder, "StartUp", "", vbNull, 2, True
        PerformDDE frmMinder, "StartUp", mstrAppPath & "Minder.exe ", "Minder", 1, True
        SaveSetting "Minder", "Windows", "Startup", "Yes"
    End If

    lstrDestinationPath = App.Path & "\"
    
    lstrSourceFile = gstrStatic.strServerPath & "Loader.exe"
    lstrDestinationFile = lstrDestinationPath & "Loader.exe"
    
    lbooCopyDone = FileCopyIfNewer(lstrSourceFile, lstrDestinationFile)
                    
    Screen.MousePointer = vbDefault
    Select Case UCase$(Command)
        Case Is <> "APP"
            Unload frmMinder
    End Select

End Sub


