Attribute VB_Name = "modUnlock"
Option Explicit
Public Type Key
    strCompanyName      As String
    strCompanyTelephone As String
    strCompanyContact   As String
    strCoverDate        As String
    strRetVal        As String
    strUnlockKey     As String
    strUserNum          As String
    strCover            As String
End Type

Global gstrKey As Key
Function DeCodeUserNum(pstrString As String) As String
Dim lintFirstChar As Integer
Dim lintSecondChar As Integer

    lintFirstChar = Asc(Left$(pstrString, 1))
    lintSecondChar = Asc(Right(pstrString, 1))
    
    DeCodeUserNum = "X"
    
    If lintSecondChar - lintFirstChar = 5 Then
        DeCodeUserNum = "05"
    End If
    If lintSecondChar - lintFirstChar = 3 Then
        DeCodeUserNum = "25"
    End If
    If lintSecondChar - lintFirstChar = -1 Then
        DeCodeUserNum = "10"
    End If
    If lintSecondChar - lintFirstChar = -3 Then
        DeCodeUserNum = "30"
    End If
    
End Function
Function DeCodeCover(pstrString As String) As String
Dim lintFirstChar As Integer
Dim lintSecondChar As Integer

    lintFirstChar = Asc(Left$(pstrString, 1))
    lintSecondChar = Asc(Right(pstrString, 1))
    
    DeCodeCover = "X"
    
    If lintSecondChar - lintFirstChar = 5 Then
        DeCodeCover = "None"
    End If
    If lintSecondChar - lintFirstChar = 3 Then
        DeCodeCover = "Basic"
    End If
    If lintSecondChar - lintFirstChar = -1 Then
        DeCodeCover = "Comp"
    End If
    If lintSecondChar - lintFirstChar = -3 Then
        DeCodeCover = "Wazz"
    End If

End Function
Function DecodeDecoyChar(pstrChar As String) As String
Dim lret As String

    Select Case Asc(pstrChar)
    Case 48 To 57 ' 0 to 9
        Select Case Asc(pstrChar)
        Case 48: lret = "4": Case 49: lret = "2": Case 50: lret = "9": Case 51: lret = "0"
        Case 52: lret = "8": Case 53: lret = "7": Case 54: lret = "3": Case 55: lret = "1"
        Case 56: lret = "6": Case 57: lret = "5": End Select
        
    Case 65 To 90 ' A to Z
        Select Case Asc(pstrChar)
        Case 65 To 77
            lret = Chr(Asc(pstrChar) + 13)
        Case 78 To 90
            lret = Chr(Asc(pstrChar) - 13)
        End Select
    Case Asc("%")
        lret = Chr(45)
    Case Asc("&")
        lret = Chr(92)
    Case Asc("£")
        lret = Chr(39)
    Case Asc("$")
        lret = Chr(32)
    End Select
    
    DecodeDecoyChar = lret
    
End Function
Function DeCodeCoverDate(pstrString As String) As String
Dim lintArrInc As Integer
Dim lstrM As String

    For lintArrInc = 1 To Len(pstrString)
        Select Case Mid$(pstrString, lintArrInc, 1)
        Case "X": lstrM = "0"
        Case "S": lstrM = "1"
        Case "G": lstrM = "2"
        Case "H": lstrM = "3"
        Case "R": lstrM = "4"
        Case "W": lstrM = "5"
        Case "P": lstrM = "6"
        Case "A": lstrM = "7"
        Case "K": lstrM = "8"
        Case "L": lstrM = "9"
        End Select
        DeCodeCoverDate = DeCodeCoverDate & lstrM
    Next lintArrInc
    
End Function
Sub GenerateKey()

    With gstrKey
        .strRetVal = Scramble(.strCompanyName & " X " & _
            .strCompanyTelephone & " X " & _
            .strCompanyContact) '& " X " & _
            Scramble(.strUsersNum)
    End With
    
End Sub
Function Scramble(pstrString As String) As String
Dim lintArrInc As Integer
Dim lstrTemp As String
    For lintArrInc = 1 To Len(pstrString)
        Scramble = Val(Scramble) + (Asc(Mid(pstrString, lintArrInc, 1)) * lintArrInc)
    Next lintArrInc

End Function
Function Decode(pstrCompanyName As String, pstrCompanyTelephoneNum As String, pstrName As String) As String
Dim lstrUnlockCode As String
Dim lstrUserNum As String
Dim lstrCoverNum As String
Dim lstrCoverDate As String

    Decode = "39"
    
    With gstrKey
        lstrUnlockCode = Mid$(.strUnlockKey, 1, 5) & _
            Mid$(.strUnlockKey, 6, 5) & Mid$(.strUnlockKey, 11, 5) & _
             Mid$(.strUnlockKey, 16, 5)
             
        If lstrUnlockCode <> "" Then
            .strUserNum = DeCodeUserNum(Mid$(lstrUnlockCode, 3, 1) & Mid$(lstrUnlockCode, 13, 1))
            .strCover = DeCodeCover(Mid$(lstrUnlockCode, 12, 1) & Mid$(lstrUnlockCode, 17, 1))
            .strCoverDate = DeCodeCoverDate(Mid$(lstrUnlockCode, 1, 1) & Mid$(lstrUnlockCode, 19, 1) & _
                Mid$(lstrUnlockCode, 5, 1) & Mid$(lstrUnlockCode, 18, 1) & Mid$(lstrUnlockCode, 6, 1))
                
            Decode = "44"
            If lstrUserNum <> "X" And lstrCoverNum <> "X" Then
                If DecodeDecoyChar(Mid$(lstrUnlockCode, 2, 1)) = Mid$(UCase$(pstrCompanyName), 4, 1) Then ', , "should be 1st word 4th char"
                    If DecodeDecoyChar(Mid$(lstrUnlockCode, 8, 1)) = Mid$(UCase$(pstrCompanyTelephoneNum), 2, 1) Then ', , "should be 2nd word 2nd char"
                        If DecodeDecoyChar(Mid$(lstrUnlockCode, 14, 1)) = Mid$(UCase$(pstrName), 5, 1) Then ', , "should be 3rd word 5th char"
                        
                            'MsgBox "Valid Key!  strUserNum=" & .strUserNum & " strCoverNum=" & .strCover & " strCoverDate=" & .strCoverDate
                            Decode = "21"
                            Exit Function
                        End If
                    End If
                End If
            Else
                ClearDecode
            End If
        End If
    End With
    
    'MsgBox "invalid key!"
End Function
Sub ClearDecode()

    With gstrKey
        .strUserNum = ""
        .strCoverDate = ""
        .strCover = ""
    End With
    
End Sub
Function ConvertJulian(JulianDate As Long)
   ConvertJulian = DateSerial(2000 + Int(JulianDate / 1000), _
                 1, JulianDate Mod 1000)
End Function



