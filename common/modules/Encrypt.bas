Attribute VB_Name = "modEncrypt"
Option Explicit
Const mconstrPassword = "DEVELOPED BY MINDWARP CONSULTANCY LTD - JULES "
Global Const gconEncryptStatic = "STATIC"
Global Const gconEncryptDataFile = "DATAFILE"

Sub Encrypt(pstrOutputFile As String, pstrOption As String, Optional pstrInputFile As String)

    Dim strHead As String
    Dim strT As String
    Dim strA As String
    Dim cphX As New clsCipher
    Dim lngN As Long
    Dim lintFreeFile As Integer

    Select Case pstrOption
    Case gconEncryptStatic
        strA = ReadBuffer
    Case gconEncryptDataFile
        lintFreeFile = FreeFile
        Open pstrInputFile For Binary As #lintFreeFile
        'Load entire file into strA
        strA = Space$(LOF(lintFreeFile))
        Get #lintFreeFile, , strA
        Close #lintFreeFile
    End Select
    
    strT = Hash(Date & Str(Timer))
    strHead = "33" & strT & Hash(strT & mconstrPassword)
    'Do the encryption
    cphX.KeyString = strHead
    cphX.Text = strA
    cphX.DoXor
    cphX.Stretch
    strA = cphX.Text
    
    lintFreeFile = FreeFile
    Open pstrOutputFile For Output As #lintFreeFile
    Print #lintFreeFile, strHead
    'Write encrypted data
    lngN = 1
    Do
        Print #lintFreeFile, Mid(strA, lngN, 70)
        lngN = lngN + 70
    Loop Until lngN > Len(strA)
    Close #lintFreeFile
    
End Sub

Sub Decrypt(pstrInputFile As String, pstrOption As String, Optional pstrOutputFile As Variant)

    Dim strHead As String
    Dim strA As String
    Dim strT As String
    Dim cphX As New clsCipher
    Dim lngN As Long
    Dim lintFreeFile As Integer
        
    lintFreeFile = FreeFile
    
    Open pstrInputFile For Input As #lintFreeFile
    Line Input #lintFreeFile, strHead
    Close #lintFreeFile
    
    strT = Mid(strHead, 9, 8)
    
    lintFreeFile = FreeFile
    Open pstrInputFile For Input As #lintFreeFile
    Line Input #lintFreeFile, strHead
    Do Until EOF(lintFreeFile)
        Line Input #lintFreeFile, strT
        strA = strA & strT
    Loop
    Close #lintFreeFile
    
    'Decrypted file contents
    cphX.KeyString = strHead
    cphX.Text = strA
    cphX.Shrink
    cphX.DoXor
    strA = cphX.Text
    
    Select Case pstrOption
    Case gconEncryptStatic
        WriteBuffer strA
    Case gconEncryptDataFile
        lintFreeFile = FreeFile
        Open pstrOutputFile For Binary As #lintFreeFile
        Put #lintFreeFile, , strA
        Close #lintFreeFile
    End Select
    
End Sub

Function Hash(strA As String) As String

    Dim cphHash As New clsCipher
    
    cphHash.KeyString = strA & "123456"
    cphHash.Text = strA & "123456"
    cphHash.DoXor
    cphHash.Stretch
    cphHash.KeyString = cphHash.Text
    cphHash.Text = "123456"
    cphHash.DoXor
    cphHash.Stretch
    Hash = cphHash.Text
    
End Function

