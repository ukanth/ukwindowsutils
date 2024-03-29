Attribute VB_Name = "Module1"
'**************************************
'Windows API/Global Declarations for :Vi
'     ew Windows XP CD Key
'**************************************
Option Explicit


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Private Const REG_BINARY = 3
    Private Const HKEY_LOCAL_MACHINE = &H80000002
    Private Const ERROR_SUCCESS = 0&
'**************************************
' Name: View Windows XP CD Key
' Description:Function: sGetXPCDKey() wi
'     ll return the CD Key for Windows XP in t
'     he format XXXXX-XXXXX-XXXXX-XXXXX-XXXXX.
'
'sGetXPCDKey() -

Public Function sGetXPCDKey() As String
    'Read the value of:
    'HKLM\SOFTWARE\MICROSOFT\Windows NT\Curr
    '     entVersion\DigitalProductId
    Dim bDigitalProductID() As Byte
    Dim bProductKey() As Byte
    Dim ilByte As Long
    Dim lDataLen As Long
    Dim hKey As Long
    'Open the registry key: HKLM\SOFTWARE\MI
    '     CROSOFT\Windows NT\CurrentVersion


    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\Windows NT\CurrentVersion", hKey) = ERROR_SUCCESS Then
        lDataLen = 164
        ReDim Preserve bDigitalProductID(lDataLen)
        'Read the value of DigitalProductID


        If RegQueryValueEx(hKey, "DigitalProductId", 0&, REG_BINARY, bDigitalProductID(0), lDataLen) = ERROR_SUCCESS Then
            'Get the Product Key, 15 bytes long, off
            '     set by 52 bytes
            ReDim Preserve bProductKey(14)


            For ilByte = 52 To 66
                bProductKey(ilByte - 52) = bDigitalProductID(ilByte)
            Next ilByte
        Else
            'ERROR: Could not read "DigitalProductID
            '     "
            sGetXPCDKey = ""
            Exit Function
        End If
    Else
        'ERROR: Could not open "HKLM\SOFTWARE\MI
        '     CROSOFT\Windows NT\CurrentVersion"
        sGetXPCDKey = ""
        Exit Function
    End If
    'Now we are going to 'base24' decode the
    '     Product Key
    Dim bKeyChars(0 To 24) As Byte
    'Possible characters in the CD Key:
    bKeyChars(0) = Asc("B")
    bKeyChars(1) = Asc("C")
    bKeyChars(2) = Asc("D")
    bKeyChars(3) = Asc("F")
    bKeyChars(4) = Asc("G")
    bKeyChars(5) = Asc("H")
    bKeyChars(6) = Asc("J")
    bKeyChars(7) = Asc("K")
    bKeyChars(8) = Asc("M")
    bKeyChars(9) = Asc("P")
    bKeyChars(10) = Asc("Q")
    bKeyChars(11) = Asc("R")
    bKeyChars(12) = Asc("T")
    bKeyChars(13) = Asc("V")
    bKeyChars(14) = Asc("W")
    bKeyChars(15) = Asc("X")
    bKeyChars(16) = Asc("Y")
    bKeyChars(17) = Asc("2")
    bKeyChars(18) = Asc("3")
    bKeyChars(19) = Asc("4")
    bKeyChars(20) = Asc("6")
    bKeyChars(21) = Asc("7")
    bKeyChars(22) = Asc("8")
    bKeyChars(23) = Asc("9")
    Dim nCur As Integer
    Dim sCDKey As String
    Dim ilKeyByte As Long
    Dim ilBit As Long


    For ilByte = 24 To 0 Step -1
        'Step through each character in the CD k
        '     ey
        nCur = 0


        For ilKeyByte = 14 To 0 Step -1
            'Step through each byte in the Product K
            '     ey
            nCur = nCur * 256 Xor bProductKey(ilKeyByte)
            bProductKey(ilKeyByte) = Int(nCur / 24)
            nCur = nCur Mod 24
        Next ilKeyByte
        sCDKey = Chr(bKeyChars(nCur)) & sCDKey
        If ilByte Mod 5 = 0 And ilByte <> 0 Then sCDKey = "-" & sCDKey
    Next ilByte
    sGetXPCDKey = sCDKey
End Function

