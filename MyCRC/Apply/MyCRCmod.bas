Attribute VB_Name = "MyCRCmod"
Option Explicit

Function AppendCRC(FilePath As String, CheckIf00 As Boolean) As String

'dim
Dim iFreeFile As Integer
Dim CBbytes As Long
Dim Rbyte As Byte
Dim strRbyte As String
Dim lnCount As Long
Dim MyCRC As String
'Dim FileStr As String
Dim BytesBefore As Long
Dim FileArr() As Byte

BytesBefore = 16

If Dir$(FilePath) = "" Then
    AppendCRC = "file not found."
    Exit Function
End If

'open the file
iFreeFile = FreeFile
Open FilePath For Binary Access Read Write As #iFreeFile

    'get CBbytes
    CBbytes = LOF(iFreeFile) - BytesBefore + 1

    'free memory space
    ReDim FileArr(LOF(iFreeFile)) As Byte

    'input byte array
    Do While Not EOF(iFreeFile)
        Get #iFreeFile, , FileArr()
    Loop
    
    'check that all bytes to write to are 00
    If CheckIf00 = True Then
        For lnCount = 0 To BytesBefore - 1
            If FileArr(CBbytes + lnCount) <> 0 Then
                AppendCRC = "Byte at position " & CBbytes + lnCount & "(" & _
                    Hex(CBbytes + lnCount) & ") is not 00 (value: " & Hex(FileArr(CBbytes + lnCount)) & ")."
                Exit Function
            End If
        Next lnCount
    End If
    
    'zero all final variables
    For lnCount = 0 To BytesBefore - 1
        FileArr(CBbytes + lnCount) = 0
    Next lnCount
    
    'compute CRC
    Dim m_CRC As clsCRC
    'start CRC engine
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    MyCRC = Hex$(m_CRC.CalculateBytes(FileArr))
    Set m_CRC = Nothing

    'check that position is ok
    If CBbytes <= 0 Then
        AppendCRC = "Writing position is before beginning of file."
        Close #iFreeFile
        Exit Function
    End If
    
    'ok bytes are 00 proceed
    For lnCount = 0 To Len(MyCRC) - 1
        Put #1, CBbytes + lnCount + 1, Asc(Mid(MyCRC, lnCount + 1, 1))
    Next lnCount
    
    AppendCRC = "Appended CRC (" & MyCRC & ") at end of file."

Close #iFreeFile

End Function



Function CheckCRC(FilePath As String) As Integer

'dim
Dim iFreeFile As Integer
Dim CBbytes As Long
Dim lnCount As Long
Dim MyCRC As String
Dim BytesBefore As Long
Dim FileArr() As Byte
Dim ReadCRC As String
Dim NotZero As Boolean

BytesBefore = 16

If Dir$(FilePath) = "" Then
    CheckCRC = "file not found."
    Exit Function
End If

'open the file
iFreeFile = FreeFile
Open FilePath For Binary Access Read As #iFreeFile

    'get CBbytes
    CBbytes = LOF(iFreeFile) - BytesBefore + 1

    'free memory space
    ReDim FileArr(LOF(iFreeFile)) As Byte

    'input byte array
    Do While Not EOF(iFreeFile)
        Get #iFreeFile, , FileArr()
    Loop
    
    'check that if all bytes are 00
    For lnCount = 0 To BytesBefore - 1
        If FileArr(CBbytes + lnCount) <> 0 Then
            NotZero = True
        End If
    Next lnCount
    
    If NotZero = False Then
        'no CRC found on file
        CheckCRC = 0
        Exit Function
    End If
    
    'store read CRC and zero all final variables
    For lnCount = 0 To BytesBefore - 1
        If FileArr(CBbytes + lnCount) <> 0 Then
            ReadCRC = ReadCRC & Chr(FileArr(CBbytes + lnCount))
        End If
        FileArr(CBbytes + lnCount) = 0
    Next lnCount
    
    'compute CRC
    Dim m_CRC As clsCRC
    'start CRC engine
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    MyCRC = Hex$(m_CRC.CalculateBytes(FileArr))
    Set m_CRC = Nothing

    'MsgBox "computed: '" & MyCRC & "'" & vbCrLf & "read: '" & ReadCRC & "'"

    If Trim(MyCRC) <> Trim(ReadCRC) Then
        'KO
        CheckCRC = 1
    Else
        'OK
        CheckCRC = 2
    End If


Close #iFreeFile

End Function


Function IsIntegrityOk() As Integer

'dim
Dim MyCRCres As Integer

MyCRCres = CheckCRC(App.Path & "\" & App.EXEName & ".exe")

If MyCRCres = 0 Then
    MsgBox "No CRC signature found."
ElseIf MyCRCres = 1 Then
    MsgBox "CRC differs: file has been patched!"
ElseIf MyCRCres = 2 Then
    MsgBox "File OK!"
End If

End Function


