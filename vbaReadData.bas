Attribute VB_Name = "readData_"
Option Explicit

'Returns text file contents as a string
'--------------------------------------
Function readTxt( _
                textFilePath As String, _
                Optional multipleFiles As Boolean = False _
                ) As String
    If multipleFiles = False Then
        'Close any open text files
        Close
    End If
    
    'Get the number of the next free text file
    Dim fileNumber As Integer 'next free file number to open
    fileNumber = FreeFile
    
    'Read file contents as a string
    Dim txt As String   'file contents
    Open textFilePath For Input As #fileNumber
    txt = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    
    readTxt = txt
End Function

'Returns length of an array
'--------------------------
Function length(arr As Variant)
    length = UBound(arr) - LBound(arr) + 1
End Function


'Returns values of a CSV in arrays of arrays
'-------------------------------------------
'(rowIndex, columnIndex)
Function parseCSV( _
                str As String, _
                delimit As Variant _
                ) As Variant
    'Clean the string input
        str = normalizeNewLines(str, "\n")
        'Remove leading and trailing linebreaks
        str = trimStr(str, "\n")
        'Remove empty lines
        str = singleStr(str, "\n")
        
    'Handle multiple delimiters
        If IsArray(delimit) Then
            Dim d As String
            d = delimit(LBound(delimit))
            Dim i As Integer
            For i = LBound(delimit) To UBound(delimit)
                str = Replace(str, delimit(i), d)
            Next i
            delimit = d
        End If
    
    'Interpret input string as CSV
        Dim lines As Variant
        lines = Split(str, "\n")
        
        Dim result As Variant
        ReDim result(UBound(lines) - LBound(lines)) As Variant
        
        For i = LBound(lines) To UBound(lines)
            result(i) = Split(lines(i), delimit)
        Next i
    
    parseCSV = result
End Function

'Returns values from a comma separated file in an array matrix
'-------------------------------------------------------------
'(rowIndex, columnIndex)
Function readCSV(csvFilePath As String, _
                Optional delimit As String = "," _
                ) As Variant
    Dim str As String
    str = readTxt(csvFilePath)

    readCSV = parseCSV(str, delimit)
End Function

'Returns values from a tab separated file in an array matrix
'-----------------------------------------------------------
'(rowIndex, columnIndex)
Function readTSV(tsvFilePath As String) As Variant
    readTSV = readCSV(tsvFilePath, vbTab)
End Function


'Returns an array of collections
'-------------------------------
'Rows is a 2d array where the
'first row is an array of column
'names for all the following arrays
Function namedRows(rows As Variant) As Variant
    Dim result As Variant
    ReDim result(length(rows) - 1)
    
    Dim headerRow As Variant
    headerRow = rows(LBound(rows))
    
    Dim i As Integer
    Dim j As Integer
    For i = LBound(rows) + 1 To UBound(rows)    'for each row
        Set result(i) = New Collection
        For j = LBound(rows(i)) To UBound(rows(i))  'for each column
            result(i).Add rows(i)(j), headerRow(j)
        Next j
    Next i
    
    namedRows = result
End Function

