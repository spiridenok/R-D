Sub r_d()
    Dim columnFormats(0 To 255) As Integer
    For i = 0 To 255
        columnFormats(i) = xlTextFormat
    Next i

    Dim filename As Variant
    filename = Application.GetOpenFilename("CVS files (*.csv),*.csv", 1, "Open", "", False)
    ' If user clicks Cancel, stop.
    If (filename = False) Then
        Exit Sub
    End If
        
    Dim ws As Excel.Worksheet
    Dim nxt As Worksheet
    Application.Workbooks.Add
    Set ws = Excel.ActiveSheet
    Set nxt = ThisWorkbook.Worksheets(2)
    Application.DisplayAlerts = False
    Sheets("Sheet3").Delete
    Application.DisplayAlerts = True

    With ws.QueryTables.Add("TEXT;" & filename, ws.Cells(1, 1))
        .FieldNames = True
        .AdjustColumnWidth = True
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileCommaDelimiter = True
        ''// This array will need as many entries as there will be columns:
        .TextFileColumnDataTypes = columnFormats
        .Refresh
    End With
    
    Dim p03_path As String
    p03_file_path = "c:\Users\dspirydz\Documents\ola\2013 - P03.xlsx"
    
    Dim p03_file As Workbook
    If Dir(p03_file_path) = "" Then
        MsgBox "File with Italy vendors is not found, please select a file (press 'Cancel' in the next file open dialog to continue without vendors list)"
        'file_path = pick_file("Select Italy Vendors")

        If file_path <> "" Then
            Set vendors_file = Workbooks.Open(file_path)
        End If
    Else
        Set p03_file = Workbooks.Open(p03_file_path)
    End If
    
    Dim p03 As New Scripting.Dictionary
    For Each person In p03_file.Worksheets(1).Rows
        Dim tmp As String
        tmp = person.Cells(2) + " "
        If Not IsEmpty(person.Cells(3)) Then
            tmp = tmp + person.Cells(3) + " "
        End If
        tmp = tmp + person.Cells(1)
        If Trim(tmp) <> "" Then p03.Add tmp, CStr(person.Cells(4)) + "," + person.Cells(7) + "," + CStr(person.Cells(5))
        If person.Row > 1000 Then Exit For
    Next person
    p03_file.Close
    
    Const NAME = 3
    Const WK_NUM = 4
    Const TOTAL_HOURS = 5
    Const PROJ_NUM = 5
    Const PROJ_DESC = 6
    Const PROJ_HOURS = 7
    
    i = 2
    Dim current_name As String
    Dim wk_n As String
    current_name = ""

    Dim names As New Scripting.Dictionary
    Dim projects As New Scripting.Dictionary
    
    Dim cnt As Integer
    
    For Each rw In ws.Rows
        If Not IsEmpty(rw.Cells(NAME)) And Not IsNumeric(rw.Cells(NAME)) Then
            If current_name <> "" Then
                For Each proj In projects.Keys
                    Dim split_name() As String
                    split_name = Split(current_name, " ")
                    nxt.Cells(i, 3) = split_name(UBound(split_name))
                    For cnt = 0 To UBound(split_name) - 1
                        nxt.Cells(i, 3) = nxt.Cells(i, 3) + ", " + split_name(ctr)
                    Next
                    
                    nxt.Cells(i, 5) = proj
                    nxt.Cells(i, 6) = projects(proj)
                    If p03.Exists(current_name) Then
                        nxt.Cells(i, 1) = CLng(Split(p03(current_name), ",")(0))
                        nxt.Cells(i, 4) = Split(p03(current_name), ",")(1)
                        nxt.Cells(i, 2) = CLng(Split(p03(current_name), ",")(2))
                    End If
                    
                    i = i + 1
                Next proj
                Set projects = Nothing
            End If
            current_name = rw.Cells(NAME)
        End If
        
        If current_name <> "" Then
'            If Not IsEmpty(rw.Cells(WK_NUM)) And IsEmpty(ws.Rows.Cells(rw.Row + 1, WK_NUM)) And IsEmpty(ws.Rows.Cells(rw.Row + 1, NAME)) Then
'                wk_n = rw.Cells(WK_NUM)
'                nxt.Cells(i, 2) = wk_n
'                i = i + 1
'            Else
'                If Not IsEmpty(rw.Cells(PROJ_NUM)) And Not IsEmpty(rw.Cells(PROJ_HOURS)) Then
'                End If
'            End If
            If Not IsEmpty(rw.Cells(PROJ_NUM)) And Not IsEmpty(rw.Cells(PROJ_HOURS)) Then
                Dim h As Integer
                'Dim proj As String
                
                h = CInt(rw.Cells(PROJ_HOURS))
                proj = rw.Cells(PROJ_NUM)
                If Not projects.Exists(proj) Then
                    projects.Add proj, h
                Else
                    projects(proj) = projects(proj) + h
                End If
            End If
        End If
        
        'just to limit the max amount of records to speed up the process
        If rw.Row > 10000 Then Exit For
    Next rw
    
End Sub
