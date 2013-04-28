Const p03_file_path = "c:\Users\dspirydz\Documents\ola\2013 - P03.xlsx"
Const internal_orders_file_path = "c:\Users\dspirydz\Documents\ola\Internal Orders in GmbH SAP.xlsx"
Const personeel_nummers_file_path = "c:\Users\dspirydz\Documents\ola\Personeelsnummers 2013 01 25.xlsx"

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
    
    Dim int_orders_wb As Workbook
    Set int_orders_wb = Workbooks.Open(internal_orders_file_path)
    
    Dim int_orders As New Scripting.Dictionary
    Dim start_found As Boolean
    start_found = False
    For Each ord In int_orders_wb.Worksheets(1).Rows
        If ord.Cells(1) = "Order number" Then start_found = True
        
        If start_found Then
            If IsEmpty(ord.Cells(1)) Then
                Exit For
            Else
                int_orders.Add CStr(ord.Cells(2)), CStr(ord.Cells(1))
            End If
        End If
    Next ord
    int_orders_wb.Close
    
    Dim pers_num_wb As Workbook
    Set pers_num_wb = Workbooks.Open(personeel_nummers_file_path, 0)
    
    Dim pers_nums As New Scripting.Dictionary
    For Each pn In pers_num_wb.Worksheets(1).Rows
        If IsEmpty(pn.Cells(1)) Then
            Exit For
        Else
            pers_nums.Add pn.Cells(4), pn.Cells(1)
        End If
    Next pn
    
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
                    tmp = Trim(Split(Split(proj, "-")(0), " ")(0))
                    If Not IsNumeric(tmp) Then
                        Dim split_name() As String
                        split_name = Split(current_name, " ")
                        nxt.Cells(i, 3 + 1) = split_name(UBound(split_name)) + ", "
                        For cnt = 0 To UBound(split_name) - 1
                            nxt.Cells(i, 3 + 1) = nxt.Cells(i, 3 + 1) + " " + split_name(cnt)
                        Next
                        If int_orders.Exists(tmp) Then
                            nxt.Cells(i, 5 + 1) = int_orders(tmp)
                        Else
                            nxt.Cells(i, 5 + 1) = tmp
                        End If
                    
                        For Each k In pers_nums.Keys
                            If Not IsEmpty(p03(current_name)) Then
                                If k = CLng(Split(p03(current_name), ",")(2)) Then a = pers_nums(k)
                            End If
                        Next k
                    
                        nxt.Cells(i, 6 + 1) = projects(proj)
                        If p03.Exists(current_name) And Not IsEmpty(p03(current_name)) Then
                            nxt.Cells(i, 1 + 1) = CLng(Split(p03(current_name), ",")(0))
                            nxt.Cells(i, 4 + 1) = Split(p03(current_name), ",")(1)
                            nxt.Cells(i, 3) = CLng(Split(p03(current_name), ",")(2))
                            k = Split(p03(current_name), ",")(2)
                            nxt.Cells(i, 1) = a
                        End If
                        
                        i = i + 1
                    End If
                Next proj
                Set projects = Nothing
            End If
            current_name = rw.Cells(NAME)
        End If
        
        If current_name <> "" Then
'            If Not IsEmpty(rw.Cells(WK_NUM)) And `(ws.Rows.Cells(rw.Row + 1, WK_NUM)) And IsEmpty(ws.Rows.Cells(rw.Row + 1, NAME)) Then
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
            
    ActiveWindow.Close savechanges:=False 'close pers numbers
    ActiveWindow.Close savechanges:=False 'close book*
    
            
End Sub
