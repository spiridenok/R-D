Const p03_file_path = "c:\Users\dspirydz\Documents\ola\2013 - P03.xlsx" 'eto glavnyj fail
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
    
    i = 3
    Dim current_name As String
    Dim wk_n As String
    current_name = ""

    Dim names As New Scripting.Dictionary
    Dim projects As New Scripting.Dictionary
    
    Dim cnt As Integer
    
    Const TGT_COMP_CODE = 2
    Const TGT_COST_CENTER = 3
    Const TGT_PERS_NO = 4
    Const TGT_NAME = 5
    Const TGT_ACTIVITY_TYPE = 6
    Const TGT_PROJECT_NO = 7
    Const TGT_ASS_HOURS = 8
    
    For Each rw In ws.Rows
        If Not IsEmpty(rw.Cells(NAME)) And Not IsNumeric(rw.Cells(NAME)) Then
            If current_name <> "" Then
                For Each proj In projects.Keys
                    tmp = Trim(Split(Split(proj, "-")(0), " ")(0))
                    If Not IsNumeric(tmp) Then
                        If int_orders.Exists(tmp) Then
                            nxt.Cells(i, TGT_PROJECT_NO) = int_orders(tmp)
                        Else
                            'project is not found directly, try to find as a substring
                            nxt.Cells(i, TGT_PROJECT_NO) = tmp
                            For Each k In int_orders
                                If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                    nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                    Exit For
                                Else
                                    tmp = Replace(tmp, "DATA", "")
                                    tmp = Replace(tmp, "DECT", "")
                                    If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                        nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                        Exit For
                                    End If
                                End If
                            Next k
                        End If
                    
                        p03_data = ""
                        If p03.Exists(current_name) And Not IsEmpty(p03(current_name)) Then
                            p03_data = p03(current_name)
                        Else 'no exact match, search in pieces
                            For Each p03_name In p03
                                If DelLeft(p03_name) = DelLeft(current_name) Then
                                    p03_data = p03(p03_name)
                                    current_name = p03_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                ElseIf Split(p03_name, " ")(0) = Split(current_name, " ")(1) Then
                                    p03_data = p03(p03_name)
                                    current_name = p03_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                End If
                            Next p03_name
                        End If
                        ' Just find the first name that matches. For some misterious reason
                        ' the names in the p03 list contains also the names from the timesheet...
'                        If p03_data = "" Then 'no exact match, search in pieces
'                            For Each p03_name In p03
'                                If Split(p03_name, " ")(0) = Split(current_name, " ")(0) Then
'                                    p03_data = p03(p03_name)
'                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 4
'                                    Exit For
'                                End If
'                            Next p03_name
'                        End If
                        
                        Dim split_name() As String
                        split_name = Split(current_name, " ")
                        nxt.Cells(i, TGT_NAME) = split_name(UBound(split_name)) + ", "
                        For cnt = 0 To UBound(split_name) - 1
                            nxt.Cells(i, TGT_NAME) = nxt.Cells(i, TGT_NAME) + " " + split_name(cnt)
                        Next
                        
                        If Len(p03_data) > 0 Then 'p03_data Is Not Nothing Then
                            For Each k In pers_nums.Keys
                                If k = CLng(Split(p03_data, ",")(2)) Then a = pers_nums(k)
                            Next k
                        
                            nxt.Cells(i, TGT_ASS_HOURS) = projects(proj)
                            nxt.Cells(i, TGT_COST_CENTER) = CLng(Split(p03_data, ",")(0))
                            nxt.Cells(i, TGT_ACTIVITY_TYPE) = Split(p03_data, ",")(1)
                            nxt.Cells(i, TGT_PERS_NO) = CLng(Split(p03_data, ",")(2))
                            nxt.Cells(i, TGT_COMP_CODE) = a
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

Function DelLeft(ByVal str As String) As String
 Dim l As Long, d As Long
 l = Len(str)
 d = InStr(1, str, " ")
 If Not d = 0 Then
    DelLeft = Right(str, l - d)
 Else
    DelLeft = str
End If
 End Function
