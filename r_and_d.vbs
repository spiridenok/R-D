Const WK_NUM = 4
Public Const TGT_NAME = 5
Public Const TGT_PROJECT_NO = 7
Public Const TGT_ASS_HOURS = 8

Sub r_d()
    
    If get_first_name(Application.UserName) = "Dzmitry" Then
        pp_file_path = "c:\Users\dspirydz\Documents\ola\2013 - P12.xlsx" 'eto glavnyj fail
        internal_orders_file_path = "c:\Users\dspirydz\Documents\ola\Internal Orders in GmbH SAP.xlsx"
        'personeel_nummers_file_path = "c:\Users\dspirydz\Documents\ola\Personeelsnummers 2013 01 25.xlsx"
    Else
        pp_file_path = "K:\Finance SiTel\Accounting & Control\2013\Accounting cycle\R&D Accounting\Employees in R&D\2013 - P04.xlsx"
        internal_orders_file_path = "K:\Finance SiTel\Accounting & Control\2013\Accounting cycle\Internal Orders\Internal Orders in GmbH SAP.xlsx"
        'personeel_nummers_file_path = "K:\Finance SiTel\Accounting & Control\2013\Accounting cycle\R&D Accounting\Definitieve R&D\P01\Personeelsnummers 2013 01 25.xlsx"
    End If
    
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
                
    Dim pp_file As Workbook
    If Dir(pp_file_path) = "" Then
        MsgBox "Pxx file is not found. Makes no sence to continue, stop executing macro..."
        Exit Sub
        If file_path <> "" Then
            Set vendors_file = Workbooks.Open(file_path)
        End If
    Else
        Set pp_file = Workbooks.Open(pp_file_path)
    End If
    
    Dim ws As Excel.Worksheet
    Application.Workbooks.add
    Set ws = Excel.ActiveSheet
    Dim nxt As Worksheet
    Set nxt = ThisWorkbook.Worksheets(1)
    Application.DisplayAlerts = False
    Sheets("Sheet3").Delete
    Application.DisplayAlerts = True

    With ws.QueryTables.add("TEXT;" & filename, ws.Cells(1, 1))
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
    
    Dim pp As New Scripting.Dictionary
    Dim pers_nums As New Scripting.Dictionary
    For Each person In pp_file.Worksheets(1).Rows
        Dim tmp As String
        tmp = person.Cells(2) + " "
        If Not IsEmpty(person.Cells(3)) Then
            tmp = tmp + person.Cells(3) + " "
        End If
        tmp = tmp + person.Cells(1)
        ' Add cost center, activity type (role) and employee number (in this particular order)
        ' TODO: this should be a new object of my own type
        If Trim(tmp) <> "" Then pp.add tmp, CStr(person.Cells(4)) + "," + person.Cells(7) + "," + CStr(person.Cells(5))
        If Trim(tmp) <> "" And IsNumeric(person.Cells(5)) Then pers_nums.add CStr(person.Cells(5)), CStr(person.Cells(9))
        If person.Row > 1000 Then Exit For
    Next person
    pp_file.Close
    
    Dim int_orders As New Scripting.Dictionary
    If Dir(internal_orders_file_path) = "" Then
        MsgBox "File with internal order numbers is not found! Original orders will not be converted!"
        Set int_orders = Nothing
    Else
        Dim int_orders_wb As Workbook
        Set int_orders_wb = Workbooks.Open(internal_orders_file_path, 0)
        Dim start_found As Boolean
        start_found = False
        For Each ord In int_orders_wb.Worksheets(1).Rows
            If ord.Cells(1) = "Order number" Then start_found = True
            If start_found Then
                If IsEmpty(ord.Cells(1)) Then
                    Exit For
                Else
                    int_orders.add CStr(ord.Cells(2)), CStr(ord.Cells(1))
                End If
            End If
        Next ord
        int_orders_wb.Close
    End If
    
'    If Dir(personeel_nummers_file_path) = "" Then
'        MsgBox "File with personeelsnummers is not found! Company code will not be filled in!"
'        Set pers_nums = Nothing
'    Else
'        Dim pers_nums_wb As Workbook
'        Set pers_nums_wb = Workbooks.Open(personeel_nummers_file_path, 0)
'        For Each pn In pers_nums_wb.Worksheets(1).Rows
'            If IsEmpty(pn.Cells(1)) Then
'                Exit For
'            Else
'                pers_nums.add pn.Cells(4), pn.Cells(1)
'            End If
'        Next pn
'    End If
    
    Const YEAR = 1
    Const COST_CENTER = 2
    Const NAME = 3
    Const TOTAL_HOURS = 5
    Const PROJ_NUM = 5
    Const PROJ_DESC = 6
    Const PROJ_HOURS = 7
    
    ' 1st row is empty, 2nd contains the header
    i = 3
    Dim current_name As String
    Dim wk_n As String
    current_name = ""

    Dim names As New Scripting.Dictionary
    Dim projects As New Scripting.Dictionary
    
    Const TGT_COMP_CODE = 2
    Const TGT_COST_CENTER = 3
    Const TGT_PERS_NO = 4
    Const TGT_ACTIVITY_TYPE = 6
    
    nxt.Range("b3", "h1000").ClearContents
    nxt.Range("b3", "h1000").Font.ColorIndex = 1
    
    Dim conv As ConversionsSheet
    Set conv = New ConversionsSheet
    Dim right_cost_center
    right_cost_center = False
    
    WeekSelectionDialog.FromWeekNum.SetFocus
    WeekSelectionDialog.Show
    
    If WeekSelectionDialog.FromWeekNum.Value = "" Or WeekSelectionDialog.ToWeekNum.Value = "" Then GoTo CLEAN_UP
    
    from_week = CInt(WeekSelectionDialog.FromWeekNum.Value)
    to_week = CInt(WeekSelectionDialog.ToWeekNum.Value)
    
    If from_week = 0 Or to_week = 0 Then GoTo CLEAN_UP
    
    For Each rw In ws.Rows
        If Not IsEmpty(rw.Cells(COST_CENTER)) Then
            If current_name <> "" Then
                For Each proj In projects.Keys
                    tmp = Trim(Split(Split(proj, "-")(0), " ")(0))
                    Dim project_found As Boolean
                    project_found = False
                    If Not IsNumeric(tmp) Then
                        If int_orders.Exists(tmp) Then
                            nxt.Cells(i, TGT_PROJECT_NO) = int_orders(tmp)
                            project_found = True
                        Else
                            'project is not found directly, try to find as a substring
                            nxt.Cells(i, TGT_PROJECT_NO) = tmp
                            For Each k In int_orders
                                If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                    nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                    project_found = True
                                    Exit For
                                Else
                                    tmp = Replace(tmp, "DATA", "")
                                    tmp = Replace(tmp, "DECT", "")
                                    If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                        nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                        project_found = True
                                        Exit For
                                    End If
                                End If
                            Next k
                        End If
                        
                        If Not project_found Then nxt.Cells(i, TGT_PROJECT_NO).Font.ColorIndex = 3
                    
                        pp_data = ""
                        If pp.Exists(current_name) Then 'And Not IsEmpty(pp(current_name)) Then
                            pp_data = pp(current_name)
                        Else 'no exact match, search in pieces
                            For Each pp_name In pp
                                If get_last_name(pp_name) = get_last_name(current_name) Then
                                    conv.add current_name, pp_name
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                ElseIf Split(pp_name, " ")(0) = Split(current_name, " ")(1) Then
                                    conv.add current_name, pp_name
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                End If
                            Next pp_name
                        End If
                        ' Just find the first name that matches.
                        If pp_data = "" Then 'no exact match, search the closest name
                            Dim org_name As String
                            org_name = current_name
                            current_name = strSimLookup(current_name, pp.Keys, 0)
                            conv.add org_name, current_name
                            pp_data = pp(current_name)
                            nxt.Cells(i, TGT_NAME).Font.ColorIndex = 4
                        End If
                        
                        Dim split_name() As String
                        split_name = Split(current_name, " ")
                        nxt.Cells(i, TGT_NAME) = split_name(UBound(split_name)) + ", "
                        For cnt = 0 To UBound(split_name) - 1
                            nxt.Cells(i, TGT_NAME) = nxt.Cells(i, TGT_NAME) + " " + split_name(cnt)
                        Next
                        
                        If Len(pp_data) > 0 Then
                        
                            nxt.Cells(i, TGT_ASS_HOURS) = projects(proj)
                            nxt.Cells(i, TGT_COST_CENTER) = CLng(Split(pp_data, ",")(0))
                            nxt.Cells(i, TGT_ACTIVITY_TYPE) = Split(pp_data, ",")(1)
                            nxt.Cells(i, TGT_PERS_NO) = CLng(Split(pp_data, ",")(2))
                            If pers_nums.Exists(Split(pp_data, ",")(2)) Then nxt.Cells(i, TGT_COMP_CODE) = pers_nums(Split(pp_data, ",")(2))
                        End If
                        
                        i = i + 1
                    End If
                Next proj
            End If
            current_name = ""
            cost_center_num = Split(rw.Cells(COST_CENTER), ":")
            If UBound(cost_center_num) > 0 Then
                cost_center_num = Trim(cost_center_num(1))
                ' empty cost center is a workaround for employees that are added later
                If cost_center_num = "" Or cost_center_num = "310001" Or cost_center_num = "320001" Or cost_center_num = "320002" Then
'                If cost_center_num = "310001" Or cost_center_num = "320001" Or cost_center_num = "320002" Then
                    right_cost_center = True
                Else
                    right_cost_center = False
                End If
            End If
        ElseIf Not IsEmpty(rw.Cells(NAME)) And Not IsNumeric(rw.Cells(NAME)) And right_cost_center Then
            If current_name <> "" Then
                For Each proj In projects.Keys
                    tmp = Trim(Split(Split(proj, "-")(0), " ")(0))
                    project_found = False
                    If Not IsNumeric(tmp) Then
                        If int_orders.Exists(tmp) Then
                            nxt.Cells(i, TGT_PROJECT_NO) = int_orders(tmp)
                            project_found = True
                        Else
                            'project is not found directly, try to find as a substring
                            nxt.Cells(i, TGT_PROJECT_NO) = tmp
                            For Each k In int_orders
                                If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                    nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                    project_found = True
                                    Exit For
                                Else
                                    tmp = Replace(tmp, "DATA", "")
                                    tmp = Replace(tmp, "DECT", "")
                                    If InStr(1, k, tmp, vbTextCompare) > 0 Then
                                        nxt.Cells(i, TGT_PROJECT_NO) = int_orders(k)
                                        project_found = True
                                        Exit For
                                    End If
                                End If
                            Next k
                        End If
                        
                        If Not project_found Then nxt.Cells(i, TGT_PROJECT_NO).Font.ColorIndex = 3
                    
                        pp_data = ""
                        If pp.Exists(current_name) Then 'And Not IsEmpty(pp(current_name)) Then
                            pp_data = pp(current_name)
                        Else 'no exact match, search in pieces
                            For Each pp_name In pp
                                If get_last_name(pp_name) = get_last_name(current_name) Then
                                    conv.add current_name, pp_name
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                ElseIf Split(pp_name, " ")(0) = Split(current_name, " ")(1) Then
                                    conv.add current_name, pp_name
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                End If
                            Next pp_name
                        End If
                        ' Just find the first name that matches.
                        If pp_data = "" Then 'no exact match, search the closest name
                            'Dim org_name As String
                            org_name = current_name
                            current_name = strSimLookup(current_name, pp.Keys, 0)
                            conv.add org_name, current_name
                            pp_data = pp(current_name)
                            nxt.Cells(i, TGT_NAME).Font.ColorIndex = 4
                        End If
                        
                        'Dim split_name() As String
                        split_name = Split(current_name, " ")
                        nxt.Cells(i, TGT_NAME) = split_name(UBound(split_name)) + ","
                        For cnt = 0 To UBound(split_name) - 1
                            nxt.Cells(i, TGT_NAME) = nxt.Cells(i, TGT_NAME) + " " + split_name(cnt)
                        Next
                        
                        If Len(pp_data) > 0 Then
                            nxt.Cells(i, TGT_ASS_HOURS) = projects(proj)
                            nxt.Cells(i, TGT_COST_CENTER) = CLng(Split(pp_data, ",")(0))
                            nxt.Cells(i, TGT_ACTIVITY_TYPE) = Split(pp_data, ",")(1)
                            nxt.Cells(i, TGT_PERS_NO) = CLng(Split(pp_data, ",")(2))
                            If pers_nums.Exists(Split(pp_data, ",")(2)) Then nxt.Cells(i, TGT_COMP_CODE) = pers_nums(Split(pp_data, ",")(2))
                        End If
                        
                        i = i + 1
                    End If
                Next proj
            End If
            Set projects = Nothing
            current_name = rw.Cells(NAME)
        End If
        
        If current_name <> "" Then
            If Not IsEmpty(rw.Cells(PROJ_NUM)) And Not IsEmpty(rw.Cells(PROJ_HOURS)) Then
                week = get_week_num(ws, rw.Row)
                If week >= from_week And week <= to_week Then
                    hours = CInt(rw.Cells(PROJ_HOURS))
                    proj = rw.Cells(PROJ_NUM)
                    If Not projects.Exists(proj) Then
                        projects.add proj, hours
                    Else
                        projects(proj) = projects(proj) + hours
                    End If
                End If
            End If
        End If
        
        'just to limit the max amount of records to speed up the process
        If rw.Cells(YEAR) = "2012" Then Exit For
    Next rw
            
CLEAN_UP:
    ActiveWindow.Close savechanges:=False 'close pers numbers
    'ActiveWindow.Close savechanges:=False 'close book*
            
End Sub

Function get_week_num(ws As Worksheet, row_index As Integer) As Integer
    wk_num_cell = ws.Cells(row_index, WK_NUM)
    If Not IsEmpty(wk_num_cell) And Not InStr(1, wk_num_cell, "week") = 0 Then
        get_week_num = CInt(Replace(wk_num_cell, "week: ", ""))
    Else
        get_week_num = get_week_num(ws, row_index - 1)
    End If
End Function

Function get_last_name(ByVal str As String) As String
 Dim d As Long
 d = InStr(1, str, " ")
 If Not d = 0 Then
    get_last_name = Right(str, Len(str) - d)
 Else
    get_last_name = str
End If
 End Function
Function get_first_name(ByVal str As String) As String
 If Not InStr(1, str, " ") Then
    get_first_name = Left(str, InStr(1, str, " ") - 1)
 Else
    get_first_name = str
 End If
 End Function

Sub SaveRawData()
    filesavename = Application.GetSaveAsFilename(fileFilter:="xlsx Files (*.xlsx), *.xlsx")
    
    If filesavename <> False Then
        Dim ThisWksht As Worksheet
        
        Set ThisWksht = ActiveSheet
        Set NewWkbk = Workbooks.add
        
        ThisWksht.Range("A1:H1000").Copy NewWkbk.Sheets(1).Range("A1")
        NewWkbk.Sheets(1).Range("A1:H1000").Select
        Selection.Columns.AutoFit
        Cells(1, 1).Select
        
        NewWkbk.SaveAs filename:=filesavename, FileFormat:=xlOpenXMLWorkbook
        If Err.Number = 1004 Then
            NewWkbk.Close
            MsgBox "File Name Not Valid " & filesavename & " Nothing is saved. Exiting."
        End If
    End If
 End Sub
