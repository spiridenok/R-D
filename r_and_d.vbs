Const pp_file_path = "c:\Users\dspirydz\Documents\ola\2013 - P03.xlsx" 'eto glavnyj fail
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
                
    Dim pp_file As Workbook
    If Dir(pp_file_path) = "" Then
        MsgBox "P0x file is not found. Makes no sence to continue, stop executing macro..."
        'file_path = pick_file("Select Italy Vendors")
        Exit Sub
        If file_path <> "" Then
            Set vendors_file = Workbooks.Open(file_path)
        End If
    Else
        Set pp_file = Workbooks.Open(pp_file_path)
    End If
    
    Dim ws As Excel.Worksheet
    Application.Workbooks.Add
    Set ws = Excel.ActiveSheet
    Dim nxt As Worksheet
    Set nxt = ThisWorkbook.Worksheets(1)
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
    
    Dim pp As New Scripting.Dictionary
    For Each person In pp_file.Worksheets(1).Rows
        Dim tmp As String
        tmp = person.Cells(2) + " "
        If Not IsEmpty(person.Cells(3)) Then
            tmp = tmp + person.Cells(3) + " "
        End If
        tmp = tmp + person.Cells(1)
        If Trim(tmp) <> "" Then pp.Add tmp, CStr(person.Cells(4)) + "," + person.Cells(7) + "," + CStr(person.Cells(5))
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
                    int_orders.Add CStr(ord.Cells(2)), CStr(ord.Cells(1))
                End If
            End If
        Next ord
        int_orders_wb.Close
    End If
    
    Dim pers_nums As New Scripting.Dictionary
    If Dir(personeel_nummers_file_path) = "" Then
        MsgBox "File with personeelsnummers is not found! Company code will not be filled in!"
        Set pers_nums = Nothing
    Else
        Dim pers_nums_wb As Workbook
        Set pers_nums_wb = Workbooks.Open(personeel_nummers_file_path, 0)
        For Each pn In pers_nums_wb.Worksheets(1).Rows
            If IsEmpty(pn.Cells(1)) Then
                Exit For
            Else
                pers_nums.Add pn.Cells(4), pn.Cells(1)
            End If
        Next pn
        'pers_nums_wb.Close
    End If
    
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
    
    nxt.Range("b3", "h1000").ClearContents
    nxt.Range("b3", "h1000").Font.ColorIndex = 1
    
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
                    
                        pp_data = ""
                        If pp.Exists(current_name) Then 'And Not IsEmpty(pp(current_name)) Then
                            pp_data = pp(current_name)
                        Else 'no exact match, search in pieces
                            For Each pp_name In pp
                                If DelLeft(pp_name) = DelLeft(current_name) Then
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                ElseIf Split(pp_name, " ")(0) = Split(current_name, " ")(1) Then
                                    pp_data = pp(pp_name)
                                    current_name = pp_name
                                    nxt.Cells(i, TGT_NAME).Font.ColorIndex = 3
                                    Exit For
                                End If
                            Next pp_name
                        End If
                        ' Just find the first name that matches.
                        If pp_data = "" Then 'no exact match, search the closest name
                            current_name = strSimLookup(current_name, pp.Keys, 0)
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
                            If pers_nums Is Nothing Then
                                a = ""
                            Else
                                For Each k In pers_nums.Keys
                                    If k = CLng(Split(pp_data, ",")(2)) Then a = pers_nums(k)
                                Next k
                            End If
                        
                            nxt.Cells(i, TGT_ASS_HOURS) = projects(proj)
                            nxt.Cells(i, TGT_COST_CENTER) = CLng(Split(pp_data, ",")(0))
                            nxt.Cells(i, TGT_ACTIVITY_TYPE) = Split(pp_data, ",")(1)
                            nxt.Cells(i, TGT_PERS_NO) = CLng(Split(pp_data, ",")(2))
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

Sub SaveRawData()
    filesavename = Application.GetSaveAsFilename(fileFilter:="xlsx Files (*.xlsx), *.xlsx")
    
    If filesavename <> False Then
        Dim ThisWksht As Worksheet
        
        Set ThisWksht = ActiveSheet
        Set NewWkbk = Workbooks.Add

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
