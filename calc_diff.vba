Const NAME_IDX = 5
Sub write_diff(ByRef sht As Sheet3, ByRef org_r() As Range, ByRef new_r() As Range)
    Dim diff() As Range
    ReDim diff(1 To 1) As Range
    
    For o_r = 1 To UBound(org_r)
        Dim project_found As Boolean
        project_found = False
        For n_r = 1 To UBound(new_r)
            If new_r(n_r).Cells(7) = org_r(o_r).Cells(7) Then
                If new_r(n_r).Cells(8) <> org_r(o_r).Cells(8) Then
                    org_r(o_r).Copy
                    Set last_row = sht.Cells(sht.rows.Count, 2).End(xlUp).Offset(1, -1)
                    sht.Paste last_row
                    last_row.Cells(1, 8) = CInt(new_r(n_r).Cells(8)) - CInt(org_r(o_r).Cells(8))
                End If
                project_found = True
            End If
        Next n_r
        ' Project is not found at all in the new data - mark all project hours as removed
        If Not project_found Then
            org_r(o_r).Copy
            Set last_row = sht.Cells(sht.rows.Count, 2).End(xlUp).Offset(1, -1)
            sht.Paste last_row
            last_row.Cells(1, 8) = "-" & last_row.Cells(1, 8) '.Value.Insert 1, "-"
        End If
    Next o_r
    
    ' Now find project in the new data which are not present in the old data
    For n_r = 1 To UBound(new_r)
        project_found = False
        For o_r = 1 To UBound(org_r)
            If new_r(n_r).Cells(7) = org_r(o_r).Cells(7) Then project_found = True
        Next o_r
        If Not project_found Then
            new_r(n_r).Copy
            Set last_row = sht.Cells(sht.rows.Count, 2).End(xlUp).Offset(1, -1)
            sht.Paste last_row
        End If
    Next n_r
End Sub

Function find_projects(name As String, ByRef r As Range) As Range()
    Dim proj_rows() As Range
    ReDim proj_rows(1 To 1) As Range
    Dim rw As Range
    
    For Each rw In r.rows
        If name = rw.Cells(NAME_IDX) Then
            Set proj_rows(UBound(proj_rows)) = rw
            ReDim Preserve proj_rows(1 To UBound(proj_rows) + 1) As Range
        End If
    Next rw
    If UBound(proj_rows) > 1 Then
        ReDim Preserve proj_rows(1 To UBound(proj_rows) - 1) As Range
    End If
    find_projects = proj_rows
End Function

Sub calc_delta()
    Dim filename As Variant
    filename = Application.GetOpenFilename("XLSM files (*.xlsm),*.xlsm", 1, "Open latest processed R&D data", "", False)
    ' If user clicks Cancel, stop.
    If filename = False Then
        Exit Sub
    End If
    
    Set last_row = Sheets("R&D").Cells(Sheet1.rows.Count, 2).End(xlUp)
    Dim sheet_with_diff As Sheet3
    
    Set sheet_with_diff = Sheets("Diff")
    sheet_with_diff.Range("b3:H2000").Value = ""
'    Set all_current_entries = Sheets("R&D").Cells("A1", "H" & last_row.Row)
    Dim org_entries As Range
    Set org_entries = Sheets("R&D").Range("A1", "H" & last_row.Row)
        
    ' 2nd parameter == 0 suppresses the "Update links" message
    Set f = Workbooks.Open(filename, 0)
    Set ws = Nothing
    On Error Resume Next
    Set ws = f.Sheets("R&D")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Could not find R&D tab in the file. The file to open must be an already processed file. Stopping now..."
        f.Close savechanges:=False
        Exit Sub
    End If
        
    Set last_row = ws.Cells(Sheet1.rows.Count, 2).End(xlUp)
    Dim new_entries As Range
    Set new_entries = ws.Range("A1", "H" & last_row.Row)
    
    Dim name As String
    
    Const PROJ_NUM = 7
    Const PROJ_HOURS = 8
    
    name = ""
    
    Dim org_projects() As Range
    Dim new_projects() As Range
    
    For Each rw In org_entries.rows
        If Not IsEmpty(rw.Cells(NAME_IDX)) And rw.Cells(NAME_IDX) <> "Name" Then
            If name <> rw.Cells(NAME_IDX) Then
                name = rw.Cells(NAME_IDX)
                org_projects = find_projects(name, org_entries)
                new_projects = find_projects(name, new_entries)
                write_diff sheet_with_diff, org_projects, new_projects
            End If
        End If
    Next rw
    
    ActiveWindow.Close savechanges:=False
    
End Sub
