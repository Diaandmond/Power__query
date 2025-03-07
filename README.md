# Power_query

## Dynamically expand all columns from file 

filter_page_only = Table.SelectRows(Source, each ([Kind] = "Page")),<br/>
#"Removed Other Columns" = Table.SelectColumns(filter_page_only,{"Data"}),<br/>
#"AddColumnWithColumnNames" = Table.AddColumn(#"Removed Other Columns", "ColumnNames", each if Type.Is(Value.Type([Data]), type table) then Table.ColumnNames([Data]) else {}),<br/>
#"AddColumnWithColumnNamesCount" = Table.AddColumn(#"AddColumnWithColumnNames", "ColumnNamesCount", each List.Count([ColumnNames])),<br/>
#"Added Index" = Table.AddIndexColumn(AddColumnWithColumnNamesCount, "Index", 1, 1, Int64.Type),<br/>
#"SortedTable" = Table.Sort(#"Added Index", {{"ColumnNamesCount", Order.Descending}}),<br/>
Index = SortedTable{0}[Index],<br/>
#"ColumnNames"= Table.ColumnNames(#"Removed Other Columns"[Data]{Index-1}),<br/>
#"Expanded Data" = Table.ExpandTableColumn(#"Removed Other Columns", "Data", ColumnNames)<br/>

## Get column names and search

get_column_names = Table.ColumnNames(#"Transposed Table3"),<br/>
replace_null = Table.ReplaceValue(#"Transposed Table3", null, "tuščia", Replacer.ReplaceValue, get_column_names),<br/>
replace_errors= Table.ReplaceErrorValues(replace_null, List.Transform(get_column_names, each {_, "tuščia"}))<br/>
change_all_columns_type = Table.TransformColumnTypes(replace_errors, List.Transform(Table.ColumnNames(replace_errors), each {_, type TYPE})),<br/>
search_ = Table.SelectRows( , each List.AnyTrue(List.Transform(Record.FieldValues(_), each Text.Contains(_, ""))))...<br/>

## Remove extra empty column spaces

add_column1_full = Table.AddColumn(add_column0_full, "column1[full]", each <br/>
  if [Data starts] = 1 then [Column1] else <br/>
  if [Data starts] = 2 then [Column2] else null),<br/>
add_column2_full = Table.AddColumn(add_column1_full,"column2[full]", each <br/>
  if [Data starts] = 1 then [Column2] else <br/>
  if [Data starts] = 2 then [Column3] else null)...<br/>

## Index only duplicate values

--Group by all rows, add custom, and expand only the index. <br/>
get_index = Table.AddIndexColumn([All Rows], "Index", 1, 1, Int64.Type)...<br/>







Source = Expression.Evaluate(Text.FromBinary(File.Contents("failas su žingsniais.txt")), #shared)

in Source

--------------------------------------
 Get table names and use it
--------------------------------------
Source = Excel.Workbook(Parameter1, null, true),
#"Filtered Rows" = Table.SelectRows(Source, each [Kind] = "Sheet"),
#"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Data"}),
#"AddColumnWithColumnNames" = Table.AddColumn(#"Removed Other Columns", "ColumnNames", each 
    if Type.Is(Value.Type([Data]), type table) then Table.ColumnNames([Data]) else {}),
#"AddColumnWithColumnNamesCount" = Table.AddColumn(#"AddColumnWithColumnNames", "ColumnNamesCount", each List.Count([ColumnNames])),
#"Added Index" = Table.AddIndexColumn(AddColumnWithColumnNamesCount, "Index", 1, 1, Int64.Type),
#"SortedTable" = Table.Sort(#"Added Index", {{"ColumnNamesCount", Order.Descending}}),
Index = SortedTable{0}[Index],
#"ColumnNames"= Table.ColumnNames(#"Removed Other Columns"[Data]{Index-1}),
#"Expanded Data" = Table.ExpandTableColumn(#"Removed Other Columns", "Data", ColumnNames),
#"search_" = Table.SelectRows(#"Changed Type", each List.AnyTrue(List.Transform(Record.FieldValues(_), each Text.Contains(_, "1.1"))))


  
--------------------------------------------
Column count use dinamicaly
-------------------------------------------  
#"ColumnCount" = Table.ColumnCount(#"Transposed Table1"),
#"ColumnNames_2" = List.Transform({1..#"ColumnCount"}, each "Column" & Text.From(_)),
#"Transformations" = List.Transform(#"ColumnNames_2", each {_, Text.Proper, type text}),
#"Capitalized Each Word" = Table.TransformColumns(#"Transposed Table1", #"Transformations")
      

--------------------------------------------
Change type
-------------------------------------------  

#"ColumnNames1" = Table.ColumnNames(#"Transposed Table1"),
#"typeTransformations" = List.Transform(ColumnNames1, each {_, type text}),
#"transformedTable" = Table.TransformColumnTypes(#"Transposed Table1", typeTransformations),


--jeigu reikia index'o tik paasikartojančiai informacijai--
--group by all rows, add custom, expand only index. 

= Table.AddIndexColumn([All Rows], "Index", 1, 1, Int64.Type)



SELECT table_name, column_name
FROM information_schema.columns
WHERE table_schema = 'infot_2_transorloja'
AND column_name LIKE '%reiso_nr%'

--------------------------------------------------
             Add missing columns
--------------------------------------------------
let
    // Original table
    Source = #"Added Custom",

    // Define the list of required columns
    RequiredColumns = {"Custom", "Index", "Origin", "Destination", "Vehicle type", "Price matrix", 
                       "Entry date", "Base step", "Validity from", "Validity to", 
                       "Base price per vehicle", "Price per vehicle", "Approx. volume"},

    // Check if each required column exists, and add empty ones if missing
    AddMissingColumns = List.Accumulate(
        RequiredColumns, 
        Source, 
        (state, column) =>
            if List.Contains(Table.ColumnNames(state), column) then
                state
            else
                Table.AddColumn(state, column, each null)
    ),

    // Reorder columns based on the desired order
    ReorderedTable = Table.ReorderColumns(AddMissingColumns, RequiredColumns)
in
    ReorderedTable

--------------------------------------------------
            VBA button to save files
--------------------------------------------------

    Sub SaveData()
    Dim wsSource As Worksheet
    Dim tbl As ListObject
    Dim rngFiltered As Range
    Dim savePath As String
    Dim criteriaColumnName As String
    Dim filterValues As Variant
    Dim filterValue As Variant
    Dim newWorkbook As Workbook
    Dim i As Integer
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Set the worksheet and table
    Set wsSource = ThisWorkbook.Sheets("Duomenys á Rivilæ")
    Set tbl = wsSource.ListObjects("Duomenys_á_Rivilæ")
    
    ' Get the base folder path from cell L1
    basePath = wsSource.Range("O1").Value
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    
    
   ' Set the save path
    savePath = basePath & "Pilni duomenys_importui.xlsx"
    
    ' Create a new workbook and copy the full data
    Set newWorkbook = Application.Workbooks.Add
    tbl.DataBodyRange.Copy Destination:=newWorkbook.Sheets(1).Range("A2")
    newWorkbook.Sheets(1).Name = "Pilni duomenys"
    
    ' Save the new workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    newWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    
    MsgBox "Duomenys iðsaugoti", vbInformation
End Sub

--------------------------------------------------
            VBA calendar
--------------------------------------------------

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    ' Check if Active Cell format matches the "Short Date" number format
    If ActiveCell.NumberFormat = "m/d/yyyy" Then
        ' If the Active Cell format is a date, make the calendar visible
        ActiveSheet.Shapes("Calendar").Visible = True
        
        ' Change the position of the calendar to be just below and to the right of the Active Cell
        ActiveSheet.Shapes("Calendar").Left = ActiveCell.Left + ActiveCell.Width
        ActiveSheet.Shapes("Calendar").Top = ActiveCell.Top + ActiveCell.Height
    
    ' If the Active Cell isn't a date, make the calendar invisible
    Else: ActiveSheet.Shapes("Calendar").Visible = False
    
    End If

End Sub


