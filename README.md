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

