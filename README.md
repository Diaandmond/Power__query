# Power_query

## Dynamically expand all columns from file 

filter_page_only = Table.SelectRows(Source, each ([Kind] = "Page")),<br/>
get_column_names= Table.ColumnNames(filter_page_only [Data]{0}),<br/>
expand_all_columns = Table.ExpandTableColumn(filter_page_only, "Data", get_column_names)...<br/>

## Get column names and search

get_column_names = Table.ColumnNames(#"Transposed Table3"),<br/>
replace_null = Table.ReplaceValue(#"Transposed Table3", null, "tuščia", Replacer.ReplaceValue, get_column_names),<br/>
replace_errors= Table.ReplaceErrorValues(replace_null, List.Transform(get_column_names, each {_, "tuščia"}))<br/>
change_all_columns_type = Table.TransformColumnTypes(replace_errors, List.Transform(Table.ColumnNames(replace_errors), each {_, type TYPE})),<br/>
search_ = Table.SelectRows(change_all_columns_type , each List.AnyTrue(List.Transform(Record.FieldValues(_), each Text.Contains(_, "SEARCH TEXT"))))...<br/>

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

