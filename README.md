# Power_query

## Dynamically expand all columns from file 

filter_page_only = Table.SelectRows(Source, each ([Kind] = "Page")),
get_column_names= Table.ColumnNames(filter_page_only [Data]{0}),
expand_all_columns = Table.ExpandTableColumn(filter_page_only, "Data", get_column_names)...

## Get column names and search

#"ColumnNames" = Table.ColumnNames(#"Transposed Table3"),
#"ReplacedValue" = Table.ReplaceValue(#"Transposed Table3", null, "tuščia", Replacer.ReplaceValue, ColumnNames),
#"ReplacedError"= Table.ReplaceErrorValues(#"ReplacedValue", List.Transform(#"ColumnNames", each {_, "tuščia"}))
#"Changed Type" = Table.TransformColumnTypes(#"ReplacedError", List.Transform(Table.ColumnNames(#"ReplacedError"), each {_, type text})),
#"search_" = Table.SelectRows(#"Changed Type", each List.AnyTrue(List.Transform(Record.FieldValues(_), each Text.Contains(_, "1.1")))),
