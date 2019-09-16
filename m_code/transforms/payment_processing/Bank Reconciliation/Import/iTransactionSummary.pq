let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "TRNSCTN_SMMRY")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "Transaction-Summary") and Text.Contains([Name], "2018")),
    #"Add [Import Month]" = Table.AddColumn(#"Filter [Name]", "Import Month", each Text.BetweenDelimiters([Name], "-", "_", 0, 0), type text),
    #"Filter [Import Month]" = Table.SelectRows(#"Add [Import Month]", each [Import Month] = import_month),
    #"Invoke Custom Function" = Table.AddColumn(#"Filter [Import Month]", "get_csv", each read_csv([Folder Path], [Name], ",", 11)),
    #"Remove Source Columns" = Table.SelectColumns(#"Invoke Custom Function",{"get_csv"}),
    #"Expand get_csv" = Table.ExpandTableColumn(#"Remove Source Columns", "get_csv", {"Administration Code", "Merchant Code", "OrderCode", "Date", "PaymentMethod", "Status", "CurrencyCode", "Amount", "Commission", "Batch ID", "Refusal Reason"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Expand get_csv",{{"Administration Code", type text}, {"Merchant Code", type text}, {"OrderCode", type text}, {"Date", type datetime}, {"PaymentMethod", type text}, {"Status", type text}, {"CurrencyCode", type text}, {"Amount", Currency.Type}, {"Commission", Currency.Type}, {"Batch ID", Int64.Type}, {"Refusal Reason", type text}})
in
    #"< Set Data Types >"