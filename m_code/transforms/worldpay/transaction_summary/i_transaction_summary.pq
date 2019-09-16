let
    folder_path = Excel.CurrentWorkbook(){[Name="folderPath"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="currencyCode"]}[Content]{0}[Column1],
    Source = Folder.Files(folder_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "TRNSCTN_SMMRY")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "Transaction-Summary")),
    #"< FILTER REPORTS >" = Table.SelectRows(#"Filter [Name]", each ([Name] = "2018-01_Transaction-Summary.csv" or [Name] = "2018-02_Transaction-Summary.csv" or [Name] = "2018-03_Transaction-Summary.csv" or [Name] = "2018-04_Transaction-Summary.csv" or [Name] = "2018-05_Transaction-Summary.csv" or [Name] = "2018-06_Transaction-Summary.csv" or [Name] = "2018-07_Transaction-Summary.csv" or [Name] = "2018-08_Transaction-Summary.csv" or [Name] = "2018-09_Transaction-Summary.csv" or [Name] = "2018-10_Transaction-Summary.csv")),
    #"Invoked Custom Function" = Table.AddColumn(#"< FILTER REPORTS >", "read_csv", each read_csv([Folder Path], [Name], ",", 11)),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoked Custom Function",{"read_csv"}),
    #"Expanded read_csv" = Table.ExpandTableColumn(#"Removed Other Columns", "read_csv", {"Administration Code", "Merchant Code", "OrderCode", "Date", "PaymentMethod", "Status", "CurrencyCode", "Amount", "Commission", "Batch ID", "Refusal Reason"}, {"admin_code", "merchant_code", "order_code", "datetime", "payment_method", "status", "currency", "amount", "commission", "batch_id", "refusal_reason"}),
    #"< SET DATA TYPES >" = Table.TransformColumnTypes(#"Expanded read_csv",{{"admin_code", type text}, {"merchant_code", type text}, {"order_code", type text}, {"datetime", type datetime}, {"payment_method", type text}, {"status", type text}, {"currency", type text}, {"amount", type number}, {"commission", type number}, {"batch_id", Int64.Type}, {"refusal_reason", type text}})
in
    #"< SET DATA TYPES >"