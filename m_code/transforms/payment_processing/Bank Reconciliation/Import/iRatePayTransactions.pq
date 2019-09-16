let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "RATEPAY") and Text.Contains([Folder Path], "TRNSCTN_SMMRY")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "settlement_report") and Text.Contains([Name], "2018")),
    #"Add [Import Month]" = Table.AddColumn(#"Filter [Name]", "import_month", each Text.Range([Name], 4, 2), type text),
    #"Add [merchant_code]" = Table.AddColumn(#"Add [Import Month]", "merchant_code", each Text.BetweenDelimiters([Name], "_", "_", 0, 0), type text),
    #"Filter [import_month]" = Table.SelectRows(#"Add [merchant_code]", each [import_month] = import_month),
    #"Left Join dMerchantAccounts" = Table.NestedJoin(#"Filter [import_month]",{"merchant_code"},dMerchantAccounts_PSP,{"company_code"},"dMerchantAccounts",JoinKind.LeftOuter),
    #"Expanded dMerchantAccounts" = Table.ExpandTableColumn(#"Left Join dMerchantAccounts", "dMerchantAccounts", {"settlement_currency"}, {"settlement_currency"}),
    #"Filter [settlement_currency]" = Table.SelectRows(#"Expanded dMerchantAccounts", each [settlement_currency] = import_currency),
    #"Filter [Import Month]" = Table.SelectRows(#"Filter [settlement_currency]", each true),
    #"Invoke Custom Function" = Table.AddColumn(#"Filter [Import Month]", "read_csv", each read_csv([Folder Path], [Name], ";", 14)),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoke Custom Function",{"settlement_currency", "read_csv"}),
    #"Expand read_csv" = Table.ExpandTableColumn(#"Removed Other Columns", "read_csv", {"SHOP_ID", "PAYMENTDATE", "PAYMENTMONTH", "SHOPNAME", "AMOUNT", "DESCRIPTOR", "SHOPINVOICE_ID", "SHOPSORDER_ID", "INVOICENUMBER", "DESCRIPTION", "FEETYPE", "ORDERDATE", "SENTDATE", "TRANSACTION_ID"}, {"SHOP_ID", "PAYMENTDATE", "PAYMENTMONTH", "SHOPNAME", "AMOUNT", "DESCRIPTOR", "SHOPINVOICE_ID", "SHOPSORDER_ID", "INVOICENUMBER", "DESCRIPTION", "FEETYPE", "ORDERDATE", "SENTDATE", "TRANSACTION_ID"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Expand read_csv",{{"SHOP_ID", type text}, {"PAYMENTDATE", type date}, {"PAYMENTMONTH", type date}, {"SHOPNAME", type text}, {"AMOUNT", Currency.Type}, {"DESCRIPTOR", type text}, {"SHOPINVOICE_ID", type text}, {"SHOPSORDER_ID", type text}, {"INVOICENUMBER", type text}, {"DESCRIPTION", type text}, {"FEETYPE", type text}, {"ORDERDATE", type date}, {"SENTDATE", type date}, {"TRANSACTION_ID", type text}})
in
    #"< Set Data Types >"