let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "BANK") and Text.Contains([Folder Path], "BNK_STTMNTS") and Text.Contains([Folder Path], "SVB-LTD") and Text.Contains([Folder Path], import_currency)),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], import_currency) and Text.Contains([Name], "SVB-LTD")),
    #"Invoke <getBankStatement>" = Table.AddColumn(#"Filter [Name]", "getBank_Svb", each getBank_Svb([Folder Path], [Name])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoke <getBankStatement>",{"getBank_Svb"}),
    #"Expanded getBank_Svb" = Table.ExpandTableColumn(#"Removed Other Columns", "getBank_Svb", {"date", "transaction_type", "bank_reference", "customer_reference", "description", "amount"}, {"date", "transaction_type", "bank_reference", "customer_reference", "description", "amount"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Expanded getBank_Svb",{{"date", type date}, {"transaction_type", type text}, {"bank_reference", type text}, {"customer_reference", type text}, {"description", type text}, {"amount", Currency.Type}})
in
    #"< Set Data Types >"