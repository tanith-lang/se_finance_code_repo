let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "BANK") and Text.Contains([Folder Path], "BNK_STTMNTS") and Text.Contains([Folder Path], "LTD") and Text.Contains([Folder Path], import_currency)),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], import_currency) and not Text.Contains([Name], "SVB")),
    #"Invoke <getBankStatement>" = Table.AddColumn(#"Filter [Name]", "getBankStatement", each getBankStatement([Folder Path], [Name])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoke <getBankStatement>",{"getBankStatement"}),
    #"Expanded getBankStatement" = Table.ExpandTableColumn(#"Removed Other Columns", "getBankStatement", {"ACCOUNT NO", "TYPE", "BANK CODE", "CURR", "ENTRY DATE", "AS AT", "AMOUNT", "TLA CODE", "CHEQUE NO", "STATUS", "DESCRIPTION"}, {"ACCOUNT NO", "TYPE", "BANK CODE", "CURR", "ENTRY DATE", "AS AT", "AMOUNT", "TLA CODE", "CHEQUE NO", "STATUS", "DESCRIPTION"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Expanded getBankStatement",{{"ACCOUNT NO", type text}, {"TYPE", type text}, {"BANK CODE", type text}, {"CURR", type text}, {"ENTRY DATE", type date}, {"AS AT", type text}, {"AMOUNT", Currency.Type}, {"TLA CODE", type text}, {"CHEQUE NO", type text}, {"STATUS", type text}, {"DESCRIPTION", type text}})
in
    #"< Set Data Types >"