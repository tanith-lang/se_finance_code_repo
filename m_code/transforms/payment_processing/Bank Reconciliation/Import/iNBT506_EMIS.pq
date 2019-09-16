let
    import_path = Excel.CurrentWorkbook(){[Name="import_path"]}[Content]{0}[Column1],
    import_month = Excel.CurrentWorkbook(){[Name="import_month"]}[Content]{0}[Column1],

    Source = Folder.Files(import_path),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "WORLDPAY") and Text.Contains([Folder Path], "NBT506")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each ([Extension] = ".CSV")),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "SECX506")),
    #"Invoke Custom Function" = Table.AddColumn(#"Filter [Name]", "getNBT506_EMIS", each getNBT506_EMIS([Folder Path], [Name])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoke Custom Function",{"getNBT506_EMIS"}),
    #"Expand getNBT506_EMIS" = Table.ExpandTableColumn(#"Removed Other Columns", "getNBT506_EMIS", {"record_type", "processed_date", "no_of_transactions", "value_of_transactions", "no_of_chargebacks", "value_of_chargebacks", "service_charge_value", "sundry_entries", "settlement_value", "company_Id"}, {"record_type", "processed_date", "no_of_transactions", "value_of_transactions", "no_of_chargebacks", "value_of_chargebacks", "service_charge_value", "sundry_entries", "settlement_value", "company_Id"})
in
    #"Expand getNBT506_EMIS"