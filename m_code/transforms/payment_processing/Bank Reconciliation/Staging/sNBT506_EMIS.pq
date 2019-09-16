let
    import_currency = Excel.CurrentWorkbook(){[Name="import_currency"]}[Content]{0}[Column1],

    Source = iNBT506_EMIS,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"processed_date", "processed_date_temp"}, {"company_Id", "company_Id_temp"}, {"no_of_transactions", "no_transactions"}, {"value_of_transactions", "gross_transaction_value"}, {"no_of_chargebacks", "no_chargebacks"}, {"value_of_chargebacks", "chargebacks_value"}, {"sundry_entries", "sundries"}, {"service_charge_value", "msc_value"}}),
    #"Trimmed Text" = Table.TransformColumns(#"Renamed Columns",{{"company_Id_temp", Text.Trim}}),
    #"Cleaned Text" = Table.TransformColumns(#"Trimmed Text",{{"company_Id_temp", Text.Clean}}),

    #"Add [processed_date]" = Table.AddColumn(#"Cleaned Text", "processed_date",
        each Date.FromText(
            Text.Range([processed_date_temp], 0, 2)
            & "-"
            & Text.Range([processed_date_temp], 2, 2)
            & "-"
            & Text.Range([processed_date_temp], 4, 4)
        ), type date
    ),
    #"Add [company_code]" = Table.AddColumn(#"Add [processed_date]", "company_code", each if Text.Length([company_Id_temp]) < 6 then "0" & [company_Id_temp] else [company_Id_temp], type text),
    #"Left Join dCompanies" = Table.NestedJoin(#"Add [company_code]",{"company_code"},dCompanies,{"company_code"},"dMerchantAccounts_PSP",JoinKind.LeftOuter),
    #"Expanded dMerchantAccounts_PSP" = Table.ExpandTableColumn(#"Left Join dCompanies", "dMerchantAccounts_PSP", {"settlement_currency"}, {"currency"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded dMerchantAccounts_PSP",{"record_type", "processed_date_temp", "company_Id_temp"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"company_code", "currency", "processed_date", "no_transactions", "gross_transaction_value", "no_chargebacks", "chargebacks_value", "msc_value", "sundries", "settlement_value"}),
    #"Filter [currency]" = Table.SelectRows(#"Reordered Columns", each ([currency] = import_currency)),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Filter [currency]",{{"no_transactions", Int64.Type}, {"gross_transaction_value", Currency.Type}, {"no_chargebacks", Int64.Type}, {"chargebacks_value", Currency.Type}, {"msc_value", Currency.Type}, {"sundries", Currency.Type}, {"settlement_value", Currency.Type}})
in
    #"< Set Data Types >"