let
    Source = Folder.Files("C:\Users\saur.dash\Dropbox (Secret Escapes)\DATA_STORE"),
    #"Filter [Folder Path]" = Table.SelectRows(Source, each Text.Contains([Folder Path], "RATEPAY") and Text.Contains([Folder Path], "TRNSCTN_SMMRY")),
    #"Filter [Extension]" = Table.SelectRows(#"Filter [Folder Path]", each [Extension] = ".csv"),
    #"Filter [Name]" = Table.SelectRows(#"Filter [Extension]", each Text.Contains([Name], "settlement_report")),
    #"< FILTER REPORT YEAR >" = Table.SelectRows(#"Filter [Name]", each Text.Contains([Name], "2018")),
    #"Invoked read_csv()" = Table.AddColumn(#"< FILTER REPORT YEAR >", "read_csv", each read_csv([Folder Path], [Name], 14, ";")),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoked read_csv()",{"read_csv"}),
    #"Expanded read_csv()" = Table.ExpandTableColumn(#"Removed Other Columns", "read_csv", {"SHOP_ID", "PAYMENTDATE", "PAYMENTMONTH", "SHOPNAME", "AMOUNT", "DESCRIPTOR", "SHOPINVOICE_ID", "SHOPSORDER_ID", "INVOICENUMBER", "DESCRIPTION", "FEETYPE", "ORDERDATE", "SENTDATE", "TRANSACTION_ID"}, {"shop_id", "payment_date", "payment_month", "shop_name", "amount", "descriptor", "shop_invoice_id", "shop_order_id", "invoice_number", "description", "fee_type", "order_date", "sent_date", "transaction_id"}),
    #"< SET DATA TYPES >" = Table.TransformColumnTypes(#"Expanded read_csv()",{{"shop_id", type text}, {"payment_date", type date}, {"payment_month", type date}, {"shop_name", type text}, {"amount", type number}, {"descriptor", type text}, {"shop_invoice_id", type text}, {"shop_order_id", type text}, {"invoice_number", Int64.Type}, {"description", type text}, {"fee_type", Int64.Type}, {"order_date", type date}, {"sent_date", type date}, {"transaction_id", type text}})
in
    #"< SET DATA TYPES >"