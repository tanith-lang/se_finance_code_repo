let
    Source = #"iAcquirerBatchDetails",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Event Type", "Status"}, {"Account", "Bank Account"}}),
    #"Remove [BackOffice Code]" = Table.RemoveColumns(#"Renamed Columns",{"BackOffice Code"}),
    #"Filter [Payment Method]" = Table.SelectRows(#"Remove [BackOffice Code]", each ([Payment Method] <> "AMEX-SSL" and [Payment Method] <> "PAYPAL-EXPRESS")),
    #"Add [Marker]" = Table.AddColumn(#"Filter [Payment Method]", "Marker", each 1),
    #"Group by [Marker]" = Table.Group(#"Add [Marker]", {"Marker"}, {{"Control", each Table.RowCount(_), type number}}),
    #"Remove [Marker]" = Table.RemoveColumns(#"Group by [Marker]",{"Marker"}),
    #"< Set Data Types >" = Table.TransformColumnTypes(#"Remove [Marker]",{{"Control", Int64.Type}})
in
    #"< Set Data Types >"