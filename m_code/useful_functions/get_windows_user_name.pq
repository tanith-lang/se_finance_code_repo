// Returns current Windows system user
let
    users = Folder.Contents("C:/Users"),
    current_user = Table.FirstN(Table.Sort(users,{{"Date accessed", Order.Descending}}),1)[Name]{0}
in
    current_user