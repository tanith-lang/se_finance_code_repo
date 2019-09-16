// Convert UNIX epoch to datetime
let
    milliseconds = Record.FieldValues([EPOCH]){0},
    seconds = milliseconds / 1000,
    datetime = #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, seconds)
in
    datetime