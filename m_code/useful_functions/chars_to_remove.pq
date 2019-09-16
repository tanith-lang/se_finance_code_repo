// Generates a list of non-numeric characters. Can be used within Text.Remove to remove
// all non-numeric characters from a column. The numbers passed to List.Transform are
// ASCII codes.

let
    chars_to_remove = List.Transform({32, 33..45,46, 47,58..126}, each Character.FromNumber(_)),
in
    chars_to-remove