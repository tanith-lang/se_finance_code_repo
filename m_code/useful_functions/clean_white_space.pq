// clean_whitespace() :: Remove all whitespace or defined character from input text
(text as text, optional char_to_trim as text) =>
let
    char = if char_to_trim = null then " " else char_to_trim,
    split = Text.Split(text, char),
    removeblanks = List.Select(split, each _ <> ""),
    result = Text.Combine(removeblanks, char)
in
    result