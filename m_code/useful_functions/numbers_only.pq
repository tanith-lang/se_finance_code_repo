// Remove all non numeric characters from a string
(source) => 
let
    numbers_only = (char) => if Character.ToNumber(char) >=48 and Character.ToNumber(char) < 58 then char else "",
    string = Text.From(source)
    len = Text.Length(string),
    acc = List.Accumulate(
                   List.Generate( () => 0, each _ < len, each _ + 1), 
                   "", 
                   (acc, index) => acc & numbers_only(Text.At(string, index))
               ),
    as_number = Number.FromText(acc)
in
    as_number