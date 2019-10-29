// function: read_csv :: Read the contents of a .csv file with specified parameters

// args: file_path (text) :: The folder path of the .csv file
// args: file_name (text) :: The name of the .csv file
// args: cols (number)     :: The number of columns the .csv file contains
// args: delimiter (text)  :: The delimiter (column separator) used in the .csv file (default="," (Comma))
// args: encoding (number) :: The character encoding code of the .csv file (default=65001 (Unicode UTF-8))

(file_path as text, file_name as text, columns as number, delimiter as text, encoding as number)=>

let
    full_path = file_path & file_name,
    source = Csv.Document(
                        File.Contents(full_path),
                        [Delimiter=delimiter, 
                        Columns=columns, 
                        Encoding=encoding, 
                        QuoteStyle=QuoteStyle.Csv]),
                        
    remove_blank_rows = Table.SelectRows(source,
                                        each not List.IsEmpty(
                                            List.RemoveMatchingItems(
                                                Record.FieldValues(_),
                                                {"", null}))),
    promote_headers = Table.PromoteHeaders(
                                        remove_blank_rows, 
                                        [PromoteAllScalars=true])
in
    promote_headers
    
