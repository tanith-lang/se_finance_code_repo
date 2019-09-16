// Nested functions

// When nesting functions, start from the inner function and work outwards. 
// For example, on line 9, Text.BetweenDelimiters() is used to extract a
// sub-string from a text field, then Text.Start() extracts the first 2
// characters from the sub-string, then Number.From() casts the resulting
// Text value as a Number.

// Finally, a datetime value is created by passing the extracted datetime
// components into the #datetime constructor

let
    year = Number.From("20" & Text.BetweenDelimiters([text_field]], "-", " ", 1, 0)),
    month = Number.From(Text.BetweenDelimiters([text_field]], "-", "-", 0, 0)),
    day = Number.From(Text.BetweenDelimiters([text_field]], " ", "-")),
    hour = Number.From(Text.Start(Text.BetweenDelimiters([text_field]], " ", ".", 1, 0), 2)),
    minute = Number.From(Text.End(Text.BetweenDelimiters([text_field]], " ", ".", 1, 0), 2)),
    second = 0,
    timestamp = #datetime(year, month, day, hour, minute, second)
in
    timestamp

// Note: "datetime" is a reserved keyword, so I have used "timestamp" to
// refer to that step.

// Input text: COMPONENT_BOOKINGS_01 15-02-19 0101.CSV
// Output datetime: 15/02/2019 01:01:00

// Notice how the year element of the Input text does not include the "20"
// of "2019". As Text.BetweenDelimiters returns a string data type, we can
// concatenate the result onto the string "20" with the "&" operator.