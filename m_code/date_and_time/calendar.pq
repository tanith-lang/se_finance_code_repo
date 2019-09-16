// d_calendar table
let
    Source = List.Dates(StartDate, Length, #duration(1, 0, 0, 0)),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "date"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"date", type date}}),
        StartDate = #date(2016, 1, 1),
        Today = DateTime.Date(Date.AddMonths(DateTime.LocalNow(), 6)),
        Length = Duration.Days(Today - StartDate),
        Custom = #"Changed Type",
    #"Add [year]" = Table.AddColumn(Custom, "year", each Date.Year([date]), Int64.Type),
    #"Add [month_name]" = Table.AddColumn(#"Add [year]", "month_name", each Date.MonthName([date]), type text),
    #"Add [day_name]" = Table.AddColumn(#"Add [month_name]", "day_name", each Date.DayOfWeekName([date]), type text),
    #"Add [month_no]" = Table.AddColumn(#"Add [day_name]", "month_no", each Date.Month([date]), Int64.Type),
    #"Add [weekday_no]" = Table.AddColumn(#"Add [month_no]", "weekday_no", each Date.DayOfWeek([date])+1, Int64.Type),
    #"Add [mmm]" = Table.AddColumn(#"Add [weekday_no]", "mmm", each Text.Start([month_name], 3), type text),
    #"Add [ddd]" = Table.AddColumn(#"Add [mmm]", "ddd", each Text.Start([day_name], 3), type text),
    #"Add [yyyymm]" = Table.AddColumn(#"Add [ddd]", "yyyymm", each ([year] * 100) + [month_no], Int64.Type),
    #"Add [month_id]" = Table.AddColumn(#"Add [yyyymm]", "month_id", each (Date.Year([date]) - Date.Year(StartDate))*12 + Date.Month([date]), Int64.Type)
in
    #"Add [month_id]"