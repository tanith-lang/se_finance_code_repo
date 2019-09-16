// Table.NestedJoin syntax (example only - code will not run)
= Table.NestedJoin
(
	table_01, {"Column 01", "Column 02"},	//LEFT TABLE COLUMNS TO MATCH
	table_02, {"Column 01", "Column 02"},	//RIGHT TABLE COLUMNS TO MATCH
	"NewColumn",				//MERGE COLUMN NAME
	JoinKind.LeftOuter			//JOIN TYPE		
)

let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],

    // we add a custom column to differentiate between the different parts of the tables
    #"Added Custom" = Table.AddColumn(Source, "Custom", each if [Rating1] = 0 then 0 else 1),

    // Now we create two tables from Table1, by grouping by the custom column
    #"Grouped Rows" = Table.Group(#"Added Custom", {"Custom"}, {{"Count", each _, type table}}),

    // We drill down to each part of the table
    rating1Table = #"Grouped Rows"{[Custom=1]}[Count],
    rating2Table = #"Grouped Rows"{[Custom=0]}[Count],

    // Join Table2 with the part in which Rating1 > 0    
    rating1WithTable2 = Table.NestedJoin(rating1Table, {"Rating1"},Table2,{"Rating1"},"NewColumn",JoinKind.Inner),
    expandRating1 = Table.ExpandTableColumn(rating1WithTable2, "NewColumn", {"Rating Description"}, {"Rating Description"}),

    // Join Table3 with the part in which Rating1 = 0    
    rating2WithTable3 = Table.NestedJoin(rating2Table, {"Rating2"},Table3,{"Rating2"},"NewColumn",JoinKind.Inner),
    expandRating2 = Table.ExpandTableColumn(rating2WithTable3, "NewColumn", {"Rating Description"}, {"Rating Description"}),
    
    // We combine the two tables
    result = Table.Combine(expandRating1 , expandRating2)
in
    result