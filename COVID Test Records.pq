let
    CurrentDir = Excel.CurrentWorkbook(){[Name = "FilePath"]}[Content]{0}[Column1],
    // Dynamic folder based on - https://www.goodly.co.in/dynamic-file-path-power-query/ + CONCATENATE (......, Source\)
    // The files need to be in the same directory as the main .xlsx file in \Source\ folder
    Source = Folder.Files(CurrentDir),
    #"Fetch excel objects" =
        Table.AddColumn(
            Source,
            "GetExcelObjects",
            each Excel.Workbook([Content], true)
        ),
    #"Unpack excel object: Data" =
        // Gets excel objects for all excel files
        Table.ExpandTableColumn(
            #"Fetch excel objects",
            "GetExcelObjects",
            {"Data"},
            {"GetExcelObjects.Data"}
        ),
    #"Remove source columns" =
        // Removes unnecessary source columns
        Table.RemoveColumns(
            #"Unpack excel object: Data",
            {
                "Extension",
                "Date accessed",
                "Date modified",
                "Date created",
                "Attributes",
                "Content"
            }
        ),
    #"Unpack object data" =
        // Unpacks objects data (i.e. unpacks all rows for all sheets in all excel files it found)
        Table.ExpandTableColumn(
            #"Remove source columns",
            "GetExcelObjects.Data",
            {
                "Column1",
                "Antigenní testování na přítomnost SARS-COV_2 (COVID-19)",
                "Column3",
                "Column4",
                "Column5"
            },
            {
                "GetExcelObjects.Data.Column1",
                "GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)",
                "GetExcelObjects.Data.Column3",
                "GetExcelObjects.Data.Column4",
                "GetExcelObjects.Data.Column5"
            }
        ),
    #"Add Index" =
        // Adds ad hoc index to all rows
        Table.AddIndexColumn(
            #"Unpack object data",
            "Index",
            0,
            1
        ),
    #"Filter out unnecessary rows from 2nd column" =
        // Filters out all unnecessary rows from second column"
        Table.SelectRows(
            #"Add Index",
            each
                (
                    [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> null
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Adresa zaměstnavatele"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Business management team 1"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "HSE Department"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "IČ zaměstnavatele"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Jméno a příjmení zaměstnance"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Název zaměstnavatele"
                    and [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Oddělení"
                // Do not get rid of "Datum testovani" so we can use it in next step
                )
        ),
    #"Get index of each new date field" =
        // Adds new column with [Index] value for each row incl. "Datum testovani", otherwise returns null
        Table.AddColumn(
            #"Filter out unnecessary rows from 2nd column",
            "Index of new date field",
            each
                if
                    [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    = "Datum testování"
                // We need to keep this field so we can easily pinpoint to where the next date is. 
                // If we don't have this here, we would have to be checking for something like "is cell value a date type..?"
                then
                    [Index]
                else
                    null
        ),
    #"Fill down to rows until next date value" =
        // Takes each index from previous collumn and fills down until next non-null value
        Table.FillDown(
            #"Get index of each new date field",
            {"Index of new date field"}
        ),
    #"Filter out rows with 'Datum testování'" =
        // Filters out "Datum testovani" which is not necessary anymore
        Table.SelectRows(
            #"Fill down to rows until next date value",
            each
                (
                    [#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]
                    <> "Datum testování"
                )
        ),
    #"Remove original index column" =
        // Removes original index value as it is not necessary anymore
        Table.RemoveColumns(
            #"Filter out rows with 'Datum testování'",
            {"Index"}
        ),
    #"Grouped" =
        /* Groups all rows by the new index and finds the minimum value for each row, which happens to be the date
        Otherwise we would have to be looking for something like "date value" within all of the available rows. */
        Table.Group(
            #"Remove original index column",
            {"Index of new date field"},
            {
                {
                    "all rows for index",
                    each _,
                    type table
                },
                {
                    "Test date",
                    each List.Min([#"GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)"]),
                    type anynonnull
                }
            }
        ),
    #"Remove new index column" =
        // Removes new index - not necessary anymore
        Table.RemoveColumns(
            Grouped,
            {"Index of new date field"}
        ),
    #"Unpacked all rows" =
        // Unpacks/expands all the rows for each of the grouped tables
        Table.ExpandTableColumn(
            #"Remove new index column",
            "all rows for index",
            {
                "Name",
                "Folder Path",
                "GetExcelObjects.Data.Column1",
                "GetExcelObjects.Data.Antigenní testování na přítomnost SARS-COV_2 (COVID-19)",
                "GetExcelObjects.Data.Column3",
                "GetExcelObjects.Data.Column4",
                "GetExcelObjects.Data.Column5"
            },
            {
                "all rows for index.Name",
                "all rows for index.Folder Path",
                "all rows for index.GetExcelObjects.Data.Column1",
                "all rows for index.GetExcelObjects.Data.Antigenní testování na přítomnost SARS-C",
                "all rows for index.GetExcelObjects.Data.Column3",
                "all rows for index.GetExcelObjects.Data.Column4",
                "all rows for index.GetExcelObjects.Data.Column5"
            }
        ),
    #"Filter non-null index rows" =
        // Filters out null value rows from Column1 (numbering for each person)
        Table.SelectRows(
            #"Unpacked all rows",
            each
                (
                    [all rows for index.GetExcelObjects.Data.Column1]
                    <> null
                )
        ),
    #"Remove distinct rows" =
        // Only distinct rows.
        Table.Distinct(
            #"Filter non-null index rows",
            {
                // This is a cummulative condition (all rows must be same in order to remove duplicity)
                // When setting this in the query editor just select more columns at the same time before removing duplicities
                "Test date",
                "all rows for index.GetExcelObjects.Data.Antigenní testování na přítomnost SARS-C",
                "all rows for index.GetExcelObjects.Data.Column3",
                "all rows for index.GetExcelObjects.Data.Column4"
            }
        ),
    #"Get plain date from file path" =
        // Gets date in plain format yymmDD from Folder path + adds 20 in front
        Table.AddColumn(
            #"Remove distinct rows",
            "Folder date yymmDD",
            each
                Text.Combine(
                    {
                        "20",
                        Text.BetweenDelimiters(
                            [all rows for index.Folder Path],
                            "\",
                            // delimiter
                            "\",
                            {
                                0,
                                // 
                                RelativePosition.FromEnd
                            // This is the last occurrence of "\"
                            },
                            {
                                0,
                                RelativePosition.FromEnd
                            // This is the last occurrence relative to the previous last occurrence ??? mindfuck
                            }
                        )
                    }
                ),
            type text
        ),
    #"Get plain date format from test date" =
        // Gets dates and changes to yyyymmDD
        Table.AddColumn(
            #"Get plain date from file path",
            "Test date yymmDD",
            each
                Text.Replace(
                    // Replaces "-" in korean date format to get plain yymmDD
                    Text.From([Test date], "ko-KR"),
                    // Gets korean version of date = yyyy-mm-DD
                    "-",
                    ""
                ),
            type text
        ),
    #"Compare yymmDD dates column" =
        Table.AddColumn(
            #"Get plain date format from test date",
            "Compare yymmDD dates",
            each
                if [Folder date yymmDD] = [Test date yymmDD] then
                    "OK"
                else
                    "CHECK ME"
        ),
    #"Reorder Columns" =
        Table.ReorderColumns(
            #"Compare yymmDD dates column",
            {
                "all rows for index.Name",
                "all rows for index.Folder Path",
                "Folder date yymmDD",
                "Compare yymmDD dates",
                "Test date yymmDD",
                "Test date",
                "all rows for index.GetExcelObjects.Data.Column1",
                "all rows for index.GetExcelObjects.Data.Antigenní testování na přítomnost SARS-C",
                "all rows for index.GetExcelObjects.Data.Column3",
                "all rows for index.GetExcelObjects.Data.Column4",
                "all rows for index.GetExcelObjects.Data.Column5"
            }
        )
in
    #"Reorder Columns"