# Dynamic Data Sources in Excel / Power Query
<sub>version 1.1, to be used with `Dynamic Data Sources for Power Query (EN).xlsx`</sub>

The configuration worksheet lets you determine which files you want to import using Power Query without having to include a fixed file location in every query.

First, you define the base locations. You then build the filename for each data source and finally use the named range to retrieve the information in Power Query.

The Power Query functions **LoadCSV**, **LoadJSON**, and **LoadXLS** are already included in the Excel template. The examples below use a CSV file and LoadCSV.

## Step 1: Define file locations

The first table contains the base locations of your data files. Each location has a unique name starting with **path_**.

This name can be selected from the drop-down list in Step 2, allowing you to update a changed location in one central place.

- **path_dynamic**\
  Refers to the location of the currently opened Excel file. If the workbook is moved together with its associated source files, the path can therefore change automatically.\
  *(This path is only available after the Excel file has been saved at least once.)*

- **path_downloads**\
  Refers to your personal Downloads folder. Replace **\_USERNAME\_** in the path with your own Windows username, for example: `C:\Users\Words-of-Code\Downloads\`

To add an additional file location:

1. Use an empty row in the table or, if necessary, add another row.
2. Give the location in column A (**ID**) a unique name, for example **path_projects**.\
   Use the same name for the named range in column B (**File location**).
3. Enter the path to the folder in column B (**File location**). The path must end with `\` or `/`, depending on the type of path: local/network or online.

*Check that the path refers to an existing folder and that the separator between the folder and filename in column J (**Data source**) is correct.*

> [!NOTE]
> **Creating a named range**\
> Select the cell you want to name. Enter the desired name, for example *csv_projects*, in the Name Box to the left of the formula bar, where the cell address is normally displayed, and press **Enter**. You can also create it via: **Ctrl+F3 → Name Manager → New**.

## Step 2: Define data sources

Create one row in the second table for every file you want to import.

Complete the columns with blue headers. Column J (**Data source**) builds the full path and filename using the `=DATA.SOURCE()` formula. The resulting value also provides an immediate check of your configuration.

- **ID**\
  Give the data source a recognizable name, for example *csv_projects*. Use the same name for the named range in column J (**Data source**).

- **File location**\
  Select a *path_* name. Because this refers to the first table, a change to a file location only needs to be made in one central place.

- **Subfolder**\
  Enter a subfolder if required, for example *source_data*. If the file is located directly in the selected folder alongside the Excel workbook, leave this field empty.\
  For example, if you only use the *source_data* subfolder for network locations, you can also handle this dynamically with a formula such as: `=IFERROR(IF(REGEXTEST(INDIRECT([@[File location]]); "://"); "source_data"; "");"")`

- **Filename (base)**\
  Enter the fixed part of the filename. For *demo_projects-20260927.csv*, this would be *demo_projects*. The date, including the preceding hyphen, and the file extension are taken from the other columns.

- **File extension**\
  Enter the extension without a period, for example *csv*, *json*, or *xlsx*.

  The template currently supports these three file types through LoadCSV, LoadJSON, and LoadXLS.

- **Include date?**\
  Enter `TRUE` or `1` to use the standard date suffix. The default pattern in the template is `eyymmdd` (year, month, day), as used in: `demo_projects-20260927.csv`. You can also enter an alternative date pattern. Check the result in **Data source**, particularly when different Excel language settings are being used.
  - The following characters are allowed in the date pattern: `d`, `m`, `y`, `j`, `e`, `-`, and ` ` (space).
  - The character `e` is the display-language-independent variant for year, for example `jjjj` or `yyyy`.
  - Use `dd` and `mm` respectively if you want leading zeroes for the day and month.

- **Override date with**\
  Enter a fixed date or another number here if you do not want to use the current date. Make sure **Include date?** is enabled as well.

- **REGEX search pattern**\
  (optional) Specify the pattern for the part of the generated filename that you want to modify.

- **REGEX replace with**\
  (optional) Specify the replacement text. Leave this empty to remove the matched part.

- **Data source**\
  Check the complete file location here and assign the cell a named range, for example *csv_projects*. Use the same name in column A (**ID**) as a reminder.

**Example configuration:**

| File location | Subfolder | Filename (base) | Extension | Include date? |
|---|---|---|---|---|
| path_dynamic | source_data | demo_projects | csv | = TRUE |

This could, for example, result in:

`https://organisationname.sharepoint.com/personal/_username_/Documents/Desktop/source_data/demo_projects-20260927.csv`

The URL above is only an example. The value in your column J (**Data source**) must point to a file that you can actually access.

## Step 3: Use the data source in Power Query

You can first let Excel create a query through **Data → Get Data → From File** or start with a blank query through **Data → Get Data → From Other Sources → Blank Query**.
Then, in the Power Query Editor, open **Home → Advanced Editor**. This allows you to view and edit the complete M code in one place.

The two source steps retrieve the path from the named range and pass it to the **LoadCSV** function.

Inside a `let` block, these lines are written without an `=` at the beginning:

```powerquery
SourceFile = Text.From(Excel.CurrentWorkbook(){[Name="csv_projects"]}[Content]{0}[Column1]),
Source = LoadCSV(SourceFile, ";")
```

- **SourceFile**\
  Reads the text value from the first cell of the *csv_projects* named range in the current workbook.

- **Source**\
  Uses LoadCSV to open the file. The second argument is the delimiter used by the CSV file. Use `,` for a comma and `;` for a semicolon.\
  If you first opened the file through the Excel interface, add `SourceFile` above the existing source step and replace that source step with `Source = LoadCSV(SourceFile, ",")`. Leave the remaining transformations in place and check the preview afterwards. Adjust subsequent steps if the structure of the loaded data requires it.

When starting with a blank query, you can paste this complete example directly into the **Advanced Editor**:

```powerquery
// Load CSV
let
    SourceFile = Text.From(Excel.CurrentWorkbook(){[Name="csv_projects"]}[Content]{0}[Column1]),
    Source = LoadCSV(SourceFile, ";")
in
    #"Source"
```

Replace *csv_projects* with the name of your named range and choose the delimiter used by the CSV file.

### JSON files

For JSON files, use `LoadJSON`. When starting with a blank query, you can paste this complete example directly into the **Advanced Editor**:

```powerquery
// Load JSON
let
    SourceFile = Text.From(Excel.CurrentWorkbook(){[Name="json_projects"]}[Content]{0}[Column1]),
    Source = LoadJSON(SourceFile)
in
    #"Source"
```

### Excel files

For Excel files, use **LoadXLS** and specify the name of the worksheet or Excel table that you want to load.

To load a worksheet:

```powerquery
// Load a worksheet
let
    SourceFile = Text.From(Excel.CurrentWorkbook(){[Name="xls_projects"]}[Content]{0}[Column1]),
    Source = LoadXLS(SourceFile),
    #"Data" = Source{[Item="WorksheetName",Kind="Sheet"]}[Data]
in
    #"Data"
```

To load an Excel table:

```powerquery
// Load an Excel table
let
    SourceFile = Text.From(Excel.CurrentWorkbook(){[Name="xls_projects"]}[Content]{0}[Column1]),
    Source = LoadXLS(SourceFile),
    #"Data" = Source{[Item="TableName",Kind="Table"]}[Data]
in
    #"Data"
```

## Finally

The template provides an easy way to load CSV, JSON, and XLSX — as well as XLS and XLSM — files without hardcoding the path and filename in your Power Query queries. For now, I am keeping the template focused on these three file types.

The first extension will probably be support for retrieving data from an online database, such as MariaDB / MySQL, once I start using that functionality myself.

In the meantime, feel free to experiment with using the approach in other queries.

And if you come up with a useful solution or improvement, please share it with me as well.

<!-- [![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6) -->
