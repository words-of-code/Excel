## [EN] LAMBDA function: DATA.SOURCE

> [!NOTE]
> This page is mainly for reference, as the function is built into the Excel workbook.

The template uses the `DATA.SOURCE` LAMBDA function to build a complete data source path consistently.

It handles local file paths and URLs, with an optional subdirectory, date, or date override. It can also replace part of the filename and normalize path separators.

The function is stored as a named formula and can be used throughout the workbook.

### Create the function

Press **CTRL + F3** to open **Name Manager**, select **New**, and enter:

**Name:** `DATA.SOURCE`

**Refers to:**

```excel
=LAMBDA(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value;
    LET(
        credits; "https://github.com/words-of-code/Excel/, see: Dynamic Data Sources for Power Query";
        get_path; INDIRECT(TRIM(filepath));
        isURL; REGEXTEST(get_path; "://");
        seperator; IF(isURL; "/"; "\");
        path; IF(get_path <> ""; get_path; "%FILEPATH" & seperator);
        dir; LET(
            _x1; TRIM(subdir);
            _x2; REGEXREPLACE(_x1; "^[/\\]+|[/\\]+$"; ""; 0; 1);
            _x3; IF(isURL; SUBSTITUTE(_x2; "\"; "/"); SUBSTITUTE(_x2; "/"; "\"));
            _x4; IF(_x3 <> ""; _x3 & seperator; "");
            _x4
        );
        use_date; IF(TRIM(include_date) <> ""; TRUE; FALSE);
        the_date; LET(
            override; TRIM(override_date);
            custom_format; REGEXREPLACE(TRIM(include_date); "[^-dmy je]"; ""; 0; 1);
            format; IF(OR(include_date=TRUE; include_date=1); "emmdd"; LOWER(custom_format));
            _x1; IF(override <> ""; override; TEXT(TODAY(); format));
            _x2; "-" & _x1;
            _x2
        );
        ext; LET(
            _x1; TRIM(extension);
            _x2; IF(_x1 <> ""; "." & REGEXREPLACE(_x1; "^\W+"; ""); ".csv");
            _x2
        );
        file; LET(
            _x1; TRIM(filename);
            _x2; IF(_x1 <> ""; _x1; "%FILENAME");
            _x3; REGEXREPLACE(_x2; ext & "$"; ""; 0; 1);
            _x4; CONCAT(_x3; IF(use_date; the_date; ""); ext);
            _x5; IF(replace_pattern <> ""; REGEXREPLACE(_x4; replace_pattern; replace_value; 0; 1); _x4);
            _x5
        );
        result; IF(
            isURL;
            REGEXREPLACE(path & dir & file; "(?<!:)/+"; "/"; 0; 1);
            REGEXREPLACE(path & dir & file; "\\+"; "\\"; 0; 1)
        );
        result
    )
)
```

The compressed version below is easier to paste into the **Refers to** field:

```excel
=LAMBDA(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value; LET( credits; "https://github.com/words-of-code/Excel/, see: Dynamic Data Sources for Power Query"; get_path; INDIRECT(TRIM(filepath)); isURL; REGEXTEST(get_path; "://"); seperator; IF(isURL; "/"; "\"); path; IF(get_path <> ""; get_path; "%FILEPATH" & seperator); dir; LET( _x1; TRIM(subdir); _x2; REGEXREPLACE(_x1; "^[/\\]+|[/\\]+$"; ""; 0; 1); _x3; IF(isURL; SUBSTITUTE(_x2; "\"; "/"); SUBSTITUTE(_x2; "/"; "\")); _x4; IF(_x3 <> ""; _x3 & seperator; ""); _x4 ); use_date; IF(TRIM(include_date) <> ""; TRUE; FALSE); the_date; LET( override; TRIM(override_date); custom_format; REGEXREPLACE(TRIM(include_date); "[^-dmy je]"; ""; 0; 1); format; IF(OR(include_date=TRUE; include_date=1); "emmdd"; LOWER(custom_format)); _x1; IF(override <> ""; override; TEXT(TODAY(); format)); _x2; "-" & _x1; _x2 ); ext; LET( _x1; TRIM(extension); _x2; IF(_x1 <> ""; "." & REGEXREPLACE(_x1; "^\W+"; ""); ".csv"); _x2 ); file; LET( _x1; TRIM(filename); _x2; IF(_x1 <> ""; _x1; "%FILENAME"); _x3; REGEXREPLACE(_x2; ext & "$"; ""; 0; 1); _x4; CONCAT(_x3; IF(use_date; the_date; ""); ext); _x5; IF(replace_pattern <> ""; REGEXREPLACE(_x4; replace_pattern; replace_value; 0; 1); _x4); _x5 ); result; IF( isURL; REGEXREPLACE(path & dir & file; "(?<!:)/+"; "/"; 0; 1); REGEXREPLACE(path & dir & file; "\\+"; "\\"; 0; 1) ); result ) )
```

### Usage

Call the function as follows:

```excel
=DATA.SOURCE(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value)
```

The function has eight parameters. Use an empty string (`""`) for optional parameters you do not use.

| Parameter | Description |
|---|---|
| `filepath` | Name of a named range containing the base path or base URL. For a URL (identified by `://`), the function uses `/`; for a local path it uses `\`. If the range is empty, `%FILEPATH` appears as a placeholder. |
| `subdir` | Optional subdirectory, for example `source_data/2026`. Leading and trailing separators are removed. |
| `filename` | Base filename. If empty, `%FILENAME` appears as a placeholder. If the selected extension is already present, it is removed first to avoid duplication. |
| `extension` | Extension such as `csv`, `.xlsx`, or `json`. The leading dot is optional; an empty value defaults to `.csv`. |
| `include_date` | Empty: no date. `TRUE` or `1`: append `-` and today's date in `emmdd` format (year, month, day). The `e` stands for the year independently of Excel's display language; e.g. `yyyy` (EN) and `jjjj` (NL). A custom format can contain `-`, spaces, `d`, `m`, `y`, `j`, and `e`; other characters are filtered out. |
| `override_date` | Replaces the generated date with the supplied value. It only affects the result when `include_date` is set. |
| `replace_pattern` | Optional regular expression to match text in the assembled filename, after the date and extension have been added. An empty value skips replacement. |
| `replace_value` | Replacement text for the match found by `replace_pattern`. This can change part of the filename; the base path and subdirectory are not included in this replacement. |

## Result

The function builds the result roughly as follows:

```text
filepath + subdir + filename + date + extension
```

For example:

```text
C:\Data\Exports\archive\products-20260928.csv
```

Or, for a URL:

```text
https://example.com/data/archive/products-20260928.csv
```

Finally, duplicate path separators are removed. For URLs, this preserves the double slash in `https://`.


## [NL] LAMBDA-functie: DATA.SOURCE

> [!NOTE]
> Deze pagina is vooral informatief van aard aangezien de functie ingebakken zit in het Excel-bestand.

De template gebruikt de LAMBDA-functie `DATA.SOURCE` om op een uniforme manier het volledige pad naar een databron samen te stellen.

De functie verwerkt lokale bestandspaden en URL's, met een optionele submap, datum of afwijkende datumwaarde. Ook kan zij een deel van de bestandsnaam vervangen en scheidingstekens in het pad normaliseren.

De functie is als benoemde formule in de werkmap opgeslagen en kan daarna vanuit de gehele werkmap worden gebruikt.

### Functie aanmaken

Druk op **CTRL + F3** om **Namen beheren** te openen, kies **Nieuw** en voer het volgende in:

**Naam:** `DATA.SOURCE`

**Verwijst naar:**

```excel
=LAMBDA(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value; 
    LET( 
        credits; "https://github.com/words-of-code/Excel/, see: Dynamic Data Sources for Power Query";
        get_path; INDIRECT( SPATIES.WISSEN( filepath ) ); 
        isURL; REGEXTEST( get_path; "://" ); 
        seperator; ALS( isURL; "/"; "\" ); 
        path; ALS( get_path <> ""; get_path; "%FILEPATH" & seperator ); 
        dir; LET( 
            _x1; SPATIES.WISSEN( subdir ); 
            _x2; REGEXVERVANGEN( _x1; "^[/\\]+|[/\\]+$"; ""; 0; 1 ); 
            _x3; ALS( isURL; SUBSTITUEREN( _x2; "\"; "/" ); SUBSTITUEREN( _x2; "/"; "\" ) ); 
            _x4; ALS( _x3 <> ""; _x3 & seperator; "" );
			_x4
        ); 
        use_date; ALS( SPATIES.WISSEN( include_date ) <> ""; WAAR; ONWAAR ); 
        the_date; LET( 
            override; SPATIES.WISSEN( override_date ); 
            custom_format; REGEXVERVANGEN( SPATIES.WISSEN( include_date ); "[^-dmy je]"; ""; 0; 1 ); 
            format; ALS( OF( include_date=WAAR; include_date=1 ); "emmdd"; KLEINE.LETTERS( custom_format ) ); 
            _x1; ALS( override <> ""; override; TEKST( VANDAAG(); format ) ); 
            _x2; "-" & _x1;
			_x2
        ); 
        ext; LET( 
            _x1; SPATIES.WISSEN( extension ); 
            _x2; ALS( _x1 <> ""; "." & REGEXVERVANGEN( _x1; "^\W+"; "" ); ".csv" );
			_x2
        ); 
        file; LET( 
            _x1; SPATIES.WISSEN( filename ); 
            _x2; ALS( _x1 <> ""; _x1; "%FILENAME" ); 
            _x3; REGEXVERVANGEN( _x2; ext & "$"; ""; 0; 1 ); 
            _x4; TEKST.SAMENV( _x3; ALS( use_date; the_date; "" ); ext ); 
            _x5; ALS( replace_pattern <> ""; REGEXVERVANGEN( _x4; replace_pattern; replace_value; 0; 1 ); _x4 );
			_x5
        ); 
        result; ALS(
			isURL;
			REGEXVERVANGEN( path & dir & file; "(?<!:)/+"; "/"; 0; 1 );
			REGEXVERVANGEN( path & dir & file; "\\+"; "\\"; 0; 1 )
		);
		result
    ) 
)
```

Voor het plakken in het veld **Verwijst naar** is het in dit geval praktischer om de onderstaande gecomprimeerde code te gebruiken:

```excel
=LAMBDA(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value; LET( credits; "https://github.com/words-of-code/Excel/, see: Dynamic Data Sources for Power Query"; get_path; INDIRECT( SPATIES.WISSEN( filepath ) ); isURL; REGEXTEST( get_path; "://" ); seperator; ALS( isURL; "/"; "\" ); path; ALS( get_path <> ""; get_path; "%FILEPATH" & seperator ); dir; LET( _x1; SPATIES.WISSEN( subdir ); _x2; REGEXVERVANGEN( _x1; "^[/\\]+|[/\\]+$"; ""; 0; 1 ); _x3; ALS( isURL; SUBSTITUEREN( _x2; "\"; "/" ); SUBSTITUEREN( _x2; "/"; "\" ) ); _x4; ALS( _x3 <> ""; _x3 & seperator; "" ); _x4 ); use_date; ALS( SPATIES.WISSEN( include_date ) <> ""; WAAR; ONWAAR ); the_date; LET( override; SPATIES.WISSEN( override_date ); custom_format; REGEXVERVANGEN( SPATIES.WISSEN( include_date ); "[^-dmy je]"; ""; 0; 1 ); format; ALS( OF( include_date=WAAR; include_date=1 ); "emmdd"; KLEINE.LETTERS( custom_format ) ); _x1; ALS( override <> ""; override; TEKST( VANDAAG(); format ) ); _x2; "-" & _x1; _x2 ); ext; LET( _x1; SPATIES.WISSEN( extension ); _x2; ALS( _x1 <> ""; "." & REGEXVERVANGEN( _x1; "^\W+"; "" ); ".csv" ); _x2 ); file; LET( _x1; SPATIES.WISSEN( filename ); _x2; ALS( _x1 <> ""; _x1; "%FILENAME" ); _x3; REGEXVERVANGEN( _x2; ext & "$"; ""; 0; 1 ); _x4; TEKST.SAMENV( _x3; ALS( use_date; the_date; "" ); ext ); _x5; ALS( replace_pattern <> ""; REGEXVERVANGEN( _x4; replace_pattern; replace_value; 0; 1 ); _x4 ); _x5 ); result; ALS( isURL; REGEXVERVANGEN( path & dir & file; "(?<!:)/+"; "/"; 0; 1 ); REGEXVERVANGEN( path & dir & file; "\\+"; "\\"; 0; 1 ) ); result ) )
```

### Gebruik

De functie wordt aangeroepen met:


```excel
=DATA.SOURCE(filepath; subdir; filename; extension; include_date; override_date; replace_pattern; replace_value)
```

De functie bevat acht parameters. Geef voor ongebruikte optionele parameters een lege tekenreeks (`""`) op.

| Parameter | Omschrijving |
|---|---|
| `filepath` | Naam van een benoemd bereik met het basispad of de basis-URL. Bij een URL (herkend aan `://`) gebruikt de functie `/`; bij een lokaal pad `\`. Als de inhoud van het bereik leeg is, verschijnt `%FILEPATH` als tijdelijke aanduiding. |
| `subdir` | Optionele submap, bijvoorbeeld `brondata/2026`. Begin- en eindscheidingstekens worden verwijderd. |
| `filename` | Basisbestandsnaam. Als deze leeg is, verschijnt `%FILENAME` als tijdelijke aanduiding. Een al aanwezige, gekozen extensie wordt eerst verwijderd om verdubbeling te voorkomen. |
| `extension` | Extensie zoals `csv`, `.xlsx` of `json`. Een punt aan het begin is optioneel; leeg betekent `.csv`. |
| `include_date` | Leeg: geen datum. `WAAR` of `1`: voeg `-` plus de huidige datum in het formaat `emmdd` toe (jaar, maand, dag). De `e` staat voor het jaar en werkt onafhankelijk van de weergavetaal van Excel; `jjjj` en `yyyy` zijn taalgebonden. Een eigen formaat kan de tekens `-`, spatie, `d`, `m`, `y`, `j` en `e` bevatten; overige tekens worden weggefilterd. |
| `override_date` | Vervangt de gegenereerde datum door de opgegeven waarde. Heeft alleen zichtbaar effect als `include_date` is ingevuld. |
| `replace_pattern` | Optionele reguliere expressie die zoekt binnen de samengestelde bestandsnaam, na toevoeging van datum en extensie. Leeg betekent geen vervanging. |
| `replace_value` | Vervangende tekst voor wat `replace_pattern` vindt. Zo kan een deel van de bestandsnaam worden aangepast; het basispad en de submap worden niet door deze vervanging bewerkt. |

## Opbouw van het resultaat

De functie bouwt het resultaat in hoofdlijnen op als:

```text
filepath + subdir + filename + datum + extension
```

Bijvoorbeeld:

```text
C:\Data\Exports\archive\products-20260928.csv
```

of bij een URL:

```text
https://example.com/data/archive/products-20260928.csv
```

Als laatste stap worden dubbele scheidingstekens uit het samengestelde pad verwijderd.

Voor URL's gebeurt dit zonder de dubbele slash uit bijvoorbeeld:

```text
https://
```

te verwijderen.

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)