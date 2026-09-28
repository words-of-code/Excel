# [EN] Dutch postcode validation, normalisation and extraction

Replace `A1` with the cell containing the postcode or address. The target format for normalisation is `2371DP`.

The expressions check the **shape** of a postcode: four digits (the first is not zero) and two letters, excluding `SA`, `SD` and `SS`. They do not check whether the postcode actually exists or belongs to the address. Only ordinary spaces are allowed between the digits and letters.

### 1. Validate a cell containing only a postcode

```excel
=REGEXTEST(TRIM(A1); "^[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}$"; 1)
```

The `^` and `$` anchors require the *whole cell* to match. `TRIM` removes spaces around the value and reduces repeated ordinary spaces inside it to one. The last argument, `1`, makes the comparison case insensitive: `2371 dp` is accepted. Because of `TRIM`, an input such as `2371   DP` is also accepted in this particular formula.

### 2. Check whether an address contains a valid postcode

```excel
=REGEXTEST(A1; "\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b"; 1)
```

This returns `TRUE` if it finds a matching postcode anywhere in the text. The `\b` tokens mark word boundaries, so surrounding spaces or punctuation do not become part of the match.

### 3. Extract a valid postcode from an address

```excel
=REGEXEXTRACT(A1; "\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b"; 0; 1)
```

The third argument, `0`, returns the **first** match; the fourth, `1`, makes the search case insensitive. Putting `1` in the *third* position would instead return **all** matches as an array. If there is no match, Excel returns `#N/A`.

### 4. Extract and normalise a valid postcode

```excel
=UPPER(SUBSTITUTE(REGEXEXTRACT(A1; "\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b"; 0; 1); " "; ""))
```

`REGEXEXTRACT` finds the postcode, `SUBSTITUTE` removes its space, and `UPPER` capitalises its letters. For example, `Main Street 12, 2371 dp Leiden` produces `2371DP`.

### 5. Normalise a cell containing only a valid postcode with REGEXREPLACE

```excel
=UPPER(REGEXREPLACE(TRIM(A1); "^([1-9][0-9]{3})[ ]?((?!SA|SD|SS)[A-Z]{2})$"; "$1$2"; 1; 1))
```

The parentheses capture the digits as `$1` and the letters as `$2`. `REGEXREPLACE` joins them, and `UPPER` capitalises the result. Its fourth argument, `1`, selects the first occurrence; its fifth argument, `1`, makes matching case insensitive. For invalid input, no replacement occurs, although `UPPER` still capitalises the original text. Validate separately if invalid values must be rejected.

### 6. Normalise only a valid postcode inside an address

This LAMBDA leaves the other words in the address exactly as they were. The final `(A1)` calls the function immediately; alternatively, save the LAMBDA in Name Manager and call it by its chosen name.

```excel
=LAMBDA(text;
    LET(
        postcode; REGEXEXTRACT(text; "\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b"; 0; 1);
        compact; UPPER(SUBSTITUTE(postcode; " "; ""));
        REPLACE(text; SEARCH(postcode; text); LEN(postcode); compact)
    )
)(A1)
```

`Main Street 12, 2371 dp Leiden` becomes `Main Street 12, 2371DP Leiden`. If no postcode is found, the formula returns `#N/A`. To keep unmatched addresses unchanged, wrap the *entire* formula in `IFNA( ... , A1)`.

> **More tolerance for messy data:** The examples use `[ ]?`, which allows **zero or one ordinary space**. If postcodes can contain several ordinary spaces between the digits and letters, replace `[ ]?` with `[ ]*` in the relevant regex. The latter allows **zero or more ordinary spaces**. This does not add support for tabs or nonbreaking spaces.

### Pattern at a glance

| Token | Meaning |
| --- | --- |
| `^` / `$` | Beginning / end of the entire cell |
| `\b` | Word boundary within longer text |
| `[1-9][0-9]{3}` | Four digits; the first cannot be zero |
| `[ ]?` | Zero or one ordinary space |
| `(?!SA\|SD\|SS)` | Exclude the three letter pairs |
| `[A-Z]{2}` | Exactly two letters, with case insensitive matching enabled by the function argument |
| `$1`, `$2` | Reuse captured groups in `REGEXREPLACE` |


---

# [NL] Nederlandse postcodes valideren, normaliseren en extraheren

Deze voorbeelden gebruiken Nederlandse Excel-functienamen en puntkomma's als scheidingsteken. Vervang `A1` door de cel met de postcode of het adres. Bij normalisatie is het gewenste formaat `2371DP`.

De expressies controleren de **opbouw** van een postcode: vier cijfers (het eerste is geen nul) en twee letters, met uitsluiting van `SA`, `SD` en `SS`. Ze controleren niet of de postcode werkelijk bestaat of bij het adres hoort. Tussen de cijfers en letters worden alleen gewone spaties toegestaan.

### 1. Een cel met alleen een postcode valideren

```excel
=REGEXTEST(SPATIES.WISSEN(A1);"^[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}$";1)
```

De ankers `^` en `$` eisen dat de *hele cel* overeenkomt. `SPATIES.WISSEN` verwijdert spaties rond de waarde en brengt herhaalde gewone spaties daarbinnen terug tot één. Het laatste argument, `1`, maakt de vergelijking ongevoelig voor hoofdletters: `2371 dp` wordt geaccepteerd. Door `SPATIES.WISSEN` wordt ook invoer zoals `2371   DP` door juist deze formule geaccepteerd.

### 2. Controleren of er een geldige postcode in een adres staat

```excel
=REGEXTEST(A1;"\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b";1)
```

Dit geeft `WAAR` terug als ergens in de tekst een overeenkomende postcode staat. De tekens `\b` duiden woordgrenzen aan: omliggende spaties of leestekens horen daardoor niet bij de overeenkomst.

### 3. De eerste geldige postcode uit een adres halen

```excel
=REGEXEXTRAHEREN(A1;"\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b";0;1)
```

Het derde argument, `0`, geeft de **eerste** overeenkomst terug; het vierde, `1`, maakt het zoeken ongevoelig voor hoofdletters. Een `1` op de *derde* positie zou juist **alle** overeenkomsten als matrix teruggeven. Zonder overeenkomst geeft Excel `#N/B` terug.

### 4. Een geldige postcode extraheren en normaliseren

```excel
=HOOFDLETTERS(SUBSTITUEREN(REGEXEXTRAHEREN(A1;"\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b";0;1);" ";""))
```

`REGEXEXTRAHEREN` zoekt de postcode, `SUBSTITUEREN` verwijdert de spatie en `HOOFDLETTERS` zet de letters in hoofdletters. Zo levert `Hoofdstraat 12, 2371 dp Leiden` de uitkomst `2371DP` op.

### 5. Een cel met alleen een postcode normaliseren met REGEXVERVANGEN

```excel
=HOOFDLETTERS(REGEXVERVANGEN(SPATIES.WISSEN(A1);"^([1-9][0-9]{3})[ ]?((?!SA|SD|SS)[A-Z]{2})$";"$1$2";1;1))
```

De haakjes leggen de cijfers vast als `$1` en de letters als `$2`. `REGEXVERVANGEN` voegt ze samen en `HOOFDLETTERS` zet de uitkomst in hoofdletters. Het vierde argument, `1`, kiest de eerste overeenkomst; het vijfde, `1`, maakt het zoeken ongevoelig voor hoofdletters. Bij ongeldige invoer vindt geen vervanging plaats, al zet `HOOFDLETTERS` de oorspronkelijke tekst nog wel in hoofdletters. Valideer apart als ongeldige waarden moeten worden afgewezen.

### 6. Alleen de postcode binnen een adres normaliseren

Deze LAMBDA laat de andere woorden in het adres precies zoals ze waren. De afsluitende `(A1)` roept de functie meteen aan; je kunt de LAMBDA ook in Naambeheer opslaan en daarna met de gekozen naam gebruiken.

```excel
=LAMBDA(tekst;
    LET(
        postcode;REGEXEXTRAHEREN(tekst;"\b[1-9][0-9]{3}[ ]?(?!SA|SD|SS)[A-Z]{2}\b";0;1);
        compact;HOOFDLETTERS(SUBSTITUEREN(postcode;" ";""));
        VERVANGEN(tekst;VIND.SPEC(postcode;tekst);LENGTE(postcode);compact)
    )
)(A1)
```

`Hoofdstraat 12, 2371 dp Leiden` wordt `Hoofdstraat 12, 2371DP Leiden`. Als er geen postcode wordt gevonden, geeft de formule `#N/B` terug. Wil je adressen zonder postcode ongewijzigd laten, zet dan `ALS.NB( ... ;A1)` om de *volledige* formule.

> **Meer tolerantie voor rommelige gegevens:** De voorbeelden gebruiken `[ ]?`: dit staat **nul of één gewone spatie** toe. Kunnen er meerdere gewone spaties tussen de cijfers en letters staan, vervang dan `[ ]?` door `[ ]*` in de betreffende regex. Daarmee zijn **nul of meer gewone spaties** toegestaan. Tabs en vaste spaties worden daardoor niet alsnog geaccepteerd.

### Het patroon in het kort

| Onderdeel | Betekenis |
| --- | --- |
| `^` / `$` | Begin / einde van de hele cel |
| `\b` | Woordgrens binnen langere tekst |
| `[1-9][0-9]{3}` | Vier cijfers; het eerste mag geen nul zijn |
| `[ ]?` | Nul of één gewone spatie |
| `(?!SA\|SD\|SS)` | Sluit deze drie letterparen uit |
| `[A-Z]{2}` | Precies twee letters, bij hoofdletterongevoelig zoeken via het functieargument |
| `$1`, `$2` | Hergebruik vastgelegde groepen in `REGEXVERVANGEN` |

