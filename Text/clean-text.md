## [EN] Text cleaner: trims text and removes hidden characters

When copying text from another source, such as a PDF, website, or external file, the copied value may contain unwanted spaces or invisible characters. These can cause problems when comparing, filtering, or matching text in Excel.

By combining `TRIM`, `CLEAN`, and `SUBSTITUTE`, you can remove many of these unwanted characters in one go.

What does each part of the formula do?

- `TRIM` removes leading and trailing spaces and reduces multiple consecutive regular spaces to a single space.
- `CLEAN` removes non-printable characters with character codes 0–31, such as:
  - `CHAR(9)` — Tab
  - `CHAR(10)` — Line Feed (LF)
  - `CHAR(13)` — Carriage Return (CR)
- `SUBSTITUTE(A1; CHAR(160); " ")` replaces non-breaking spaces (`CHAR(160)`) with regular spaces. These are commonly introduced when copying text from websites or documents.

```excel
=TRIM(CLEAN(SUBSTITUTE(A1; CHAR(160); " ")))
```

If you use this formula regularly, consider turning it into a custom `LAMBDA` function using a named range.

Press **CTRL + F3** to open the **Name Manager**, select **New**, and enter:

**Name:** `CLEANTEXT`

**Refers to:**

```excel
=LAMBDA(
	txt;
	TRIM(CLEAN(SUBSTITUTE(txt; CHAR(160); " ")))
)
```

You can then use the function anywhere in the workbook:

```excel
=CLEANTEXT(A1)
```

## [NL] Tekst opschonen: verwijdert overtollige spaties en verborgen tekens

Wanneer je tekst kopieert uit een andere bron, zoals een PDF, website of extern bestand, kan de gekopieerde waarde ongewenste spaties of onzichtbare tekens bevatten. Deze kunnen problemen veroorzaken bij het vergelijken, filteren of opzoeken van tekst in Excel.

Door `SPATIES.WISSEN`, `WISSEN.CONTROL` en `SUBSTITUEREN` te combineren, kun je veel van deze ongewenste tekens in één keer verwijderen.

Wat doet elk onderdeel van de formule?

- `SPATIES.WISSEN` verwijdert spaties aan het begin en einde van de tekst en brengt meerdere opeenvolgende gewone spaties terug tot één spatie.
- `WISSEN.CONTROL` verwijdert niet-afdrukbare tekens met tekencodes 0–31, zoals:
  - `TEKEN(9)` — Tab
  - `TEKEN(10)` — Line Feed (LF)
  - `TEKEN(13)` — Carriage Return (CR)
- `SUBSTITUEREN(A1; TEKEN(160); " ")` vervangt vaste spaties of non-breaking spaces (`TEKEN(160)`) door gewone spaties. Deze komen vaak mee bij het kopiëren van tekst uit websites of documenten.

```excel
=SPATIES.WISSEN(WISSEN.CONTROL(SUBSTITUEREN(A1; TEKEN(160); " ")))
```

Als je deze formule regelmatig gebruikt, kun je er een aangepaste `LAMBDA`-functie van maken via een benoemd bereik.

Druk op **CTRL + F3** om **Namen beheren** te openen, kies **Nieuw** en vul het volgende in:

**Naam:** `CLEANTEXT`

**Verwijst naar:**

```excel
=LAMBDA(
	txt;
	SPATIES.WISSEN(WISSEN.CONTROL(SUBSTITUEREN(txt; TEKEN(160); " ")))
)
```

Daarna kun je de functie overal in de werkmap gebruiken:

```excel
=CLEANTEXT(A1)
```

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
