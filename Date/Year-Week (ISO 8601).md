## [EN] Year-Week (ISO 8601)

Excel does have a built-in function for returning the ISO 8601 weeknumber, but lacks a function to return the ISO 8601 year. Here is a simple solution to tackle that problem.

> [!NOTE]
> [Wikipedia: ISO 8601, Week dates](https://en.wikipedia.org/wiki/ISO_8601#Week_dates)

To return the ISO year and week in the format `2026-W53`:

```excel
=LET(
	date; A1;
	type; 2; correction; 4;
	isoYear; YEAR(date - WEEKDAY(date; type) + correction);
	isoWeek; TEXT(ISO.WEEKNUMBER(date);"-W00");
	isoYear & isoWeek;
)
```

If you use this formula regularly, consider turning it into a custom `LAMBDA` function using a named range.

Press **CTRL + F3** to open the **Name Manager**, select **New**, and enter:

**Name:** `ISO.YEARWEEK`

**Refers to:**

```excel
=LAMBDA(date;
	LET(
		type; 2; correction; 4;
		isoYear; YEAR(date - WEEKDAY(date; type) + correction);
		isoWeek; TEXT(ISO.WEEKNUMBER(date);"-W00");
		isoYear & isoWeek;
	)
)
```

You can then use the function anywhere in your workbook:

```excel
=ISO.YEARWEEK(A1)
```




````less
## [NL] Jaar-Week (ISO 8601)

Excel heeft een ingebouwde functie om het ISO 8601-weeknummer te bepalen, maar heeft geen functie om het bijbehorende ISO 8601-jaar te retourneren. Hieronder staat een eenvoudige oplossing voor dit probleem.

> [!NOTE]
> [Wikipedia: ISO 8601, Weekdatumnotatie](https://nl.wikipedia.org/wiki/ISO_8601#Weeknummering)

Om het ISO-jaar en weeknummer te retourneren in het formaat `2026-W53`:

```excel
=LET(
	datum; A1;
	type; 2; correctie; 4;
	isoJaar; JAAR(datum - WEEKDAG(datum; type) + correctie);
	isoWeek; TEKST(ISO.WEEKNUMMER(datum);"-W00");
	isoJaar & isoWeek;
)
```

Als je deze formule regelmatig gebruikt, kun je er een aangepaste `LAMBDA`-functie van maken via een benoemd bereik.

Druk op **CTRL + F3** om **Namen beheren** te openen, kies **Nieuw** en voer het volgende in:

**Naam:** `ISO.JAARWEEK`

**Verwijst naar:**

```excel
=LAMBDA(datum;
	LET(
		type; 2; correctie; 4;
		isoJaar; JAAR(datum - WEEKDAG(datum; type) + correctie);
		isoWeek; TEKST(ISO.WEEKNUMMER(datum);"-W00");
		isoJaar & isoWeek;
	)
)
```

Je kunt de functie vervolgens overal in je werkmap gebruiken:

```excel
=ISO.JAARWEEK(A1)
```
````

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
