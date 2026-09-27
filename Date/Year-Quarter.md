## [EN] Year-Quarter

Excel doesn't have a built-in function for returning the quarter of a date, so here's a simple solution.

In the examples below, I've included the year as well. If you only need the quarter number, you can simply use:

```excel
=CEILING(MONTH(A1)/3; 1)
```

To return the year and quarter in the format `2026-Q3`:

```excel
=LET(
	date; A1;
	quarter; CEILING(MONTH(date)/3; 1);
	CONCAT(YEAR(date); "-Q"; quarter)
)
```

If you use this formula regularly, consider turning it into a custom `LAMBDA` function using a named range.

Press **CTRL + F3** to open the **Name Manager**, select **New**, and enter:

**Name:** `YEAR.Q`

**Refers to:**

```excel
=LAMBDA(date;
	LET(
		quarter; CEILING(MONTH(date)/3; 1);
		CONCAT(YEAR(date); "-Q"; quarter)
	)
)
```

You can then use the function anywhere in the workbook:

```excel
=YEAR.Q(A1)
```


## [NL] Jaar-Kwartaal

Excel heeft geen ingebouwde functie om het kwartaal van een datum terug te geven, dus hieronder een eenvoudige oplossing.

In de voorbeelden hieronder heb ik ook het jaar toegevoegd. Als je alleen het kwartaalnummer nodig hebt, kun je simpelweg dit gebruiken:

```excel
=AFRONDEN.BOVEN(MAAND(A1)/3; 1)
```

Om het jaar en kwartaal terug te geven in het formaat `2026-Q3`:

```excel
=LET(
	datum; A1;
	kwartaal; AFRONDEN.BOVEN(MAAND(datum)/3; 1);
	TEKST.SAMENV(JAAR(datum); "-Q"; kwartaal);
)
```

Als je deze formule regelmatig gebruikt, kun je er een aangepaste `LAMBDA`-functie van maken via een benoemd bereik.

Druk op **CTRL + F3** om **Namen beheren** te openen, kies **Nieuw** en vul het volgende in:

**Naam:** `JAAR.KW`

**Verwijst naar:**

```excel
=LAMBDA(datum;
	LET(
		kwartaal; AFRONDEN.BOVEN(MAAND(datum)/3; 1);
		TEKST.SAMENVOEGEN(JAAR(datum); "-Q"; kwartaal)
	)
)
```

Daarna kun je de functie overal in de werkmap gebruiken:

```excel
=JAAR.KW(A1)
```

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
