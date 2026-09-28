## [EN] Validate a BSN number using the eleven test

A BSN can be checked using the **eleven test**. This lets you determine whether a number complies with the validation rule used for BSN numbers. This does not necessarily mean that the BSN actually exists or has been issued.

A BSN consists of **9 digits**. When the first digit is `0`, this leading zero is sometimes omitted, resulting in an 8-digit value. The formula takes this into account and automatically adds the missing leading zero.

Using `REGEXEXTRACT`, the formula searches the supplied text for a sequence of **8 or 9 digits** that is not directly part of a longer sequence of digits. This means the formula can recognize both a standalone BSN and a BSN embedded in text, for example:

- `123456782`
- `BSN 123456782`
- `BSN: 123456782`

Any surrounding text therefore does not have to be removed manually first.

The formula returns **TRUE** when the extracted number passes the eleven test and **FALSE** when it does not.

```excel
=LET(
    bsn; REGEXEXTRACT(A1; "(?<![0-9])[0-9]{8,9}(?![0-9])");
    n; TEXT(bsn; "000000000");
    MOD(
        9*VALUE(MID(n;1;1)) + 8*VALUE(MID(n;2;1)) + 7*VALUE(MID(n;3;1)) +
        6*VALUE(MID(n;4;1)) + 5*VALUE(MID(n;5;1)) + 4*VALUE(MID(n;6;1)) +
        3*VALUE(MID(n;7;1)) + 2*VALUE(MID(n;8;1)) - VALUE(MID(n;9;1));
        11
    )=0
)
```

If you use this formula regularly, consider turning it into a custom `LAMBDA` function using a named range.

Press **CTRL + F3** to open the **Name Manager**, select **New**, and enter:

**Name:** `IS.BSN`

**Refers to:**

```excel
=LAMBDA(txt;
    LET(
        bsn; REGEXEXTRACT(txt; "(?<![0-9])[0-9]{8,9}(?![0-9])");
        n; TEXT(bsn; "000000000");
        MOD(
            9*VALUE(MID(n;1;1)) + 8*VALUE(MID(n;2;1)) + 7*VALUE(MID(n;3;1)) +
            6*VALUE(MID(n;4;1)) + 5*VALUE(MID(n;5;1)) + 4*VALUE(MID(n;6;1)) +
            3*VALUE(MID(n;7;1)) + 2*VALUE(MID(n;8;1)) - VALUE(MID(n;9;1));
            11
        )=0
    )
)
```

You can then use the function anywhere in the workbook:

```excel
=IS.BSN(A1)
```

## [NL] Valideer BSN nummer via elfproef

Een BSN kan met de **elfproef** worden gecontroleerd. Hiermee kun je vaststellen of een nummer voldoet aan de controle-regel die voor BSN-nummers wordt gebruikt. Dit betekent niet automatisch dat het BSN daadwerkelijk bestaat of is uitgegeven.

Een BSN bestaat uit **9 cijfers**. Wanneer het eerste cijfer een `0` is, wordt deze voorloopnul soms weggelaten en bestaat de ingevoerde waarde daardoor uit slechts 8 cijfers. De formule houdt hier rekening mee en vult de ontbrekende voorloopnul automatisch aan.

Met REGEXEXTRAHEREN zoekt de formule in de opgegeven tekst naar een reeks van **8 of 9 cijfers** die niet direct onderdeel is van een langere cijferreeks. Hierdoor kan de formule zowel een los BSN als een BSN in tekst herkennen, bijvoorbeeld:

- `123456782`
- `BSN 123456782`
- `BSN: 123456782`

Eventuele tekst rondom het nummer hoeft daardoor niet eerst handmatig te worden verwijderd.

De formule geeft **WAAR** terug wanneer het gevonden nummer aan de elfproef voldoet en **ONWAAR** wanneer dit niet het geval is.

```excel
=LET(
	bsn; REGEXEXTRAHEREN(A1; "(?<![0-9])[0-9]{8,9}(?![0-9])");
    n; TEKST(bsn; "000000000");
    REST(
        9*WAARDE(DEEL(n;1;1)) + 8*WAARDE(DEEL(n;2;1)) + 7*WAARDE(DEEL(n;3;1)) +
        6*WAARDE(DEEL(n;4;1)) + 5*WAARDE(DEEL(n;5;1)) + 4*WAARDE(DEEL(n;6;1)) +
        3*WAARDE(DEEL(n;7;1)) + 2*WAARDE(DEEL(n;8;1)) - WAARDE(DEEL(n;9;1));
        11
    )=0
)
```


Als je deze formule regelmatig gebruikt, kun je er een aangepaste `LAMBDA`-functie van maken via een benoemd bereik.

Druk op **CTRL + F3** om **Namen beheren** te openen, kies **Nieuw** en vul het volgende in:

**Naam:** `IS.BSN`

**Verwijst naar:**

```excel
=LAMBDA(txt;
	LET(
		bsn; REGEXEXTRAHEREN(txt; "(?<![0-9])[0-9]{8,9}(?![0-9])");
		n; TEKST(bsn; "000000000");
		REST(
			9*WAARDE(DEEL(n;1;1)) + 8*WAARDE(DEEL(n;2;1)) + 7*WAARDE(DEEL(n;3;1)) +
			6*WAARDE(DEEL(n;4;1)) + 5*WAARDE(DEEL(n;5;1)) + 4*WAARDE(DEEL(n;6;1)) +
			3*WAARDE(DEEL(n;7;1)) + 2*WAARDE(DEEL(n;8;1)) - WAARDE(DEEL(n;9;1));
			11
		)=0
	)
)
```

Daarna kun je de functie overal in de werkmap gebruiken:

```excel
=IS.BSN(A1)
```

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
