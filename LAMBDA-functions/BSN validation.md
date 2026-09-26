## [NL] Valideer BSN nummer via elfproef

BSN nummers zijn te controleren via de elfproef. Met deze proef kan vastgesteld worden of het een geldig nummer is.

BSN nummers hebben een vaste lengte van 9 cijfers, maar soms wordt de voorloop nul weggelaten waardoor er slechts 8 cijfers zijn.
De formule houdt hier rekening mee. 

Verder verwijderd de formule ook `BSN` en `BSN:` (eventueel gevolgd door een spatie) als dat ervoor staat. Hiermee is de formule ook in te zetten zonder eerst die tekst te moeten verwijderen.

```excel
=LET(
  bsn; REGEXVERVANGEN(
    REGEXEXTRAHEREN(A1; "(?<![0-9])(?:BSN)?[0-9]{8,9}(?![0-9])");
    "^BSN";
    ""
  );
  n; TEKST(bsn; "000000000");
  REST(
    9*WAARDE(DEEL(n;1;1)) + 8*WAARDE(DEEL(n;2;1)) + 7*WAARDE(DEEL(n;3;1)) +
    6*WAARDE(DEEL(n;4;1)) + 5*WAARDE(DEEL(n;5;1)) + 4*WAARDE(DEEL(n;6;1)) +
    3*WAARDE(DEEL(n;7;1)) + 2*WAARDE(DEEL(n;8;1)) - WAARDE(DEEL(n;9;1));
    11
  )=0
)
```

Mocht je binnen een bestand veel gebruik van deze formule maken, overweeg dan om hem als LAMBDA functie te gebruiken.

Ga daarvoor via **CTRL + F3** naar **Namen beheren** en klik op **Nieuw**. Gebruik verder:
- Naam: IS.BSN
- Bereik: Werkmap
- Verwijst naar:
  ```excel
  =LAMBDA(txt;
  LET(
      bsn; REGEXVERVANGEN(
          REGEXEXTRAHEREN(txt; "(?<![0-9])(?:BSN)?[0-9]{8,9}(?![0-9])");
          "^BSN";
          ""
      );
      n; TEKST(bsn; "000000000");
      REST(
          9*WAARDE(DEEL(n;1;1)) + 8*WAARDE(DEEL(n;2;1)) + 7*WAARDE(DEEL(n;3;1)) +
          6*WAARDE(DEEL(n;4;1)) + 5*WAARDE(DEEL(n;5;1)) + 4*WAARDE(DEEL(n;6;1)) +
          3*WAARDE(DEEL(n;7;1)) + 2*WAARDE(DEEL(n;8;1)) - WAARDE(DEEL(n;9;1));
          11
      )=0
  ) )
  ```

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
