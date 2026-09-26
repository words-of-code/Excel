## [EN] Cleanup text

When copying text from another file (e.g. PDF) or from a website, then the chances are that you get spaces on one or both sides around the text you need, or that you get non visible characters. You can get rid if these all easily with the following formula:

```excel
=TRIM(CLEAN(SUBSTITUTE(A1; CHAR(160); "")))
```

**TIP**: store this formula as a LAMBDA-function within a named range, e.g. CLEANUP, so you can use everywhere with ease. This way you can use it as a function: `=CLEANUP(A1)`.
Hit `CTRL + F3` to open the 'Manage names' window and click New. Name: CLEANUP. Points to:

```excel
=LAMBDA(
	txt;
	TRIM(CLEAN(SUBSTITUTE(txt; CHAR(160); "")))
)
```

### What does the formula parts do?
- TRIM: removes leading and trailing spaces, and reduces multiple spaces behind each other to a single space
- CLEAN: removes invisible characters (0 - 31), think about:
  - CHAR(9): tab
  - CHAR(10): Line Feed (LF)
  - CHAR(13): Carriage Return (CR)
- SUBSTITUTE CHAR(160): replaces Non-Breaking Space (NBSP) with a regular space

## [NL] Tekst opschonen

Wanneer je tekst kopieert vanuit een ander bestand (bijv PDF) of van een website, dan bestaat de kans dat er naast voorloop / naloop spaties ook niet zichtbare tekens meekomen. Hier kan je eenvoudig vanaf komen met de volgende formule:

```excel
=SPATIES.WISSEN(WISSEN.CONTROL(SUBSTITUEREN(A1; TEKEN(160); "")))
```

**TIP**: sla deze formule op als een benoemde LAMBDA-functie, bijv: OPSCHONEN. Hiermee kan je de formule aanroepen als een formule: `=OPSCHONEN(A1)`.
Open met `CTRL + F3` het 'Namen beheren' venster en klik op Nieuw. Naam: OPSCHONEN. Verwijst naar:
```excel
=LAMBDA(
	txt;
	SPATIES.WISSEN(WISSEN.CONTROL(SUBSTITUEREN(txt; TEKEN(160); "")))
)
```

### Wat doen de verschillende delen van de formule?
- SPATIES.WISSEN: spaties aan weerzijden verwijderen
- WISSEN.CONTROL: verwijder onzichtbare tekens (0 - 31), denk hierbij aan tekens als:
  - TEKEN(9) (tab)
  - TEKEN(10) (Line Feed / LF ofwel regeleinde)
  - TEKEN(13) (Carriage Return / CR ofwel harde return)
- SUBSTITUEREN TEKEN(160) (Non-Breaking Space / NBSP ofwel vast spatie) door niets

---

If you enjoy what I do, please consider supporting me. Especially when it solved an issue for you or just saved you time. Thank you!

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)
