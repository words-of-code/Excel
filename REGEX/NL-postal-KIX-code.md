## [EN] Dutch postal (PostNL) KIX Code

This Excel `LAMBDA` function creates a Dutch KIX code from:

- a Dutch postal code;
- a house number;
- an optional house number addition.

The function normalizes the supplied values, validates them, and returns either:

- a valid KIX code; or
- a descriptive validation error.

The generated KIX value follows this structure:

```text
POSTCODE + HOUSE NUMBER + X + ADDITION
```

The `X` is only added when a house number addition is present.

Examples:

```text
1234 AB + 12      → 1234AB12
1234 AB + 12 A    → 1234AB12XA
1234 AB + 12 + A  → 1234AB12XA
```

---

### Function parameters

| Parameter | Description | Required |
|---|---|---:|
| `postcode` | Dutch postal code, with or without a space | Yes |
| `huisnummer` | House number, optionally including the addition | Yes |
| `toevoeging` | Separate house number addition | No |

The house number addition may therefore be supplied in either of these ways:

```text
huisnummer = 12A
toevoeging = ""
```

or:

```text
huisnummer = 12
toevoeging = A
```

Both result in:

```text
1234AB12XA
```

---

### Normalization

Before validation, the function normalizes the input.

#### Postal code

The postal code is:

- converted to uppercase;
- trimmed;
- stripped of spaces.

Example:

```text
" 1234 ab " → "1234AB"
```

#### House number

The house number is:

- converted to uppercase;
- trimmed.

The numeric part is extracted separately from a possible addition.

Examples:

```text
"12"       → number: 12
"12A"      → number: 12, addition: A
"12-A"     → number: 12, addition: A
"12 A"     → number: 12, addition: A
```

#### Addition

The addition is:

- converted to uppercase;
- trimmed;
- stripped of spaces and hyphens.

Examples:

```text
"a"        → "A"
"A-1"      → "A1"
"A 1"      → "A1"
"A-B-C"    → "ABC"
```

The normalized addition may contain a maximum of **6 alphanumeric characters**.

---

### Validation rules

#### Postal code

A postal code must:

- start with a digit from `1` through `9`;
- contain exactly 4 digits;
- end with exactly 2 letters;
- not use `SA`, `SD`, or `SS` as the letter combination.

Valid examples:

```text
1234AB
1234 AB
9999ZZ
```

Invalid examples:

```text
0123AB
123AB
12345AB
1234A
1234ABC
1234SA
1234SD
1234SS
```

#### House number

The numerical house number must contain:

- at least 1 digit;
- at most 5 digits.

A house number may also contain an addition.

Examples:

```text
1
12
12345
12A
12-A
12 A
```

A 6-digit house number is invalid.

#### House number addition

After normalization, an addition must:

- contain only letters `A-Z` and digits `0-9`;
- contain no more than 6 characters.

Spaces and hyphens do not count towards this maximum because they are removed during normalization.

For example:

```text
A-B-C
```

is normalized to:

```text
ABC
```

and is therefore valid.

---

### Valid and invalid examples

| Postcode | House number | Addition | Normalized KIX / Result | Valid |
|---|---:|---|---|:---:|
| `1234 AB` | `12` |  | `1234AB12` | ✓ |
| `1234 ab` | `12` |  | `1234AB12` | ✓ |
| `1234 AB` | `12A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12-A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12 A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12` | `A` | `1234AB12XA` | ✓ |
| `1234 AB` | `12` | `A-1` | `1234AB12XA1` | ✓ |
| `1234 AB` | `12 A-B-C` |  | `1234AB12XABC` | ✓ |
| `1234 AB` | `12` | `ABCDEF` | `1234AB12XABCDEF` | ✓ |
| `1234 AB` | `12345` |  | `1234AB12345` | ✓ |
| `0123 AB` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SA` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SD` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SS` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 AB` | `123456` |  | `Ongeldig huisnummer of toevoeging!` | ✗ |
| `1234 AB` | `12 ABC-DEFG` |  | `Ongeldige toevoeging!` | ✗ |
| `1234 AB` | `12` | `ABCDEFG` | `Ongeldige toevoeging!` | ✗ |

> `ABC-DEFG` is normalized to `ABCDEFG`. Because that contains 7 characters, the addition is invalid.

---

### Combination of additions

The function supports an addition supplied:

- as part of `huisnummer`;
- through the separate `toevoeging` parameter;
- or through both.

The current logic behaves as follows:

| House number | Addition parameter | Effective addition |
|---|---|---|
| `12` | empty | none |
| `12A` | empty | `A` |
| `12` | `A` | `A` |
| `12A` | `A` | `A` |
| `12A` | `B` | `AB` |

When both values are present and different, they are concatenated.

---

### Validation messages

The function may return the following error messages:

| Message | Meaning |
|---|---|
| `Ongeldige postcode!` | The postal code does not meet the validation rules |
| `Ongeldig huisnummer of toevoeging!` | The house number or embedded addition has an invalid basic format |
| `Ongeldige toevoeging!` | The normalized addition contains invalid characters or more than 6 characters |

### Usage

After saving the formula as a named Excel function, for example:

```text
KIX
```

it can be used as:

```excel
=KIX(A1; B1; C1)
```

where:

- `A1` contains the postal code;
- `B1` contains the house number;
- `C1` contains the optional addition.

Example:

```text
A1 = 1234 AB
B1 = 12-A
C1 =

Result: 1234AB12XA
```

## [NL] PostNL KIX code

Deze Excel-`LAMBDA`-functie maakt een Nederlandse KIX-code op basis van:

- een Nederlandse postcode;
- een huisnummer;
- een optionele huisnummertoevoeging.

De functie normaliseert eerst de aangeleverde waarden, controleert deze vervolgens en retourneert daarna:

- een geldige KIX-code; of
- een duidelijke foutmelding.

De KIX-code wordt opgebouwd volgens:

```text
POSTCODE + HUISNUMMER + X + TOEVOEGING
```

De `X` wordt alleen toegevoegd wanneer er een huisnummertoevoeging aanwezig is.

Voorbeelden:

```text
1234 AB + 12      → 1234AB12
1234 AB + 12 A    → 1234AB12XA
1234 AB + 12 + A  → 1234AB12XA
```

---

### Parameters

| Parameter | Omschrijving | Verplicht |
|---|---|---:|
| `postcode` | Nederlandse postcode, met of zonder spatie | Ja |
| `huisnummer` | Huisnummer, eventueel inclusief toevoeging | Ja |
| `toevoeging` | Losse huisnummertoevoeging | Nee |

De huisnummertoevoeging mag dus op beide manieren worden aangeleverd:

```text
huisnummer = 12A
toevoeging = ""
```

of:

```text
huisnummer = 12
toevoeging = A
```

Beide leveren op:

```text
1234AB12XA
```

---

### Normalisatie

Voor de validatie worden de aangeleverde waarden eerst genormaliseerd.

#### Postcode

De postcode wordt:

- omgezet naar hoofdletters;
- ontdaan van overtollige spaties;
- volledig ontdaan van spaties.

Voorbeeld:

```text
" 1234 ab " → "1234AB"
```

#### Huisnummer

Het huisnummer wordt:

- omgezet naar hoofdletters;
- ontdaan van overtollige spaties.

Het numerieke huisnummer wordt vervolgens losgetrokken van een eventueel aanwezige toevoeging.

Voorbeelden:

```text
"12"       → nummer: 12
"12A"      → nummer: 12, toevoeging: A
"12-A"     → nummer: 12, toevoeging: A
"12 A"     → nummer: 12, toevoeging: A
```

#### Toevoeging

Een toevoeging wordt:

- omgezet naar hoofdletters;
- ontdaan van overtollige spaties;
- ontdaan van spaties en koppeltekens.

Voorbeelden:

```text
"a"        → "A"
"A-1"      → "A1"
"A 1"      → "A1"
"A-B-C"    → "ABC"
```

De **genormaliseerde toevoeging** mag maximaal **6 alfanumerieke karakters** bevatten.

---

### Validatieregels

#### Postcode

Een postcode moet:

- beginnen met een cijfer van `1` t/m `9`;
- exact 4 cijfers bevatten;
- eindigen op exact 2 letters;
- niet eindigen op `SA`, `SD` of `SS`.

Geldige voorbeelden:

```text
1234AB
1234 AB
9999ZZ
```

Ongeldige voorbeelden:

```text
0123AB
123AB
12345AB
1234A
1234ABC
1234SA
1234SD
1234SS
```

#### Huisnummer

Het numerieke huisnummer moet bestaan uit:

- minimaal 1 cijfer;
- maximaal 5 cijfers.

Achter het huisnummer mag ook direct een toevoeging staan.

Voorbeelden:

```text
1
12
12345
12A
12-A
12 A
```

Een huisnummer van 6 cijfers is ongeldig.

#### Huisnummertoevoeging

Na normalisatie mag een toevoeging:

- uitsluitend de letters `A-Z` en cijfers `0-9` bevatten;
- maximaal 6 karakters lang zijn.

Spaties en koppeltekens tellen niet mee voor deze maximale lengte, omdat deze tijdens de normalisatie worden verwijderd.

Bijvoorbeeld:

```text
A-B-C
```

wordt:

```text
ABC
```

en is dus geldig.

---

### Geldige en ongeldige voorbeelden

| Postcode | Huisnummer | Toevoeging | KIX / resultaat | Geldig |
|---|---:|---|---|:---:|
| `1234 AB` | `12` |  | `1234AB12` | ✓ |
| `1234 ab` | `12` |  | `1234AB12` | ✓ |
| `1234 AB` | `12A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12-A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12 A` |  | `1234AB12XA` | ✓ |
| `1234 AB` | `12` | `A` | `1234AB12XA` | ✓ |
| `1234 AB` | `12` | `A-1` | `1234AB12XA1` | ✓ |
| `1234 AB` | `12 A-B-C` |  | `1234AB12XABC` | ✓ |
| `1234 AB` | `12` | `ABCDEF` | `1234AB12XABCDEF` | ✓ |
| `1234 AB` | `12345` |  | `1234AB12345` | ✓ |
| `0123 AB` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SA` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SD` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 SS` | `12` |  | `Ongeldige postcode!` | ✗ |
| `1234 AB` | `123456` |  | `Ongeldig huisnummer of toevoeging!` | ✗ |
| `1234 AB` | `12 ABC-DEFG` |  | `Ongeldige toevoeging!` | ✗ |
| `1234 AB` | `12` | `ABCDEFG` | `Ongeldige toevoeging!` | ✗ |

> `ABC-DEFG` wordt genormaliseerd naar `ABCDEFG`. Omdat dit 7 karakters bevat, is de toevoeging ongeldig.

---

### Combineren van toevoegingen

De functie ondersteunt een toevoeging die:

- onderdeel is van `huisnummer`;
- via de losse parameter `toevoeging` wordt aangeleverd;
- of via beide wordt aangeleverd.

De huidige logica werkt als volgt:

| Huisnummer | Parameter toevoeging | Gebruikte toevoeging |
|---|---|---|
| `12` | leeg | geen |
| `12A` | leeg | `A` |
| `12` | `A` | `A` |
| `12A` | `A` | `A` |
| `12A` | `B` | `AB` |

Wanneer beide toevoegingen aanwezig zijn maar niet gelijk zijn, worden ze samengevoegd.

---

### Foutmeldingen

De functie kan de volgende foutmeldingen retourneren:

| Melding | Betekenis |
|---|---|
| `Ongeldige postcode!` | De postcode voldoet niet aan de validatieregels |
| `Ongeldig huisnummer of toevoeging!` | Het huisnummer of de daarin opgenomen toevoeging heeft geen geldig basisformaat |
| `Ongeldige toevoeging!` | De genormaliseerde toevoeging bevat ongeldige tekens of meer dan 6 karakters |

---

## Excel LAMBDA

```excel
=LAMBDA(postcode; huisnummer; toevoeging;
    LET(
        pc; HOOFDLETTERS( SUBSTITUEREN( SPATIES.WISSEN( postcode & "" );" ";"" ));
        hn; HOOFDLETTERS( SPATIES.WISSEN( huisnummer & "" ));
        tv; REGEXVERVANGEN( HOOFDLETTERS( SPATIES.WISSEN( toevoeging & "" )); "[ -]"; "" );

        pc_geldig; ALS.FOUT(
            REGEXTEST( pc; "^[1-9][0-9]{3}(?!SA|SD|SS)[A-Z]{2}$"; 1 );
            ONWAAR
        );

        hn_geldig; ALS.FOUT(
            REGEXTEST( hn; "^[0-9]{1,5}(?:[ -]?[A-Z0-9][A-Z0-9 -]*)?$"; 1 );
            ONWAAR
        );

        nr; ALS( hn_geldig; REGEXEXTRAHEREN( hn; "^[0-9]{1,5}"); "" );

        tv_hn; ALS(
            hn_geldig;
            ALS.FOUT(
                REGEXVERVANGEN(
                    REGEXVERVANGEN( hn; "^[0-9]{1,5}"; "" );
                    "[ -]";
                    ""
                );
                ""
            );
            ""
        );

        tv_hn_geldig; OF(
            tv_hn="";
            ALS.FOUT(
                REGEXTEST( tv_hn; "^[A-Z0-9]{1,6}$"; 1 );
                ONWAAR
            )
        );

        tv_geldig; OF(
            tv="";
            ALS.FOUT(
                REGEXTEST( tv; "^[A-Z0-9]{1,6}$"; 1 );
                ONWAAR
            )
        );

        tv_def; ALS(
            tv_hn="";
            tv;
            ALS(
                tv="";
                tv_hn;
                ALS( tv_hn=tv; tv_hn; tv_hn & tv )
            )
        );

        SCHAKELEN(
            WAAR;
            NIET(pc_geldig); "Ongeldige postcode!";
            NIET(hn_geldig); "Ongeldig huisnummer of toevoeging!";
            NIET(tv_hn_geldig); "Ongeldige toevoeging!";
            NIET(tv_geldig); "Ongeldige toevoeging!";
            pc & nr & ALS( tv_def<>""; "X" & tv_def; "" )
        )
    )
)
```

---

## Gebruik

Nadat de formule als benoemde Excel-functie is opgeslagen, bijvoorbeeld als:

```text
KIX
```

kan deze als volgt worden gebruikt:

```excel
=KIX(A1; B1; C1)
```

waarbij:

- `A1` de postcode bevat;
- `B1` het huisnummer bevat;
- `C1` de optionele toevoeging bevat.

Voorbeeld:

```text
A1 = 1234 AB
B1 = 12-A
C1 =

Resultaat: 1234AB12XA
```
