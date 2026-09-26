## Year-Quarter [EN]

```excel
=LET(
	date; A1;
	quarter; CEILING(MONTH(date)/3; 1);
	CONCAT(YEAR(date); "-Q"; quarter);
)
```

## Jaar-Kwartaal [NL]

```excel
=LET(
	datum; A1;
	kwartaal; AFRONDEN.BOVEN(MAAND(datum)/3; 1);
	TEKST.SAMENV(JAAR(datum); "-Q"; kwartaal);
)
```
