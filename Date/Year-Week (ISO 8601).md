# Intro

Excel ships with ISO.WEEKNUMBER() which is great for getting the ISO 8601 based weeknumber BUT it lacks an ISO.YEAR() function.

Luckily we can get around that with a small formula.

## [EN] Year-Week (ISO 8601)

```excel
=LET(
	date; A1;
	type; 2; correction; 4;
	isoYear; YEAR(date - WEEKDAY(date; type) + correction);
	isoWeek; TEXT(ISO.WEEKNUMBER(date);"-W00");
	isoYear & isoWeek;
)
```

## [NL] Jaar-Week (ISO 8601)

```excel
=LET(
	datum; A1;
	type; 2; correctie; 4;
	isoJaar; JAAR(datum - WEEKDAG(datum; type) + correctie);
	isoWeek; TEKST(ISO.WEEKNUMMER(datum);"-W00");
	isoJaar & isoWeek;
)
```
