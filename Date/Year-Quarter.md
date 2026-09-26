# Intro



## Year-Quarter [EN]

``
=LET(
	date; A1;
	quarter; CEILING(MONTH(date)/3; 1);
	CONCAT(YEAR(date); "-Q"; quarter);
)
``

## Jaar-Kwartaal [NL]

``
=LET(
	datum; A1;
	kwartaal; AFRONDEN.BOVEN(MAAND(datum)/3; 1);
	TEKST.SAMENV(JAAR(datum); "-Q"; kwartaal);
)
``