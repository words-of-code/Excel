## [NL] Valideer BSN nummer via elfproef

```excel
=LET(
  bsn; REGEXVERVANGEN(
    REGEXEXTRAHEREN(A1;"(?<![0-9])(?:BSN)?[0-9]{8,9}(?![0-9])");
    "^BSN";
    ""
  );
  n; TEKST(bsn; "000000000");
  REST(
    9*WAARDE(DEEL(n;1;1)) + 8*WAARDE(DEEL(n;2;1)) + 7*WAARDE(DEEL(n;3;1))+ 6*WAARDE(DEEL(n;4;1)) +
    5*WAARDE(DEEL(n;5;1)) + 4*WAARDE(DEEL(n;6;1)) + 3*WAARDE(DEEL(n;7;1)) + 2*WAARDE(DEEL(n;8;1)) – 
    WAARDE(DEEL(n;9;1));
    11
  )=0
)
```
