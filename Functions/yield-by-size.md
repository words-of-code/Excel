# Yield By (Bottle) Size

Documentation will follow.

`YIELDBYSIZE(bottles; shares; [full_size_ml]`


```excel
=LAMBDA(bottles; shares; [full_size_ml];
  LET(
    size; OR( ISOMITTED(full_size_ml); TRIM(full_size_ml)=""; full_size_ml=0 );
    IF(
      OR(
        AND( size<>700; size<>500 );
        shares<=0;
        MOD(shares; 1)<>0;
        bottles<0
      );
      "Invalid input";
      LET(
        fraction; MOD(ROUND(bottles/shares; 10); 1);
        small_size; IF(
          size=700;
          IF(fraction>=0,7; 500; IF(fraction>=0,5; 350; 0));
          IF(fraction>=0,7; 350; 0)
        );
		small_count; IF( small_size=0; 0; IF(fraction>=0,7; shares; CEILING.MATH(shares;2)) );
        remaining_ml; bottles * size-small_count * small_size;
        full_count;INT(ROUND(remaining_ml/size; 10));
        IF(
          remaining_ml < 0;
          "Insufficient volume";
          full_count & "x" & size &
          IF(small_count > 0;" + " & small_count & "x" & small_size;"")
        )
      )
    )
  )
)
```
