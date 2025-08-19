## ISRC Validation
```excel
=LET(
  x, UPPER(C2),
  isAZ, LAMBDA(ch, AND(CODE(ch)>=65, CODE(ch)<=90)),
  is09, LAMBDA(ch, AND(CODE(ch)>=48, CODE(ch)<=57)),
  isAN, LAMBDA(ch, OR(isAZ(ch), is09(ch))),
  OR(
    x="NOT FOUND",
    AND(
      LEN(x)=12,
      isAZ(MID(x,1,1)), isAZ(MID(x,2,1)),
      isAN(MID(x,3,1)), isAN(MID(x,4,1)), isAN(MID(x,5,1)),
      is09(MID(x,6,1)), is09(MID(x,7,1)),
      is09(MID(x,8,1)), is09(MID(x,9,1)), is09(MID(x,10,1)), is09(MID(x,11,1)), is09(MID(x,12,1))
    )
  )
)
```