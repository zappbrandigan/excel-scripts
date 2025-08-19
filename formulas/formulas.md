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

*Paste Version*

```excel
=LET(x, UPPER(I2), isAZ, LAMBDA(ch, AND(CODE(ch)>=65, CODE(ch)<=90)), is09, LAMBDA(ch, AND(CODE(ch)>=48, CODE(ch)<=57)), isAN, LAMBDA(ch, OR(isAZ(ch), is09(ch))), OR( x="NOT FOUND", AND( LEN(x)=12, isAZ(MID(x,1,1)), isAZ(MID(x,2,1)), isAN(MID(x,3,1)), isAN(MID(x,4,1)), isAN(MID(x,5,1)), is09(MID(x,6,1)), is09(MID(x,7,1)), is09(MID(x,8,1)), is09(MID(x,9,1)), is09(MID(x,10,1)), is09(MID(x,11,1)), is09(MID(x,12,1)))))
```
*no LET/LAMBDA version*

```exel
=OR(UPPER(C2)="NOT FOUND", AND( LEN(C2)=12, AND(CODE(UPPER(MID(C2,1,1)))>=65,CODE(UPPER(MID(C2,1,1)))<=90), AND(CODE(UPPER(MID(C2,2,1)))>=65,CODE(UPPER(MID(C2,2,1)))<=90), OR(AND(CODE(UPPER(MID(C2,3,1)))>=65,CODE(UPPER(MID(C2,3,1)))<=90),AND(CODE(MID(C2,3,1))>=48,CODE(MID(C2,3,1))<=57)), OR(AND(CODE(UPPER(MID(C2,4,1)))>=65,CODE(UPPER(MID(C2,4,1)))<=90),AND(CODE(MID(C2,4,1))>=48,CODE(MID(C2,4,1))<=57)), OR(AND(CODE(UPPER(MID(C2,5,1)))>=65,CODE(UPPER(MID(C2,5,1)))<=90),AND(CODE(MID(C2,5,1))>=48,CODE(MID(C2,5,1))<=57)), AND(CODE(MID(C2,6,1))>=48,CODE(MID(C2,6,1))<=57), AND(CODE(MID(C2,7,1))>=48,CODE(MID(C2,7,1))<=57), AND(CODE(MID(C2,8,1))>=48,CODE(MID(C2,8,1))<=57), AND(CODE(MID(C2,9,1))>=48,CODE(MID(C2,9,1))<=57), AND(CODE(MID(C2,10,1))>=48,CODE(MID(C2,10,1))<=57), AND(CODE(MID(C2,11,1))>=48,CODE(MID(C2,11,1))<=57), AND(CODE(MID(C2,12,1))>=48,CODE(MID(C2,12,1))<=57)))
```