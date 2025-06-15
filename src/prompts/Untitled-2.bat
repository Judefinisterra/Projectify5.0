<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH1; row1="(D)|~Recurring Revenue:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<SPREAD-E; indent="1"; bold="true"; topborder="false"; row1="V1(D)|New ARR Added Per Year(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$100000000(Y1)|$100000000(Y2)|$100000000(Y3)|$100000000(Y4)|$100000000(Y5)|$100000000(Y6)|">
<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH3; row1="(D)|~ARR:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<OFFSETCOLUMN-S; sumif="offsetyear"; driver1="R2"; indent="2"; bold="false"; topborder="false"; italic="true"; row1="V3(D)|~Beginning(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|~$F(Y1)|~$F(Y2)|~$F(Y3)|~$F(Y4)|~$F(Y5)|~$F(Y6)|">
<DIRECT-S; driver1="V1"; indent="2"; bold="false"; topborder="false"; italic="true"; sumif="year"; row1="V4(D)|~New(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|~$F(Y1)|~$F(Y2)|~$F(Y3)|~$F(Y4)|~$F(Y5)|~$F(Y6)|">
<SUBTOTAL2-S; driver1="V3"; driver2="V4"; bold="true"; italic="true"; sumif="yearend"; topborder="true"; indent="1"; row1="R2(D)|~Ending(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|~$F(Y1)|~$F(Y2)|~$F(Y3)|~$F(Y4)|~$F(Y5)|~$F(Y6)|">
<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<DEANNUALIZE-S; bold="true"; indent="1"; topborder="false"; driver1="R2"; row1="V5(D)|Revenue - Subscriptions(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">