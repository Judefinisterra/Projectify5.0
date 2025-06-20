<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH1; row1="(D)|~Recurring Revenue:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<SPREAD-E; indent="1"; bold="true"; topborder="false"; row1="V1(D)|COGS(L)|is: direct costs(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|-$100000000(Y1)|-$200000000(Y2)|-$300000000(Y3)|-$400000000(Y4)|-$100000000(Y5)|-$100000000(Y6)|">
<CONST-E; indent="2"; bold="false"; topborder="false"; row1="V2(D)|~# of Days - Inventory Purchased (+- sale date)(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|~180(Y1)|~180(Y2)|~180(Y3)|~180(Y4)|~180(Y5)|~180(Y6)|">
<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH3; row1="(D)|~Inventory:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<OFFSETCOLUMN-S; sumif="offsetyear"; driver1="R2"; indent="2"; bold="false"; topborder="false"; italic="true"; row1="V3(D)|Beginning(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
<FORMULA-S; customformula="abs(offset(rd{V1},0,roundup(rd{V2}/30,0)))[30 days per month]"; sumif="year"; driver1="V1"; driver2="V2"; bold="false";  sumif="year"; topborder="false"; indent="2"; row1="Y2(D)|Plus: New Inventory Purchased(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
<FORMULA-S; customformula="-min(abs(rd{V1}),rd{Y2}+rd{V3})"; sumif="year"; driver1="V1"; indent="2"; bold="false"; topborder="false"; italic="true"; row1="V4(D)|Less: COGS(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
<SUM2-S; driver1="V3"; driver2="V4"; sumif="yearend"; bold="true";  sumif="year"; topborder="true"; indent="1"; row1="Z2(D)|Inventory(L)|bs: current assets(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
<BR; row1="(D)|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<CHANGE-S; driver1="Z2"; sumif="year"; bold="true";  sumif="year"; topborder="false"; indent="1"; row1="R2(D)|Change in Inventory(L)|cf: wc(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">