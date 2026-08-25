function refreshFormulas() {
  setPromoFormulas_(); setBtoFormulas_(); setNonPromoFormulas_(); setMonthlyFormulas_();
}

function putArray_(sheet, cell, formula) { sheet.getRange(cell).setFormula(formula); }

function setPromoFormulas_() {
  const s=mbSheet_('Promos');
  putArray_(s,'A2','=ARRAYFORMULA(IF(B2:B="","","P-"&TEXT(B2:B,"yyyymmdd")&"-"&TEXT(ROW(B2:B)-1,"0000")))');
  putArray_(s,'I2','=ARRAYFORMULA(IF(D2:D="","",IFNA(VLOOKUP(D2:D,Settings!A2:B,2,FALSE),NA())))');
  putArray_(s,'J2','=ARRAYFORMULA(IF(F2:F="","",F2:F*I2:I))');
  putArray_(s,'M2','=ARRAYFORMULA(IF((B2:B="")+(F2:F=""),"",N(K2:K)+N(L2:L)))');
  putArray_(s,'N2','=ARRAYFORMULA(IF(F2:F="","",M2:M/F2:F))');
  putArray_(s,'O2','=ARRAYFORMULA(IF(I2:I="","",IFNA(VLOOKUP(I2:I,SORT({Settings!E2:E7,Settings!D2:D7},1,TRUE),2,TRUE),NA())))');
  putArray_(s,'P2','=ARRAYFORMULA(IF(B2:B="","",DATE(YEAR(B2:B),MONTH(B2:B),1)))');
}

function setBtoFormulas_() {
  const s=mbSheet_('BTO');
  putArray_(s,'A2','=ARRAYFORMULA(IF(B2:B="","","B-"&TEXT(B2:B,"yyyymmdd")&"-"&TEXT(ROW(B2:B)-1,"0000")))');
  putArray_(s,'J2','=ARRAYFORMULA(IF((F2:F="")+(G2:G=""),"",IF(H2:H="",F2:F*Settings!E10,F2:F*(G2:G-1)/H2:H)))');
  putArray_(s,'K2','=ARRAYFORMULA(IF(F2:F="","",J2:J/F2:F))');
  putArray_(s,'Q2','=ARRAYFORMULA(IF((B2:B="")+(F2:F=""),"",N(L2:L)+N(P2:P)))');
  putArray_(s,'R2','=ARRAYFORMULA(IF(F2:F="","",Q2:Q/F2:F))');
  putArray_(s,'S2','=ARRAYFORMULA(IF(B2:B="","",DATE(YEAR(B2:B),MONTH(B2:B),1)))');
}

function setNonPromoFormulas_() {
  const s=mbSheet_('NonPromos');
  putArray_(s,'A2','=ARRAYFORMULA(IF(B2:B="","","N-"&TEXT(B2:B,"yyyymmdd")&"-"&TEXT(ROW(B2:B)-1,"0000")))');
  putArray_(s,'J2','=ARRAYFORMULA(IF((F2:F="")+(H2:H=""),"",F2:F*((G2:G/H2:H)-1)))');
  putArray_(s,'K2','=ARRAYFORMULA(IF(F2:F="","",IF(J2:J="","",J2:J/F2:F)))');
  putArray_(s,'N2','=ARRAYFORMULA(IF((B2:B="")+(F2:F=""),"",N(L2:L)+N(M2:M)))');
  putArray_(s,'O2','=ARRAYFORMULA(IF(F2:F="","",N2:N/F2:F))');
  putArray_(s,'P2','=ARRAYFORMULA(IF(B2:B="","",DATE(YEAR(B2:B),MONTH(B2:B),1)))');
}

function setMonthlyFormulas_() {
  const s=mbSheet_('Monthly Metrics'); mbEnsureSize_(s,500,20);
  s.getRange(1,1,1,20).setValues([MB.MONTH_HEADERS]);
  putArray_(s,'A2','=IFNA(SORT(UNIQUE(FILTER({Promos!P2:P;BTO!S2:S;NonPromos!P2:P},{Promos!P2:P;BTO!S2:S;NonPromos!P2:P}<>""))),"")');
  const fs = [
    '=ARRAYFORMULA(IF(A2:A="","",SUMIF(Promos!P:P,A2:A,Promos!F:F)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(Promos!P:P,A2:A,Promos!J:J)))','=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",C2:C/B2:B)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(Promos!P:P,A2:A,Promos!M:M)))','=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",E2:E/B2:B)))',
    '=ARRAYFORMULA(IF(A2:A="","",SUMIF(BTO!S:S,A2:A,BTO!F:F)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(BTO!S:S,A2:A,BTO!J:J)))','=ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",H2:H/G2:G)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(BTO!S:S,A2:A,BTO!Q:Q)))','=ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",J2:J/G2:G)))',
    '=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!P:P,A2:A,NonPromos!F:F)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!P:P,A2:A,NonPromos!J:J)))','=ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",M2:M/L2:L)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!P:P,A2:A,NonPromos!N:N)))','=ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",O2:O/L2:L)))','=ARRAYFORMULA(IF(A2:A="","",C2:C+H2:H+M2:M))','=ARRAYFORMULA(IF(A2:A="","",E2:E+J2:J+O2:O))','=ARRAYFORMULA(IF(A2:A="","",SCAN(0,Q2:Q,LAMBDA(a,v,a+v))))','=ARRAYFORMULA(IF(A2:A="","",SCAN(0,R2:R,LAMBDA(a,v,a+v))))'];
  fs.forEach((f,i)=>putArray_(s,String.fromCharCode(66+i)+'2',f));
}
