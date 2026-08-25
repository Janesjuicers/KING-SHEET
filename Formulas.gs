function refreshFormulas() {
  setPromoFormulas_(); setBtoFormulas_(); setNonPromoFormulas_(); setMonthlyFormulas_();
}

function putArray_(sheet, cell, formula) { sheet.getRange(cell).setFormula(formula); }

function setPromoFormulas_() {
  const s=mbSheet_('Promos');
  putArray_(s,'H2','=ARRAYFORMULA(IF(E2:E="","",E2:E*IFNA(VLOOKUP(C2:C,Settings!A2:B,2,FALSE),NA())))');
  putArray_(s,'I2','=ARRAYFORMULA(IF(G2:G="","",IF(G2:G="W",E2:E*(F2:F-1),IF((G2:G="L")+(G2:G="B"),-E2:E,IF(G2:G="Void",0,"")))))');
  putArray_(s,'J2','=ARRAYFORMULA(IF(G2:G="","",IF(G2:G="B",E2:E,IF((G2:G="W")+(G2:G="L")+(G2:G="Void"),0,""))))');
}

function setBtoFormulas_() {
  const s=mbSheet_('BTO');
  putArray_(s,'I2','=ARRAYFORMULA(IF((E2:E="")+(F2:F=""),"",IF(G2:G="",E2:E*Settings!E10,E2:E*(F2:F-1)/G2:G)))');
  putArray_(s,'J2','=ARRAYFORMULA(IF(E2:E="","",I2:I/E2:E))');
  putArray_(s,'P2','=ARRAYFORMULA(IF((A2:A="")+(E2:E=""),"",N(K2:K)+N(O2:O)))');
  putArray_(s,'Q2','=ARRAYFORMULA(IF(E2:E="","",P2:P/E2:E))');
}

function setNonPromoFormulas_() {
  const s=mbSheet_('NonPromos');
  putArray_(s,'I2','=ARRAYFORMULA(IF((E2:E="")+(G2:G=""),"",E2:E*((F2:F/G2:G)-1)))');
  putArray_(s,'J2','=ARRAYFORMULA(IF(E2:E="","",IF(I2:I="","",I2:I/E2:E)))');
  putArray_(s,'M2','=ARRAYFORMULA(IF((A2:A="")+(E2:E=""),"",N(K2:K)+N(L2:L)))');
  putArray_(s,'N2','=ARRAYFORMULA(IF(E2:E="","",M2:M/E2:E))');
  putArray_(s,'O2','=ARRAYFORMULA(IF(A2:A="","",DATE(YEAR(A2:A),MONTH(A2:A),1)))');
}

function setMonthlyFormulas_() {
  const s=mbSheet_('Monthly Metrics'); mbEnsureSize_(s,500,20); s.getRange(1,1,1,20).setValues([MB.MONTH_HEADERS]);
  putArray_(s,'A2','=IFNA(SORT(UNIQUE(FILTER({IF(Promos!A2:A="","",DATE(YEAR(Promos!A2:A),MONTH(Promos!A2:A),1));IF(BTO!A2:A="","",DATE(YEAR(BTO!A2:A),MONTH(BTO!A2:A),1));NonPromos!O2:O},{Promos!A2:A;BTO!A2:A;NonPromos!O2:O}<>""))),"")');
  const fs=[
    '=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(Promos!E:E,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1)))))','=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(Promos!H:H,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1)))))','=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",C2:C/B2:B)))','=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(Promos!I:I,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1))+SUMIFS(Promos!J:J,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1)))))','=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",E2:E/B2:B)))',
    '=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(BTO!E:E,BTO!A:A,">="&m,BTO!A:A,"<"&EDATE(m,1)))))','=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(BTO!I:I,BTO!A:A,">="&m,BTO!A:A,"<"&EDATE(m,1)))))','=ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",H2:H/G2:G)))','=MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(BTO!P:P,BTO!A:A,">="&m,BTO!A:A,"<"&EDATE(m,1)))))','=ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",J2:J/G2:G)))',
    '=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!O:O,A2:A,NonPromos!E:E)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!O:O,A2:A,NonPromos!I:I)))','=ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",M2:M/L2:L)))','=ARRAYFORMULA(IF(A2:A="","",SUMIF(NonPromos!O:O,A2:A,NonPromos!M:M)))','=ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",O2:O/L2:L)))','=ARRAYFORMULA(IF(A2:A="","",C2:C+H2:H+M2:M))','=ARRAYFORMULA(IF(A2:A="","",E2:E+J2:J+O2:O))','=ARRAYFORMULA(IF(A2:A="","",SCAN(0,Q2:Q,LAMBDA(a,v,a+v))))','=ARRAYFORMULA(IF(A2:A="","",SCAN(0,R2:R,LAMBDA(a,v,a+v))))'];
  fs.forEach((f,i)=>putArray_(s,String.fromCharCode(66+i)+'2',f));
}
