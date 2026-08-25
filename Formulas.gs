function refreshFormulas() {
  setPromoFormulas_(); setBtoFormulas_(); setNonPromoFormulas_(); setMonthlyFormulas_(); refreshAccounts_();
}

function putArray_(sheet, header, formula) {
  const col=mbColumn_(sheet,header); sheet.getRange(2,col,Math.max(1,sheet.getMaxRows()-1),1).clearContent();
  sheet.getRange(2,col).setFormula(formula);
}

function setPromoFormulas_() {
  const s=mbSheet_('Promos');
  putArray_(s,'Estimated EV','=ARRAYFORMULA(IF(E2:E="","",E2:E*IFNA(VLOOKUP(C2:C,Settings!A2:B,2,FALSE),NA())))');
  putArray_(s,'Cash Change','=ARRAYFORMULA(IF(G2:G="","",IF(G2:G="W",E2:E*(F2:F-1),IF((G2:G="L")+(G2:G="B"),-E2:E,IF(G2:G="Void",0,"")))))');
  putArray_(s,'Bonus Change','=ARRAYFORMULA(IF(G2:G="","",IF(G2:G="B",E2:E,IF((G2:G="W")+(G2:G="L")+(G2:G="Void"),0,""))))');
}

function setBtoFormulas_() {
  const s=mbSheet_('BTO');
  putArray_(s,'EV Taken','=ARRAYFORMULA(IF((E2:E="")+(F2:F=""),"",IF(G2:G="",E2:E*Settings!E10,E2:E*(F2:F-1)/G2:G)))');
  putArray_(s,'Bookie Change','=ARRAYFORMULA(IF(H2:H="","",IF(H2:H="W",E2:E*(F2:F-1),IF((H2:H="L")+(H2:H="B"),-E2:E,IF(H2:H="Void",0,"")))))');
  putArray_(s,'Betfair Change','=ARRAYFORMULA(IF(H2:H="","",IF(H2:H="Void",0,IF((K2:K="")+(L2:L=""),"",IF(H2:H="W",-K2:K*(L2:L-1),IF((H2:H="L")+(H2:H="B"),K2:K*(1-N(M2:M)),""))))))');
}

function setNonPromoFormulas_() {
  const s=mbSheet_('NonPromos');
  putArray_(s,'Expected EV','=ARRAYFORMULA(IF((E2:E="")+(G2:G=""),"",E2:E*((F2:F/G2:G)-1)))');
  putArray_(s,'EV %','=ARRAYFORMULA(IF((E2:E="")+(I2:I=""),"",I2:I/E2:E))');
  putArray_(s,'Bookie Change','=ARRAYFORMULA(IF(H2:H="","",IF(H2:H="W",E2:E*(F2:F-1),IF((H2:H="L")+(H2:H="B"),-E2:E,IF(H2:H="Void",0,"")))))');
  putArray_(s,'Betfair Change','=ARRAYFORMULA(IF(H2:H="","",IF(H2:H="Void",0,IF((L2:L="")+(M2:M=""),"",IF(H2:H="W",-L2:L*(M2:M-1),IF((H2:H="L")+(H2:H="B"),L2:L*(1-N(N2:N)),""))))))');
}

function setMonthlyFormulas_() {
  const s=mbSheet_('Monthly Metrics'); mbEnsureSize_(s,500,20); s.getRange(1,1,1,20).setValues([MB.MONTH_HEADERS]);
  s.getRange(2,1,s.getMaxRows()-1,20).clearContent();
  s.getRange('A2').setFormula('=IFNA(SORT(UNIQUE(FILTER({IF(Promos!A2:A="","",EOMONTH(Promos!A2:A,-1)+1);IF(BTO!A2:A="","",EOMONTH(BTO!A2:A,-1)+1);IF(NonPromos!A2:A="","",EOMONTH(NonPromos!A2:A,-1)+1)},{Promos!A2:A;BTO!A2:A;NonPromos!A2:A}<>""))),"")');
  const sum=(sheet,col)=>`MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(${sheet}!${col}:${col},${sheet}!A:A,">="&m,${sheet}!A:A,"<"&EDATE(m,1)))))`;
  const formulas=[sum('Promos','E'),sum('Promos','H'),'ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",C2:C/B2:B)))',`MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(Promos!I:I,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1))+SUMIFS(Promos!J:J,Promos!A:A,">="&m,Promos!A:A,"<"&EDATE(m,1)))))`,'ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",E2:E/B2:B)))',sum('BTO','E'),sum('BTO','I'),'ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",H2:H/G2:G)))',`MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(BTO!J:J,BTO!A:A,">="&m,BTO!A:A,"<"&EDATE(m,1))+SUMIFS(BTO!N:N,BTO!A:A,">="&m,BTO!A:A,"<"&EDATE(m,1)))))`,'ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,"",J2:J/G2:G)))',sum('NonPromos','E'),sum('NonPromos','I'),'ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",M2:M/L2:L)))',`MAP(A2:A,LAMBDA(m,IF(m="","",SUMIFS(NonPromos!K:K,NonPromos!A:A,">="&m,NonPromos!A:A,"<"&EDATE(m,1))+SUMIFS(NonPromos!O:O,NonPromos!A:A,">="&m,NonPromos!A:A,"<"&EDATE(m,1)))))`,'ARRAYFORMULA(IF(A2:A="","",IF(L2:L=0,"",O2:O/L2:L)))','ARRAYFORMULA(IF(A2:A="","",C2:C+H2:H+M2:M))','ARRAYFORMULA(IF(A2:A="","",E2:E+J2:J+O2:O))','ARRAYFORMULA(IF(A2:A="","",SCAN(0,Q2:Q,LAMBDA(a,v,a+N(v)))))','ARRAYFORMULA(IF(A2:A="","",SCAN(0,R2:R,LAMBDA(a,v,a+N(v)))))'];
  formulas.forEach((f,i)=>s.getRange(2,i+2).setFormula('='+f));
}
