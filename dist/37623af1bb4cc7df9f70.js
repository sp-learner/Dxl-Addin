const Synonyms={ROUND:["BRILLIANT","RBC","BR","Round Brilliant","ROUND","RB","Round"],HEART:["HE","HB","Ht","S.HEART","HT","ROSE HEART","HRT JODI","Heart-P8-P8","Heart"],EMERALD:["EM","Em","E","EMERALD 4STEP","EMERALD","ASYMEMERALD4S","EM-HD","EMRALD","SQUARE EMERALD","EMR","Emerald Cut","EMERALD"],PEAR:["PE","SGP-100","PB","Pear Brilliant","PEAR MODIFIED BRILLIANT","PEAR"],OVAL:["OB","OVAL","OV","S.OVAL","OMB","Oval Brilliant","OVL","OVAL MODIFIED BRILLIANT","OVAL"],TRIANGLE:["TR","TS","Triangular","TRI","TRIANGLE"],VVS1:["VVS1","VVS1-","vvs1","VVS 1"],VVS2:["VVS2","VVS2-","vvs2","VVS 2"],IF:["IF","LOUPE-CLEAN","LC","Internally Flawless"],SI1:["SI1","SI1-","SI 1"],SI2:["SI2","SI2-","SI 2"],EX:["EXCELLENT","Excellent","EXC","IDEAL","EX","EX-2","EX-1","EX-3","X1","X2","X3","X4","ex1","ex2","I","Ex Ideal","EX3"],VG:["VERY GOOD","V. GOOD","VG","VV","VX","V"],GD:["GOOD","Good","GD","G","GV","GX","ex6","G"],FR:["FAIR","F","FR","FA","F"],PR:["P","POOR","PR","PU"],NONE:["N","NO","NIL","FL0","NN","Non","None"],FNT:["FAINT","FNT","NEGLIGIBLE","FL1","FA"],MED:["M","MEDIUM","MED","FL2","MD","MB","MEDIUM BLUE","ME","MD-BL","MD-YL"],STG:["STRONG","STG","ST","FL3","S","STRONG BLUE","STR","STO","ST-BL","ST-WH","ST-YL"],VST:["VERY STRONG","VST","FL4","VSTG","VST-BL"],SLT:["SL","SLIGHT","SLI","SLT"],VSL:["VERY SLIGHT","VSLG","VSLT","VSL","VS"],GIA:["G.I.A","GIA","GA","GIA"],IGI:["IGI"],NCERT:["NONE","NON","NC","NONCERT","NON CERT"]};function getUniformWord(E,S){for(const[L,R]of Object.entries(Synonyms))if(S===L)return R.includes(E)?L:E;return E}async function UniformData(){Excel.run((async E=>{let S=E.workbook.getSelectedRange();S.load("values"),await E.sync(),S.values=S.values.map((E=>E.map((E=>{for(const S in Synonyms){let L=getUniformWord(E,S);if(L!==E)return L}return E})))),await E.sync()}))}