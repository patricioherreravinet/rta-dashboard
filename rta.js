
var SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT7gJBxlbAi-NeigWcsGex11tIKY3WmwncZz_PsXQkGStXkTgkagV33VS4CkHwZqR2guWxNvfSPgNUE/pub?gid=0&single=true&output=csv";
var LAST_SYNC = null;
var MES=["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
var DAILY=[],MONTHLY=[],ANNUAL=[],CAT=[],MO=[];
var sY=[],sM=[],sW=[],gran="dia",kpi="hh";
var myCh=null,pCh=null,pMet="un",topMet="un",pCd=null,pAi=-1,pSYrs=[];
var s1c="dt",s1a=true,s2c="k",s2a=true,s3c="yr",s3a=true;
var KCONF={
  hh:{c:"#4f8ef7",u:"HH",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  m2:{c:"#27c4a0",u:"m2",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  ml:{c:"#06b6d4",u:"ml",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  pp:{c:"#f97316",u:"pzs",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  cep:{c:"#ec4899",u:"ud",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  sku:{c:"#9b7dea",u:"SKUs",f:function(v){return Math.round(v).toLocaleString("es-CL");}},
  lote:{c:"#f0a500",u:"u/OF",f:function(v){return parseFloat(v).toFixed(1);}}
};
var AVG_KEYS=["lt","hpd","mpd","mlpd","pppd","sk"];

function pF(v){if(v===undefined||v===null||v==="")return 0;var s=String(v).trim().replace(/\./g,"").replace(",",".");return parseFloat(s)||0;}
function fN(v){return Math.round(v).toLocaleString("es-CL");}
function fD(v){return parseFloat(v).toFixed(1);}
function hv(a,v){var i;for(i=0;i<a.length;i++){if(a[i]===v)return true;}return false;}
function rv(a,v){var r=[],i;for(i=0;i<a.length;i++){if(a[i]!==v)r.push(a[i]);}return r;}
function uniq(a){var s={},r=[],i;for(i=0;i<a.length;i++){if(!s[a[i]]){s[a[i]]=1;r.push(a[i]);}}return r;}
function qsa(sel){return document.querySelectorAll(sel);}
function gel(id){return document.getElementById(id);}

function parseCSV(text){
  var lines=text.split("\n"),result=[],i,line,vals,inQ,val,j,ch;
  for(i=0;i<lines.length;i++){
    line=lines[i].replace(/\r$/,"");
    if(!line.trim())continue;
    vals=[];val="";inQ=false;
    for(j=0;j<line.length;j++){
      ch=line[j];
      if(ch==='"'){inQ=!inQ;}
      else if(ch===","&&!inQ){vals.push(val.trim());val="";}
      else{val+=ch;}
    }
    vals.push(val.trim());
    result.push(vals);
  }
  return result;
}

function setLoading(txt){
  var el=gel("loading-txt");
  if(el)el.textContent=txt;
}

function setSyncStatus(state){
  var dot=gel("sdot"),stxt=gel("stxt");
  if(!dot)return;
  dot.className="sync-dot"+(state==="loading"?" loading":state==="error"?" error":"");
  if(stxt)stxt.textContent=state==="loading"?"Sincronizando...":state==="error"?"Error - Reintentar":"Actualizar";
}

function processRawData(rows){
  var headers=rows[0],i,r;
  var COL={};
  for(i=0;i<headers.length;i++){
    var hk=headers[i].trim();
    hk=hk.replace(/A\u00c3\u00b1o/g,"Ano").replace(/A\u00c3\u00b1/g,"An").replace(/\u00c3\u00b1/g,"n").replace(/\u00c3\u0093/g,"O").replace(/\u00c3\u00a9/g,"e").replace(/\u00c3\u00a1/g,"a").replace(/\u00c3\u00b3/g,"o");
    COL[hk]=i;
    COL[headers[i].trim()]=i;
  }

  var hhI=COL["hh Total"],m2I=COL["M2 Total"],mlI=COL["ml Total"],
      ppI=COL["Total Piezas Perforadas"],cepI=COL["Cantidad CEP"],
      unI=COL["Unidades Total"],cdI=COL["CODIGO"],ofI=COL["OF Asignada"],
      dtI=COL["FECHA ENTREGA"],
      yrI=COL["A\u00f1o"]||COL["Ano"]||COL["Anio"]||COL["ANO"],
      moI=COL["Mes"]||COL["MES"],
      wkI=COL["semana"]||COL["Semana"],
      descI=COL["DESCRIPCION"]||COL["DESCRIPCION"]||COL["DESCRIPCION"],
      famI=COL["FAMILIA"],catI=COL["Categoria Nivel 0"];

  console.log("yrI="+yrI+" moI="+moI+" dtI="+dtI+" hhI="+hhI+" wkI="+wkI);
  var r4=rows[4]||rows[3]||[];
  console.log("DataRow cols:"+r4.length+" yr="+r4[yrI]+" mo="+r4[moI]+" hh="+r4[hhI]+" un="+r4[unI]+" m2="+r4[m2I]+" cd="+r4[cdI]);
  console.log("Headers found:", JSON.stringify(Object.keys(COL).slice(0,15)));
  console.log("Row4 sample:", JSON.stringify(rows[4]?rows[4].slice(0,8):[]));
  if(dtI===undefined&&yrI===undefined){
    console.log("Headers:",Object.keys(COL));return false;
  }

  var dayMap={},prodMap={};

  var cnt={total:0,short:0,dots:0,empty:0,noyr:0,nodt:0,bighh:0,nocd:0};
  for(i=1;i<rows.length;i++){
    r=rows[i];cnt.total++;
    if(!r||r.length<3){cnt.short++;continue;}
    if(!String(r[yrI]||r[dtI]||"").trim()){cnt.empty++;continue;}
    var yr,mo2,wk,dt,dtRaw,pts,yy,mm,dd;

    if(yrI!==undefined&&r[yrI]&&String(r[yrI]).trim()){
      yr=parseInt(r[yrI],10);
      mo2=moI!==undefined?parseInt(r[moI],10):0;
      wk=wkI!==undefined?parseInt(r[wkI],10)||1:1;
      if(isNaN(yr)||yr<2000||isNaN(mo2)||mo2<1||mo2>12)continue;
      dd="01";
      if(dtI!==undefined&&r[dtI]){
        dtRaw=String(r[dtI]).trim();
        if(dtRaw.indexOf("/")>=0){pts=dtRaw.split("/");dd=pts[0].length===1?"0"+pts[0]:pts[0].slice(0,2);}
        else if(dtRaw.indexOf("-")>=0){dd=dtRaw.slice(8,10)||"01";}
      }
      dt=yr+"-"+(mo2<10?"0":"")+mo2+"-"+dd;
    } else if(dtI!==undefined&&r[dtI]&&String(r[dtI]).trim()){
      dtRaw=String(r[dtI]).trim();dt="";
      if(dtRaw.indexOf("/")>=0){
        pts=dtRaw.split("/");
        if(pts.length===3){
          if(pts[2].length===4){
            yy=pts[2];mm=pts[1].length===1?"0"+pts[1]:pts[1];dd=pts[0].length===1?"0"+pts[0]:pts[0];
          } else {
            yy=pts[0].length===4?pts[0]:"20"+pts[0];mm=pts[1].length===1?"0"+pts[1]:pts[1];dd=pts[2].length===1?"0"+pts[2]:pts[2];
          }
          dt=yy+"-"+mm+"-"+dd;
        }
      } else if(dtRaw.indexOf("-")>=0){
        pts=dtRaw.split("-");
        if(pts.length===3){
          if(pts[0].length===4){dt=dtRaw.slice(0,10);}
          else{
            yy=pts[2].length===4?pts[2]:"20"+pts[2];
            mm=pts[1].length===1?"0"+pts[1]:pts[1];
            dd=pts[0].length===1?"0"+pts[0]:pts[0];
            dt=yy+"-"+mm+"-"+dd;
          }
        }
      } else if(!isNaN(parseFloat(dtRaw))){
        var serial=parseInt(dtRaw,10);
        var d2=new Date(new Date(1899,11,30).getTime()+serial*86400000);
        yr=d2.getFullYear();mo2=d2.getMonth()+1;dd=d2.getDate();
        dt=yr+"-"+(mo2<10?"0":"")+mo2+"-"+(dd<10?"0":"")+dd;
      }
      if(!dt||dt.length<10)continue;
      yr=parseInt(dt.slice(0,4),10);mo2=parseInt(dt.slice(5,7),10);
      var wd=new Date(dt),j1=new Date(yr,0,1);
      wk=Math.ceil((((wd-j1)/86400000)+j1.getDay()+1)/7);
    } else { continue; }

    if(!yr||!mo2||isNaN(yr)||isNaN(mo2)){cnt.noyr++;continue;}

    var hh=pF(r[hhI]);
    if(hh>5000){cnt.bighh++;continue;}
    var m2v=pF(r[m2I]);
    var ml=pF(r[mlI]);
    var pp=pF(r[ppI]);
    var cep=pF(r[cepI]);
    var un=pF(r[unI]);
    var cd=String(r[cdI]!==undefined?r[cdI]:"").trim();
    var of=String(ofI!==undefined&&r[ofI]!==undefined?r[ofI]:"").trim()||cd;
    if(!cd&&un===0&&hh===0&&m2v===0){cnt.nocd++;continue;}
    var desc=String(descI!==undefined&&r[descI]?r[descI]:"").trim().replace(/&#160;/g," ")||cd;
    var fam=String(famI!==undefined&&r[famI]?r[famI]:"").trim();
    var cat=String(catI!==undefined&&r[catI]?r[catI]:"").trim();

    if(Object.keys(dayMap).length===0)console.log("First valid row: dt="+dt+" yr="+yr+" mo="+mo2+" hh="+hh+" un="+un+" cd="+cd);
    if(!dayMap[dt])dayMap[dt]={dt:dt,yr:yr,mo:mo2,wk:wk,hh:0,m2:0,ml:0,pp:0,cep:0,un:0,ofs:new Set(),cds:new Set()};
    var dm=dayMap[dt];
    dm.hh+=hh;dm.m2+=m2v;dm.ml+=ml;dm.pp+=pp;dm.cep+=cep;dm.un+=un;
    if(of)dm.ofs.add(of);if(cd)dm.cds.add(cd);

    if(cd){
      if(!prodMap[cd])prodMap[cd]={cd:cd,desc:desc,fam:fam,cat:cat,months:{}};
      var pm=prodMap[cd];
      if(desc&&!pm.desc)pm.desc=desc;
      var mk=yr+"-"+mo2;
      if(!pm.months[mk])pm.months[mk]={yr:yr,mo:mo2,un:0,m2:0,hh:0};
      pm.months[mk].un+=un;pm.months[mk].m2+=m2v;pm.months[mk].hh+=hh;
    }
  }

  buildAggregates();

  CAT=[];MO=[];
  var pks=Object.keys(prodMap),j2;
  for(i=0;i<pks.length;i++){
    var p2=prodMap[pks[i]];
    var mks2=Object.keys(p2.months);
    var totU=0,totM2v=0,totH=0,firstDt="9999",lastDt="0000";
    for(j2=0;j2<mks2.length;j2++){
      var mmx=p2.months[mks2[j2]];
      totU+=mmx.un;totM2v+=mmx.m2;totH+=mmx.hh;
      var ms2=mmx.yr+"-"+(mmx.mo<10?"0":"")+mmx.mo;
      if(ms2<firstDt)firstDt=ms2;if(ms2>lastDt)lastDt=ms2;
    }
    CAT.push([p2.cd,p2.desc||p2.cd,p2.fam,p2.cat,Math.round(totU),Math.round(totM2v*10)/10,Math.round(totH*10)/10,firstDt,lastDt,mks2.length]);
    for(j2=0;j2<mks2.length;j2++){
      var mmx2=p2.months[mks2[j2]];
      MO.push([p2.cd,mmx2.yr,mmx2.mo,Math.round(mmx2.un),Math.round(mmx2.m2*10)/10,Math.round(mmx2.hh*10)/10]);
    }
  }
  console.log("Counters:",JSON.stringify(cnt));gel("pshin").textContent=CAT.length+" productos disponibles";console.log("DAILY:"+DAILY.length+" MONTHLY:"+MONTHLY.length+" CAT:"+CAT.length);if(DAILY.length>0)console.log("First daily:"+JSON.stringify(DAILY[0]));
  return true;
}

function buildAggregates(){
  var i,r,k,mo2,yr,d,m,a;
  var moMap={},yrMap={};

  for(i=0;i<DAILY.length;i++){
    r=DAILY[i];
    yr=r[1];mo2=r[2];
    k=yr+"-"+mo2;
    if(!moMap[k])moMap[k]={yr:yr,mo:mo2,hh:0,m2:0,ml:0,pp:0,cep:0,un:0,sk_s:0,of:0,n:0,ds:0};
    m=moMap[k];
    m.hh+=r[4];m.m2+=r[5];m.ml+=r[6];m.pp+=r[7];m.cep+=r[8];
    m.un+=r[9];m.sk_s+=r[10];m.of+=r[11];m.n++;m.ds++;

    if(!yrMap[yr])yrMap[yr]={yr:yr,hh:0,m2:0,ml:0,pp:0,cep:0,un:0,sk_s:0,of:0,n:0,ds:0};
    a=yrMap[yr];
    a.hh+=r[4];a.m2+=r[5];a.ml+=r[6];a.pp+=r[7];a.cep+=r[8];
    a.un+=r[9];a.sk_s+=r[10];a.of+=r[11];a.n++;a.ds++;
  }

  MONTHLY=[];
  var mks=Object.keys(moMap).sort();
  for(i=0;i<mks.length;i++){
    m=moMap[mks[i]];d=m.ds;
    var lt=m.of>0?+(m.un/m.of).toFixed(1):0;
    MONTHLY.push({yr:m.yr,mo:m.mo,hh:Math.round(m.hh*10)/10,m2:Math.round(m.m2*10)/10,
      ml:Math.round(m.ml*10)/10,pp:Math.round(m.pp),cep:Math.round(m.cep),un:Math.round(m.un),
      sk:Math.round(m.sk_s/m.n),of:Math.round(m.of),lt:lt,ds:d,
      hpd:Math.round(m.hh/d*10)/10,mpd:Math.round(m.m2/d*10)/10,
      mlpd:Math.round(m.ml/d*10)/10,pppd:Math.round(m.pp/d)});
  }

  ANNUAL=[];
  var yks=Object.keys(yrMap).sort();
  for(i=0;i<yks.length;i++){
    a=yrMap[yks[i]];d=a.ds;
    var lt2=a.of>0?+(a.un/a.of).toFixed(1):0;
    ANNUAL.push({yr:a.yr,hh:Math.round(a.hh*10)/10,m2:Math.round(a.m2*10)/10,
      ml:Math.round(a.ml*10)/10,pp:Math.round(a.pp),cep:Math.round(a.cep),un:Math.round(a.un),
      sk:Math.round(a.sk_s/a.n),of:Math.round(a.of),lt:lt2,ds:d,
      hpd:Math.round(a.hh/d*10)/10,mpd:Math.round(a.m2/d*10)/10,
      mlpd:Math.round(a.ml/d*10)/10,pppd:Math.round(a.pp/d)});
  }
}

function syncData(){
  setSyncStatus("loading");
  setLoading("Sincronizando con Google Sheets...");
  var ts="&ts="+Date.now();
  fetch(SHEET_URL+ts)
    .then(function(r){
      if(!r.ok)throw new Error("HTTP "+r.status);
      return r.text();
    })
    .then(function(text){
      setLoading("Procesando datos...");
      var rows=parseCSV(text);
      if(!processRawData(rows)){
        throw new Error("Error en columnas");
      }
      LAST_SYNC=new Date();
      setSyncStatus("ok");
      gel("loading").style.display="none";
      initChips();
      upAll();
    })
    .catch(function(err){
      console.error("Error:",err);
      setSyncStatus("error");
      gel("loading-txt").textContent="Error al conectar. Verifica que la hoja este publicada.";
      setTimeout(function(){gel("loading").style.display="none";},3000);
    });
}

function getDR(){var f=gel("dr-from").value,t=gel("dr-to").value;return{f:f||null,t:t||null};}
function clrDr(){gel("dr-from").value="";gel("dr-to").value="";}

function initChips(){
  var yrs=uniq(DAILY.map(function(r){return r[1];})).sort(function(a,b){return a-b;});
  var mns=uniq(DAILY.map(function(r){return r[2];})).sort(function(a,b){return a-b;});
  var wks=uniq(DAILY.map(function(r){return r[3];})).sort(function(a,b){return a-b;});
  var cy=gel("cy"),cm=gel("cm"),cw=gel("cw"),el,i;
  cy.innerHTML="";cm.innerHTML="";cw.innerHTML="";
  for(i=0;i<yrs.length;i++){
    el=document.createElement("div");el.className="chip";
    el.setAttribute("data-t","y");el.setAttribute("data-v",yrs[i]);
    el.textContent=yrs[i];
    (function(y){el.onclick=function(){togChip(this,"y",y);};})(yrs[i]);
    cy.appendChild(el);
  }
  for(i=0;i<mns.length;i++){
    el=document.createElement("div");el.className="chip";
    el.setAttribute("data-t","m");el.setAttribute("data-v",mns[i]);
    el.textContent=MES[mns[i]];
    (function(mo){el.onclick=function(){togChip(this,"m",mo);};})(mns[i]);
    cm.appendChild(el);
  }
  for(i=0;i<wks.length;i++){
    el=document.createElement("div");el.className="chip";
    el.setAttribute("data-t","w");el.setAttribute("data-v",wks[i]);
    el.textContent="S"+wks[i];
    (function(wk){el.onclick=function(){togChip(this,"w",wk);};})(wks[i]);
    cw.appendChild(el);
  }
}

function togChip(el,t,v){
  if(t==="y"){if(hv(sY,v)){sY=rv(sY,v);el.classList.remove("on");}else{sY.push(v);el.classList.add("on");}}
  else if(t==="m"){if(hv(sM,v)){sM=rv(sM,v);el.classList.remove("on");}else{sM.push(v);el.classList.add("on");}}
  else if(t==="w"){if(hv(sW,v)){sW=rv(sW,v);el.classList.remove("on");}else{sW.push(v);el.classList.add("on");}}
  upAll();
}

function clrF(t){
  var els,i;
  if(t==="y"){sY=[];els=qsa(".chip[data-t='y']");for(i=0;i<els.length;i++)els[i].classList.remove("on");}
  else if(t==="m"){sM=[];els=qsa(".chip[data-t='m']");for(i=0;i<els.length;i++)els[i].classList.remove("on");}
  else if(t==="w"){sW=[];els=qsa(".chip[data-t='w']");for(i=0;i<els.length;i++)els[i].classList.remove("on");}
  upAll();
}

function resetAll(){
  var els=qsa(".chip"),i;
  sY=[];sM=[];sW=[];
  for(i=0;i<els.length;i++)els[i].classList.remove("on");
  clrDr();upAll();
}

function setG(btn,g){
  var btns=qsa(".gc"),i;
  for(i=0;i<btns.length;i++)btns[i].classList.remove("on");
  btn.classList.add("on");gran=g;upAll();
}

function selK(k){
  var btns,i;
  btns=qsa(".kc");for(i=0;i<btns.length;i++)btns[i].classList.remove("on");
  btns=qsa(".mfb");for(i=0;i<btns.length;i++)btns[i].classList.remove("on");
  var kc=gel("kc-"+k),mf=gel("mf-"+k);
  if(kc)kc.classList.add("on");if(mf)mf.classList.add("on");
  kpi=k;upChart();
}

function fDay(){
  var dr=getDR(),out=[],r,i;
  for(i=0;i<DAILY.length;i++){
    r=DAILY[i];
    if(sY.length&&!hv(sY,r[1]))continue;
    if(sM.length&&!hv(sM,r[2]))continue;
    if(sW.length&&!hv(sW,r[3]))continue;
    if(dr.f&&r[0]<dr.f)continue;
    if(dr.t&&r[0]>dr.t)continue;
    out.push(r);
  }
  return out;
}

function fMo(){
  var dr=getDR(),out=[],r,i,ym;
  for(i=0;i<MONTHLY.length;i++){
    r=MONTHLY[i];
    if(sY.length&&!hv(sY,r.yr))continue;
    if(sM.length&&!hv(sM,r.mo))continue;
    if(dr.f){ym=r.yr+"-"+(r.mo<10?"0":"")+r.mo;if(ym<dr.f.slice(0,7))continue;}
    if(dr.t){ym=r.yr+"-"+(r.mo<10?"0":"")+r.mo;if(ym>dr.t.slice(0,7))continue;}
    out.push(r);
  }
  return out;
}

function fYr(){
  var dr=getDR(),out=[],r,i;
  for(i=0;i<ANNUAL.length;i++){
    r=ANNUAL[i];
    if(sY.length&&!hv(sY,r.yr))continue;
    if(dr.f&&r.yr<parseInt(dr.f.slice(0,4),10))continue;
    if(dr.t&&r.yr>parseInt(dr.t.slice(0,4),10))continue;
    out.push(r);
  }
  return out;
}

function agg(rows){
  var i,r,k,l,map,ks,v,out,x;
  if(gran==="dia"){
    out=[];
    for(i=0;i<rows.length;i++){
      r=rows[i];
      out.push({p:r[0].slice(5).replace("-","/"),k:r[0],hh:r[4],m2:r[5],ml:r[6],pp:r[7],cep:r[8],un:r[9],sk:r[10],of:r[11],lt:r[12]});
    }
    return out;
  }
  map={};
  for(i=0;i<rows.length;i++){
    r=rows[i];
    k=gran==="mes"?r[1]+"-"+(r[2]<10?"0":"")+r[2]:String(r[1]);
    l=gran==="mes"?MES[r[2]]+" "+r[1]:String(r[1]);
    if(!map[k])map[k]={p:l,k:k,hh:0,m2:0,ml:0,pp:0,cep:0,un:0,of:0,ss:0,ln:0};
    x=map[k];
    x.hh+=r[4];x.m2+=r[5];x.ml+=r[6];x.pp+=r[7];x.cep+=r[8];
    x.un+=r[9];x.ss+=r[10];x.of+=r[11];x.ln++;
  }
  ks=Object.keys(map).sort();out=[];
  for(i=0;i<ks.length;i++){
    v=map[ks[i]];
    out.push({p:v.p,k:ks[i],hh:Math.round(v.hh),m2:Math.round(v.m2),ml:Math.round(v.ml),
      pp:Math.round(v.pp),cep:Math.round(v.cep),un:Math.round(v.un),of:Math.round(v.of),
      sk:Math.round(v.ss/v.ln),lt:v.of>0?+(Math.round(v.un)/Math.round(v.of)).toFixed(1):0});
  }
  return out;
}

function updBdg(){
  var p=[],i;
  if(sY.length)p.push(sY.slice().sort().join(", "));
  if(sM.length)p.push(sM.slice().sort(function(a,b){return a-b;}).map(function(m){return MES[m];}).join(", "));
  if(sW.length)p.push("S"+sW.slice().sort(function(a,b){return a-b;}).join("+S"));
  gel("bi").textContent=p.length?p.join(" - "):"Todo el periodo";
}

function upCards(){
  var rows=agg(fDay()),n=rows.length||1,tH=0,tM=0,tML=0,tPP=0,i;
  for(i=0;i<rows.length;i++){tH+=rows[i].hh;tM+=rows[i].m2;tML+=rows[i].ml;tPP+=rows[i].pp;}
  var sfx=gran==="dia"?"/dia":gran==="mes"?"/mes":"";
  gel("kv-hh").textContent=fN(tH);gel("kv-m2").textContent=fN(tM);
  gel("kv-ml").textContent=fN(tML);gel("kv-pp").textContent=fN(tPP);
  gel("ka-hh").textContent=gran!=="yr"?"prom"+sfx+": "+fN(tH/n)+" HH":"";
  gel("ka-m2").textContent=gran!=="yr"?"prom"+sfx+": "+fN(tM/n)+" m2":"";
  gel("ka-ml").textContent=gran!=="yr"?"prom"+sfx+": "+fN(tML/n)+" ml":"";
  gel("ka-pp").textContent=gran!=="yr"?"prom"+sfx+": "+fN(tPP/n)+" pzs":"";
}

function upChart(){
  var rows=agg(fDay()),cfg=KCONF[kpi],kmap={hh:"hh",m2:"m2",ml:"ml",pp:"pp",cep:"cep",sku:"sk",lote:"lt"};
  var kk=kmap[kpi],labels=[],vals=[],avg=0,isL=gran==="dia",i;
  for(i=0;i<rows.length;i++){labels.push(rows[i].p);vals.push(rows[i][kk]||0);}
  for(i=0;i<vals.length;i++)avg+=vals[i];
  if(vals.length)avg/=vals.length;
  gel("ch-t").textContent=cfg.u+" - por "+(gran==="yr"?"anio":gran);
  gel("ch-m").textContent="prom: "+cfg.f(avg)+" "+cfg.u;
  if(myCh){myCh.destroy();myCh=null;}
  myCh=new Chart(gel("mc"),{
    type:isL?"line":"bar",
    data:{labels:labels,datasets:[{data:vals,borderColor:cfg.c,backgroundColor:isL?cfg.c+"20":cfg.c+"cc",borderWidth:isL?2:0,pointRadius:0,fill:isL,tension:0.35,borderRadius:isL?0:4}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return cfg.f(c.parsed.y)+" "+cfg.u;}}}},
      scales:{x:{ticks:{color:"#3e4459",font:{size:10},maxRotation:45,autoSkip:true,maxTicksLimit:16},grid:{color:"rgba(255,255,255,0.04)"}},
              y:{ticks:{color:"#3e4459",font:{size:11},callback:function(v){return cfg.f(v);}},grid:{color:"rgba(255,255,255,0.04)"}}}
    }
  });
}

function mhd(cols,sc,sa,fn){
  var h="<tr>",i,c,ar;
  for(i=0;i<cols.length;i++){c=cols[i];ar=sc===c.k?(sa?"&#8593;":"&#8595;"):"&#8597;";h+="<th onclick=\""+fn+"('"+c.k+"')\" class=\""+(sc===c.k?"s":"")+"\" style=\"text-align:"+(c.fl?"left":"right")+"\">"+c.l+" <span style=\"opacity:.4\">"+ar+"</span></th>";}
  return h+"</tr>";
}
function mcell(k,v,mx,av,col){
  var w,tag,fmtd;
  if(v==null)return"<td>-</td>";
  w=Math.round((v/Math.max(mx[k]||1,1))*38);
  tag=v>=(av[k]||0)*1.15?"<span class=\"tag hi\">+</span>":v<=(av[k]||0)*0.85?"<span class=\"tag lo\">-</span>":"";
  fmtd=KCONF[k]?KCONF[k].f(v):fN(v);
  return"<td><div class=\"bc\">"+tag+"<div class=\"mbar\" style=\"width:"+w+"px;background:"+col+"55\"></div>"+fmtd+"</div></td>";
}
function mma(rows,cols){
  var mx={},av={},c,vs,sum,i,ci,v;
  for(ci=0;ci<cols.length;ci++){c=cols[ci];if(c.fl)continue;vs=[];sum=0;for(i=0;i<rows.length;i++){v=rows[i][c.k];if(v!=null&&!isNaN(v)){vs.push(v);sum+=v;}}mx[c.k]=vs.length?Math.max.apply(null,vs):1;av[c.k]=vs.length?sum/vs.length:0;}
  return{mx:mx,av:av};
}
function mft(rows,cols){
  var h="<tr>",n=rows.length||1,ci,c,isA,sum,v,i,tun,tof,kf;
  for(ci=0;ci<cols.length;ci++){
    c=cols[ci];if(c.fl){h+="<td>TOTAL/PROM</td>";continue;}
    if(c.k==="lt"){tun=0;tof=0;for(i=0;i<rows.length;i++){tun+=(rows[i].un||0);tof+=(rows[i].of||0);}v=tof>0?+(tun/tof).toFixed(1):0;}
    else{isA=hv(AVG_KEYS,c.k);sum=0;for(i=0;i<rows.length;i++)sum+=(rows[i][c.k]||0);v=isA?sum/n:sum;}
    kf=c.k==="sk"?"sku":c.k;h+="<td>"+(KCONF[kf]?KCONF[kf].f(v):fN(v))+"</td>";
  }
  return h+"</tr>";
}

var C1=[{k:"p",l:"Periodo",fl:true},{k:"hh",l:"HH",col:"#4f8ef7"},{k:"m2",l:"M2",col:"#27c4a0"},{k:"ml",l:"ML",col:"#06b6d4"},{k:"pp",l:"Pzs Perf.",col:"#f97316"},{k:"cep",l:"CEP",col:"#ec4899"},{k:"sk",l:"SKUs",col:"#9b7dea"},{k:"lt",l:"Lote",col:"#f0a500"},{k:"un",l:"Unidades",col:"#5a6080"},{k:"of",l:"OFs",col:"#5a6080"}];
var C2=[{k:"_l",l:"Mes/Anio",fl:true},{k:"hh",l:"HH",col:"#4f8ef7"},{k:"m2",l:"M2",col:"#27c4a0"},{k:"ml",l:"ML",col:"#06b6d4"},{k:"pp",l:"Pzs Perf.",col:"#f97316"},{k:"cep",l:"CEP",col:"#ec4899"},{k:"hpd",l:"HH/dia",col:"#7bbdff"},{k:"mpd",l:"m2/dia",col:"#7de8d0"},{k:"mlpd",l:"ML/dia",col:"#a5f3fc"},{k:"pppd",l:"PP/dia",col:"#fed7aa"},{k:"sk",l:"SKUs",col:"#9b7dea"},{k:"lt",l:"Lote",col:"#f0a500"},{k:"un",l:"Unidades",col:"#5a6080"},{k:"of",l:"OFs",col:"#5a6080"},{k:"ds",l:"Dias",col:"#5a6080"}];
var C3=[{k:"yr",l:"Anio",fl:true},{k:"hh",l:"HH",col:"#4f8ef7"},{k:"m2",l:"M2",col:"#27c4a0"},{k:"ml",l:"ML",col:"#06b6d4"},{k:"pp",l:"Pzs Perf.",col:"#f97316"},{k:"cep",l:"CEP",col:"#ec4899"},{k:"hpd",l:"HH/dia",col:"#7bbdff"},{k:"mpd",l:"m2/dia",col:"#7de8d0"},{k:"mlpd",l:"ML/dia",col:"#a5f3fc"},{k:"pppd",l:"PP/dia",col:"#fed7aa"},{k:"sk",l:"SKUs",col:"#9b7dea"},{k:"lt",l:"Lote",col:"#f0a500"},{k:"un",l:"Unidades",col:"#5a6080"},{k:"of",l:"OFs",col:"#5a6080"},{k:"ds",l:"Dias",col:"#5a6080"}];

function rT1(){var s=(gel("t1i").value||"").toLowerCase(),rows=agg(fDay()),f=[],i,r,ma,tb,ci;if(s){for(i=0;i<rows.length;i++)if(rows[i].p.toLowerCase().indexOf(s)>=0)f.push(rows[i]);rows=f;}rows.sort(function(a,b){var va=s1c==="p"?a.k:a[s1c],vb=s1c==="p"?b.k:b[s1c];return s1a?(va>vb?1:-1):(va<vb?1:-1);});gel("t1s").textContent="Detalle por "+(gran==="yr"?"anio":gran);gel("t1t").textContent=rows.length+" registros";gel("t1h").innerHTML=mhd(C1,s1c,s1a,"sT1");if(!rows.length){gel("t1b").innerHTML="<tr><td colspan=\""+C1.length+"\" class=\"emp\">Sin datos</td></tr>";gel("t1f").innerHTML="";return;}ma=mma(rows,C1);tb="";for(i=0;i<rows.length;i++){r=rows[i];tb+="<tr><td>"+r.p+"</td>";for(ci=1;ci<C1.length;ci++)tb+=mcell(C1[ci].k,r[C1[ci].k],ma.mx,ma.av,C1[ci].col);tb+="</tr>";}gel("t1b").innerHTML=tb;gel("t1f").innerHTML=mft(rows,C1);}
function sT1(c){if(s1c===c)s1a=!s1a;else{s1c=c;s1a=true;}rT1();}
function rT2(){var s=(gel("t2i").value||"").toLowerCase(),raw=fMo(),rows=[],f=[],i,r,ma,tb,ci;for(i=0;i<raw.length;i++){r=raw[i];rows.push({_l:MES[r.mo]+" "+r.yr,k:r.yr+"-"+(r.mo<10?"0":"")+r.mo,hh:r.hh,m2:r.m2,ml:r.ml,pp:r.pp,cep:r.cep,hpd:r.hpd,mpd:r.mpd,mlpd:r.mlpd,pppd:r.pppd,sk:r.sk,lt:r.lt,un:r.un,of:r.of,ds:r.ds});}if(s){for(i=0;i<rows.length;i++)if(rows[i]._l.toLowerCase().indexOf(s)>=0)f.push(rows[i]);rows=f;}rows.sort(function(a,b){var va=s2c==="_l"?a.k:a[s2c],vb=s2c==="_l"?b.k:b[s2c];return s2a?(va>vb?1:-1):(va<vb?1:-1);});gel("t2t").textContent=rows.length+" meses";gel("t2h").innerHTML=mhd(C2,s2c,s2a,"sT2");if(!rows.length){gel("t2b").innerHTML="<tr><td colspan=\""+C2.length+"\" class=\"emp\">Sin datos</td></tr>";gel("t2f").innerHTML="";return;}ma=mma(rows,C2);tb="";for(i=0;i<rows.length;i++){r=rows[i];tb+="<tr><td>"+r._l+"</td>";for(ci=1;ci<C2.length;ci++)tb+=mcell(C2[ci].k,r[C2[ci].k],ma.mx,ma.av,C2[ci].col);tb+="</tr>";}gel("t2b").innerHTML=tb;gel("t2f").innerHTML=mft(rows,C2);}
function sT2(c){if(s2c===c)s2a=!s2a;else{s2c=c;s2a=true;}rT2();}
function rT3(){var rows=fYr(),i,r,ma,tb,ci;rows.sort(function(a,b){var va=a[s3c],vb=b[s3c];return s3a?(va>vb?1:-1):(va<vb?1:-1);});gel("t3t").textContent=rows.length+" anios";gel("t3h").innerHTML=mhd(C3,s3c,s3a,"sT3");if(!rows.length){gel("t3b").innerHTML="<tr><td colspan=\""+C3.length+"\" class=\"emp\">Sin datos</td></tr>";gel("t3f").innerHTML="";return;}ma=mma(rows,C3);tb="";for(i=0;i<rows.length;i++){r=rows[i];tb+="<tr><td>"+r.yr+"</td>";for(ci=1;ci<C3.length;ci++)tb+=mcell(C3[ci].k,r[C3[ci].k],ma.mx,ma.av,C3[ci].col);tb+="</tr>";}gel("t3b").innerHTML=tb;gel("t3f").innerHTML=mft(rows,C3);}
function sT3(c){if(s3c===c)s3a=!s3a;else{s3c=c;s3a=true;}rT3();}

function setTopM(m,btn){var btns=qsa("#top-un,#top-m2,#top-hh"),i;for(i=0;i<btns.length;i++)btns[i].classList.remove("on");btn.classList.add("on");topMet=m;rTop();}
function rTop(){
  var dr=getDR(),totals={},r,i,cd,rows,keys,top10,gT,catMap,maxV,tb,pct,bw,info,rc;
  for(i=0;i<MO.length;i++){
    r=MO[i];
    if(sY.length&&!hv(sY,r[1]))continue;
    if(sM.length&&!hv(sM,r[2]))continue;
    if(dr.f){var ds=r[1]+"-"+(r[2]<10?"0":"")+r[2];if(ds<dr.f.slice(0,7))continue;}
    if(dr.t){var de=r[1]+"-"+(r[2]<10?"0":"")+r[2];if(de>dr.t.slice(0,7))continue;}
    cd=r[0];if(!totals[cd])totals[cd]={cd:cd,un:0,m2:0,hh:0};
    totals[cd].un+=r[3];totals[cd].m2+=r[4];totals[cd].hh+=r[5];
  }
  rows=[];keys=Object.keys(totals);for(i=0;i<keys.length;i++)rows.push(totals[keys[i]]);
  rows.sort(function(a,b){return b[topMet]-a[topMet];});
  top10=rows.slice(0,10);gT=0;for(i=0;i<rows.length;i++)gT+=rows[i][topMet];
  catMap={};for(i=0;i<CAT.length;i++)catMap[CAT[i][0]]={desc:CAT[i][1],fam:CAT[i][2]};
  var lbl={un:"Unidades",m2:"M2",hh:"HH"};gel("top-tt").textContent="Top 10 por "+lbl[topMet];
  maxV=top10.length?top10[0][topMet]:1;tb="";
  for(i=0;i<top10.length;i++){
    r=top10[i];info=catMap[r.cd]||{desc:r.cd,fam:"-"};
    pct=gT>0?((r[topMet]/gT)*100).toFixed(1):0;bw=Math.round((r[topMet]/maxV)*80);
    rc=i===0?"#f0a500":i===1?"#7c849a":i===2?"#cd7f32":"#3e4459";
    tb+="<tr><td style=\"text-align:left;font-family:monospace;color:"+rc+";font-weight:700\">"+(i+1)+"</td>";
    tb+="<td style=\"text-align:left;max-width:200px\"><div style=\"font-size:10px;color:#4f8ef7;font-family:monospace\">"+r.cd+"</div><div style=\"white-space:nowrap;overflow:hidden;text-overflow:ellipsis\">"+info.desc+"</div></td>";
    tb+="<td style=\"text-align:left\"><span style=\"font-size:10px;padding:2px 6px;border-radius:10px;background:#202535;color:#7c849a\">"+info.fam+"</span></td>";
    tb+="<td><div class=\"bc\"><div class=\"mbar\" style=\"width:"+bw+"px;background:#4f8ef755\"></div>"+fN(r.un)+"</div></td>";
    tb+="<td>"+fD(r.m2)+"</td><td>"+fD(r.hh)+"</td><td style=\"color:#7c849a\">"+pct+"%</td></tr>";
  }
  if(!top10.length)tb="<tr><td colspan=\"7\" class=\"emp\">Sin datos</td></tr>";
  gel("topb").innerHTML=tb;
}

function upAll(){updBdg();upCards();upChart();rT1();rT2();rT3();rTop();}

function psIn(){var q=gel("psi").value.trim().toLowerCase(),hits=[],ac,h,i,p;pAi=-1;if(q.length<2){psCA();return;}for(i=0;i<CAT.length;i++){p=CAT[i];if(p[0].toLowerCase().indexOf(q)>=0||p[1].toLowerCase().indexOf(q)>=0){hits.push(p);if(hits.length>=10)break;}}if(!hits.length){psCA();return;}ac=gel("psac");h="";for(i=0;i<hits.length;i++){p=hits[i];h+="<div class=\"aci\" onclick=\"psPk('"+p[0]+"')\"><div class=\"acode\">"+p[0]+"</div><div class=\"adesc\">"+p[1]+"</div><div class=\"afam\">"+p[2]+"</div></div>";}ac.innerHTML=h;ac.style.display="block";}
function psKd(e){var ac=gel("psac"),its=ac.querySelectorAll(".aci"),i;if(e.key==="ArrowDown"){pAi=Math.min(pAi+1,its.length-1);for(i=0;i<its.length;i++)its[i].classList.remove("sel");if(pAi>=0&&its[pAi])its[pAi].classList.add("sel");e.preventDefault();}else if(e.key==="ArrowUp"){pAi=Math.max(pAi-1,0);for(i=0;i<its.length;i++)its[i].classList.remove("sel");if(pAi>=0&&its[pAi])its[pAi].classList.add("sel");e.preventDefault();}else if(e.key==="Enter"){if(pAi>=0&&its[pAi])its[pAi].click();else psSrch();}else if(e.key==="Escape")psCA();}
function psCA(){gel("psac").style.display="none";pAi=-1;}
function psPk(cd){pCd=cd;var p=null,i;for(i=0;i<CAT.length;i++)if(CAT[i][0]===cd){p=CAT[i];break;}if(p)gel("psi").value=p[0]+" - "+p[1];psCA();psRen(cd);}
function psSrch(){var q=gel("psi").value.trim(),i;if(!q)return;for(i=0;i<CAT.length;i++){if(CAT[i][0].toLowerCase()===q.toLowerCase()||CAT[i][1].toLowerCase()===q.toLowerCase()){psRen(CAT[i][0]);return;}}for(i=0;i<CAT.length;i++){if(CAT[i][0].toLowerCase().indexOf(q.toLowerCase())>=0){psRen(CAT[i][0]);return;}}for(i=0;i<CAT.length;i++){if(CAT[i][1].toLowerCase().indexOf(q.toLowerCase())>=0){psRen(CAT[i][0]);return;}}gel("pshin").textContent="Producto no encontrado.";}
function psClr(){gel("psi").value="";pCd=null;psCA();gel("pres").style.display="none";}
function setPM(m,btn){var btns=qsa(".mb"),i;for(i=0;i<btns.length;i++)btns[i].classList.remove("on");btn.classList.add("on");pMet=m;if(pCd)psRCh(pGMo(pCd));}
function pGMo(cd){var rows=[],i;for(i=0;i<MO.length;i++)if(MO[i][0]===cd)rows.push(MO[i]);rows.sort(function(a,b){return a[1]!==b[1]?a[1]-b[1]:a[2]-b[2];});return rows;}
function psRen(cd){
  var p=null,i,md,yrs,ph,el;pCd=cd;
  for(i=0;i<CAT.length;i++)if(CAT[i][0]===cd){p=CAT[i];break;}if(!p)return;
  gel("pres").style.display="block";gel("prcd").textContent=p[0];gel("prnm").textContent=p[1];
  gel("prtgs").innerHTML="<span class=\"prtg\">Familia: <b>"+p[2]+"</b></span><span class=\"prtg\">Cat: <b>"+p[3]+"</b></span><span class=\"prtg\">Desde: <b>"+p[7]+"</b></span><span class=\"prtg\">Hasta: <b>"+p[8]+"</b></span>";
  gel("pk-un").textContent=fN(p[4]);gel("pk-m2").textContent=fD(p[5]);gel("pk-hh").textContent=fD(p[6]);gel("pk-rng").textContent=p[7].slice(0,4)+" - "+p[8].slice(0,4);gel("pk-days").textContent=p[9]+" meses con produccion";
  md=pGMo(cd);yrs=uniq(md.map(function(r){return r[1];})).sort(function(a,b){return a-b;});
  pSYrs=yrs.slice();ph=gel("yrps");ph.innerHTML="";
  for(i=0;i<yrs.length;i++){el=document.createElement("div");el.className="yrp on";el.textContent=yrs[i];(function(y){el.onclick=function(){psTgY(this,y);};})(yrs[i]);ph.appendChild(el);}
  psRCh(md);psRMo(md,p[4]);psRAn(md);
  gel("pres").scrollIntoView({behavior:"smooth",block:"start"});
}
function psTgY(el,y){var md,p,i;if(hv(pSYrs,y)){if(pSYrs.length<=1)return;pSYrs=rv(pSYrs,y);el.classList.remove("on");}else{pSYrs.push(y);el.classList.add("on");}md=pGMo(pCd);p=null;for(i=0;i<CAT.length;i++)if(CAT[i][0]===pCd){p=CAT[i];break;}psRCh(md);psRMo(md,p?p[4]:0);}
function psRCh(md){var rows=[],cols={un:"#4f8ef7",m2:"#27c4a0",hh:"#9b7dea"},mi={un:3,m2:4,hh:5},labels=[],vals=[],i,col;for(i=0;i<md.length;i++)if(hv(pSYrs,md[i][1]))rows.push(md[i]);for(i=0;i<rows.length;i++){labels.push(MES[rows[i][2]]+" "+String(rows[i][1]).slice(2));vals.push(rows[i][mi[pMet]]||0);}if(pCh){pCh.destroy();pCh=null;}col=cols[pMet];pCh=new Chart(gel("pc"),{type:"bar",data:{labels:labels,datasets:[{data:vals,backgroundColor:col+"cc",borderRadius:3,borderSkipped:false}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return fN(c.parsed.y)+" "+(pMet==="un"?"unidades":pMet==="m2"?"m2":"HH");}}}},scales:{x:{ticks:{color:"#3e4459",font:{size:9},maxRotation:45,autoSkip:true,maxTicksLimit:24},grid:{color:"rgba(255,255,255,0.03)"}},y:{ticks:{color:"#3e4459",font:{size:10},callback:function(v){return fN(v);}},grid:{color:"rgba(255,255,255,0.03)"}}}}});}
function psRMo(md,totUn){var rows=[],i,r,maxU=0,tb="",tU=0,tM=0,tH=0,pct,bw;for(i=0;i<md.length;i++)if(hv(pSYrs,md[i][1]))rows.push(md[i]);rows.sort(function(a,b){return a[1]!==b[1]?b[1]-a[1]:b[2]-a[2];});for(i=0;i<rows.length;i++)if(rows[i][3]>maxU)maxU=rows[i][3];for(i=0;i<rows.length;i++){r=rows[i];pct=totUn>0?((r[3]/totUn)*100).toFixed(1):0;bw=maxU>0?Math.round((r[3]/maxU)*48):0;tU+=r[3];tM+=r[4];tH+=r[5];tb+="<tr><td>"+MES[r[2]]+" "+r[1]+"</td><td><span style=\"display:inline-block;width:"+bw+"px;height:3px;background:#4f8ef755;border-radius:2px;vertical-align:middle;margin-right:4px\"></span>"+fN(r[3])+"</td><td>"+fD(r[4])+"</td><td>"+fD(r[5])+"</td><td style=\"color:#7c849a\">"+pct+"%</td></tr>";}gel("ptb").innerHTML=tb;gel("ptf").innerHTML="<tr><td>TOTAL</td><td>"+fN(tU)+"</td><td>"+fD(tM)+"</td><td>"+fD(tH)+"</td><td>100%</td></tr>";}
function psRAn(md){var ym={},i,r,yrs,y,maxU=0,tb="",bw,pr;for(i=0;i<md.length;i++){r=md[i];if(!ym[r[1]])ym[r[1]]={yr:r[1],un:0,m2:0,hh:0,mos:0};ym[r[1]].un+=r[3];ym[r[1]].m2+=r[4];ym[r[1]].hh+=r[5];ym[r[1]].mos++;}yrs=Object.keys(ym).sort(function(a,b){return b-a;});for(i=0;i<yrs.length;i++)if(ym[yrs[i]].un>maxU)maxU=ym[yrs[i]].un;for(i=0;i<yrs.length;i++){y=ym[yrs[i]];bw=maxU>0?Math.round((y.un/maxU)*48):0;pr=y.mos>0?Math.round(y.un/y.mos):0;tb+="<tr><td><b>"+y.yr+"</b></td><td><span style=\"display:inline-block;width:"+bw+"px;height:3px;background:#4f8ef755;border-radius:2px;vertical-align:middle;margin-right:4px\"></span>"+fN(y.un)+"</td><td>"+fD(y.m2)+"</td><td>"+fD(y.hh)+"</td><td>"+y.mos+" meses</td><td>"+fN(pr)+"/mes</td></tr>";}gel("pab").innerHTML=tb;}

document.addEventListener("click",function(e){if(!e.target.closest(".psiw"))psCA();});
window.onload=function(){syncData();};
