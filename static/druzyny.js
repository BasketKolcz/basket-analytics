(function(){
var DB=window.DB_INIT||[];
var PLAYERS=window.PLAYERS_INIT||{};
var LOGOS=window.LOGOS_INIT||{};
var KI=-1,KEXT=false,KN="",KSEZON="",STEP=-1;
var TEAM_KI=-1,TEAM_S="",TEAM_D="";
var _Q=new URLSearchParams(location.search);
var EMBED=_Q.get("embed")==="1";
var NEW_PENDING=-1;   // indeks klubu tworzonego wlasnie przez ?new=1 (jeszcze nie zatwierdzonego)
var EDIT_PI=-1,pendDel=-1,pendDelPl=-1;
var SP_EDIT=false;    // profil zawodnika: tryb edycji na miejscu
var SP_DIRTY=false;   // sa niezapisane zmiany (ocena w zakladce Profil)
var CARD_PALS=[
  {bg:"#1a2b4a",acc:"#EF9F27"},
  {bg:"#085041",acc:"#d4f5e6"},
  {bg:"#7c1616",acc:"#ffd966"},
  {bg:"#4a3fa8",acc:"#d9d4ff"},
  {bg:"#1a5c8a",acc:"#b8dcf5"},
  {bg:"#3d631a",acc:"#c3e88d"},
  {bg:"#8a4f1a",acc:"#ffcc88"},
  {bg:"#2d6e76",acc:"#b8edf5"}
];
var SILHOUETTES=[
  '<img src="/static/sil0.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil1.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil2.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil3.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil4.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil5.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil6.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">',
  '<img src="/static/sil7.png" style="display:block;margin:0 auto 0;max-height:74px;max-width:90%;object-fit:contain;pointer-events:none">'
];
function playerSilhIdx(p){
  var s=(p.imie||'')+(p.nazwisko||'');
  var h=0;
  for(var i=0;i<s.length;i++) h=(h*31+s.charCodeAt(i))&0xFFFF;
  return h%SILHOUETTES.length;
}

function saveDB(){
  fetch("/druzyny/save",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({kluby:DB,players:PLAYERS})});
}

function esc(s){
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function draw(){
  var root=document.getElementById("tree-root");
  if(!root)return;
  if(!DB.length){
    root.innerHTML='<div style="padding:40px;text-align:center;color:#bbb;font-size:12px">Brak klubów. Kliknij + Dodaj.</div>';
    return;
  }
  var h="";
  for(var ki=0;ki<DB.length;ki++){
    var k=DB[ki],ext=k.ext||false;
    h+='<div style="margin-bottom:10px;border-radius:9px;border:1px solid #eee;overflow:hidden">';
    h+='<div style="background:#1a2b4a;padding:8px 12px;display:flex;justify-content:space-between;align-items:center">';
    h+='<div style="display:flex;align-items:center;gap:7px;font-size:12px;font-weight:500;color:#fff">';
    h+='<span style="width:8px;height:8px;border-radius:50%;background:'+(ext?'#5DCAA5':'#EF9F27')+';display:inline-block"></span>'+esc(k.name)+'</div>';
    h+='<div style="position:relative">';
    h+='<button class="mbtn" data-ki="'+ki+'" style="background:rgba(255,255,255,.18);border:none;color:#fff;cursor:pointer;padding:3px 9px;border-radius:5px;font-size:15px">&#8943;</button>';
    h+='<div id="mdd'+ki+'" style="display:none;position:absolute;top:calc(100% + 4px);right:0;background:#fff;border:1px solid #e0e0e0;border-radius:8px;min-width:110px;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.1)">';
    h+='<div class="dde" data-ki="'+ki+'" style="padding:9px 14px;cursor:pointer;font-size:12px;color:#222">Edytuj</div>';
    h+='<div style="height:1px;background:#f0f0f0"></div>';
    h+='<div class="ddd" data-ki="'+ki+'" style="padding:9px 14px;cursor:pointer;font-size:12px;color:#A32D2D">Usuń</div>';
    h+='</div></div></div>';
    var sezony=k.sezony||{};
    for(var s in sezony){
      h+='<div style="padding:5px 12px;background:#f8f9fa;border-top:1px solid #eee;font-size:11px;font-weight:500;color:#666">'+esc(s)+'</div>';
      h+='<div style="padding:6px 12px 8px;display:flex;flex-wrap:wrap;gap:5px">';
      var tt=sezony[s]||[];
      if(!tt.length) h+='<span style="font-size:11px;color:#ccc">brak drużyn</span>';
      for(var ti=0;ti<tt.length;ti++){
        h+='<button class="tbtn" data-ki="'+ki+'" data-s="'+encodeURIComponent(s)+'" data-d="'+encodeURIComponent(tt[ti])+'" style="display:inline-flex;align-items:center;border:1px solid #e0e0e0;border-radius:20px;padding:4px 12px;font-size:11px;color:#222;cursor:pointer;background:#fff">'+esc(tt[ti])+' &#8594;</button>';
      }
      h+='</div>';
    }
    h+='</div>';
  }
  root.innerHTML=h;
  bind();
}

function bind(){
  document.querySelectorAll(".mbtn").forEach(function(b){
    b.addEventListener("click",function(e){
      e.stopPropagation();
      var ki=parseInt(this.dataset.ki);
      document.querySelectorAll("[id^=mdd]").forEach(function(d){d.style.display="none";});
      document.getElementById("mdd"+ki).style.display="block";
    });
  });
  document.querySelectorAll(".dde").forEach(function(el){el.addEventListener("click",function(){doEdit(parseInt(this.dataset.ki));});});
  document.querySelectorAll(".ddd").forEach(function(el){el.addEventListener("click",function(){askDel(parseInt(this.dataset.ki));});});
  document.querySelectorAll(".tbtn").forEach(function(b){
    b.addEventListener("click",function(){
      goTeam(parseInt(this.dataset.ki),decodeURIComponent(this.dataset.s),decodeURIComponent(this.dataset.d));
    });
  });
}

document.addEventListener("click",function(){
  document.querySelectorAll("[id^=mdd]").forEach(function(d){d.style.display="none";});
  document.querySelectorAll("[id^=cdrop]").forEach(function(d){d.style.display="none";});
});

function idle(){
  // Rezygnacja z tworzenia klubu — usuwamy szkielet. saveDB() na wypadek,
  // gdyby posrednie akcje (dodanie sezonu/druzyny) zdazyly go juz utrwalic.
  if(NEW_PENDING>=0){
    if(DB[NEW_PENDING]) DB.splice(NEW_PENDING,1);
    NEW_PENDING=-1; saveDB();
  }
  STEP=-1;KI=-1;
  // W ramce (zakladka Roster) wracamy do skladu druzyny — bez opuszczania HUB-u.
  if(EMBED && TEAM_KI>=0 && DB[TEAM_KI]){ renderTeam(); return; }
  TEAM_KI=-1;
  // Poza ramka nie ma juz ekranu startowego — wracamy na liste klubow.
  (window.top||window).location.href='/klub';
}

function panel(body,foot,bar){
  var rb=document.getElementById("rbody");rb.style.padding="0";rb.style.textAlign="";rb.innerHTML=body;
  var rf=document.getElementById("rfoot");rf.innerHTML=foot;rf.style.display=foot?"flex":"none";
  var sb=document.getElementById("sbar");
  if(bar){sb.innerHTML=bar;sb.style.display="flex";}else{sb.style.display="none";}
}

function fld(id,lbl,ph,val){
  return '<div style="margin-bottom:10px"><div style="font-size:11px;color:#666;margin-bottom:3px">'+lbl+'</div>'
    +'<input id="'+id+'" placeholder="'+ph+'" value="'+esc(val||'')+'" style="width:100%;padding:7px 10px;border:1px solid #e0e0e0;border-radius:8px;font-size:12px;box-sizing:border-box"></div>';
}

function askDel(ki){
  document.querySelectorAll("[id^=mdd]").forEach(function(d){d.style.display="none";});
  pendDel=ki;
  document.getElementById("ov-title").textContent="Usunąć „"+DB[ki].name+"”?";
  document.getElementById("ov-msg").textContent="Klub zostanie usunięty. Operacji nie można cofnąć.";
  document.getElementById("ov-ok").onclick=function(){DB.splice(pendDel,1);NEW_PENDING=-1;saveDB();closeOv();draw();idle();};
  document.getElementById("ov-del").style.display="flex";
}
document.getElementById("ov-cancel").addEventListener("click",closeOv);
function closeOv(){document.getElementById("ov-del").style.display="none";}

function uploadLogo(clubName,file){
  var fd=new FormData();
  fd.append('logo',file);
  fd.append('klub',clubName);
  fetch('/druzyny/upload-logo',{method:'POST',body:fd})
    .then(function(r){return r.json();})
    .then(function(d){
      if(d.ok && d.logo){
        var safe=(clubName||'').replace(/ /g,'_').replace(/\//g,'-');
        LOGOS[safe]=d.logo;
        var prev=document.getElementById('edit-logo-img');
        if(prev){
          prev.outerHTML='<img id="edit-logo-img" src="'+d.logo+'" style="max-width:80%;max-height:68px;object-fit:contain">';
        }
        // odśwież karty na ekranie startowym (jeśli widoczne)
      } else {
        alert('Błąd wgrywania logo: '+(d.error||'nieznany'));
      }
    })
    .catch(function(err){alert('Błąd: '+(err.message||err));});
}

function doEdit(ki){
  document.querySelectorAll("[id^=mdd]").forEach(function(d){d.style.display="none";});
  KI=ki;var k=DB[KI];
  var safe=(k.name||'').replace(/ /g,'_').replace(/\//g,'-');
  var existLogo=LOGOS[safe]||'';
  var color=k.ext?'#085041':'#1a2b4a';
  var b='<div style="overflow:hidden">';
  // Nagłówek z logo
  b+='<div style="background:'+color+';height:80px;position:relative;display:flex;align-items:center;justify-content:center;overflow:hidden">';
  if(existLogo){
    b+='<img id="edit-logo-img" src="'+existLogo+'" style="max-width:80%;max-height:68px;object-fit:contain">';
  } else {
    var inits=(k.name||'?').split(' ').map(function(w){return w[0]||'';}).join('').substring(0,3).toUpperCase();
    b+='<span id="edit-logo-img" style="font-size:28px;font-weight:800;color:rgba(255,255,255,.25)">'+inits+'</span>';
  }
  b+='<label style="position:absolute;bottom:7px;right:9px;background:rgba(0,0,0,.32);color:#fff;font-size:10px;padding:3px 9px;border-radius:10px;cursor:pointer;user-select:none">';
  b+='<input id="logo-file-inp" type="file" accept=".png,.jpg,.jpeg,.webp" style="display:none">';
  b+='↑ Logo</label>';
  b+='</div>';
  b+='<div style="padding:14px 16px;overflow-y:auto;max-height:390px">';
  b+='<div style="font-size:14px;font-weight:600;color:#1a2b4a;margin-bottom:12px">Edycja: '+esc(k.name)+'</div>';
  b+='<div style="background:#f8f9fa;border-radius:9px;padding:12px;margin-bottom:12px">';
  b+='<input id="ename" value="'+esc(k.name)+'" style="width:100%;padding:7px 10px;border:1px solid #e0e0e0;border-radius:7px;font-size:12px;margin-bottom:8px;box-sizing:border-box">';
  b+='<div style="display:flex;gap:8px">';
  b+='<div id="ec" style="flex:1;padding:6px;text-align:center;border:1px solid '+(k.ext?'#e0e0e0':'#1a2b4a')+';border-radius:7px;cursor:pointer;font-size:11px;background:'+(k.ext?'#fff':'#E6F1FB')+';color:'+(k.ext?'#888':'#0C447C')+'">Klub sportowy</div>';
  b+='<div id="ee" style="flex:1;padding:6px;text-align:center;border:1px solid '+(k.ext?'#1a2b4a':'#e0e0e0')+';border-radius:7px;cursor:pointer;font-size:11px;background:'+(k.ext?'#E6F1FB':'#fff')+';color:'+(k.ext?'#0C447C':'#888')+'">Kadra</div>';
  b+='</div></div>';
  var sk=Object.keys(k.sezony||{});
  for(var si=0;si<sk.length;si++){
    var s=sk[si],tt=k.sezony[s]||[];
    b+='<div style="margin-bottom:8px;border:1px solid #eee;border-radius:8px;overflow:hidden">';
    b+='<div style="display:flex;align-items:center;gap:6px;padding:7px 10px;background:#f8f9fa">';
    b+='<input value="'+esc(s)+'" id="sn'+si+'" style="flex:1;padding:5px 9px;border:1px solid #e0e0e0;border-radius:6px;font-size:12px">';
    b+='<button class="ds" data-s="'+encodeURIComponent(s)+'" style="background:none;border:none;cursor:pointer;font-size:11px;color:#A32D2D;padding:3px 6px">Usuń</button></div>';
    b+='<div style="padding:8px 10px"><div style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:6px">';
    for(var di=0;di<tt.length;di++){
      b+='<div style="display:inline-flex;align-items:center;gap:3px;border:1px solid #eee;border-radius:14px;padding:3px 9px;font-size:11px">'+esc(tt[di]);
      b+='<button class="dd2" data-s="'+encodeURIComponent(s)+'" data-di="'+di+'" style="background:none;border:none;cursor:pointer;color:#bbb;font-size:12px;line-height:1">&#10005;</button></div>';
    }
    if(!tt.length) b+='<span style="font-size:11px;color:#ccc">Brak drużyn</span>';
    b+='</div><div style="display:flex;gap:5px">';
    b+='<input placeholder="+ Nowa drużyna..." id="nd'+si+'" style="flex:1;padding:5px 9px;border:1px solid #e0e0e0;border-radius:6px;font-size:11px">';
    b+='<button class="ad" data-s="'+encodeURIComponent(s)+'" data-si="'+si+'" style="background:#1a2b4a;color:#fff;border:none;padding:5px 10px;border-radius:6px;font-size:11px;cursor:pointer">Dodaj</button></div></div></div>';
  }
  b+='<div style="display:flex;gap:6px;margin-top:4px">';
  b+='<input id="nsi" placeholder="Nowy sezon, np. 2026/2027" style="flex:1;padding:6px 10px;border:1px solid #e0e0e0;border-radius:7px;font-size:12px">';
  b+='<button id="as" style="background:#EF9F27;color:#fff;border:none;padding:6px 12px;border-radius:7px;font-size:12px;cursor:pointer;white-space:nowrap">Dodaj sezon</button></div></div></div>';
  var f='<button id="ekdel" style="background:none;border:1px solid #f0d5d5;color:#A32D2D;padding:8px 12px;border-radius:8px;cursor:pointer;font-size:12px">Usuń klub</button>';
  f+='<button id="ec2" style="flex:1;background:none;border:1px solid #ddd;color:#888;padding:8px;border-radius:8px;cursor:pointer;font-size:12px">Anuluj</button>';
  f+='<button id="es" style="flex:2;background:#EF9F27;color:#fff;border:none;padding:8px;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600">Zapisz</button>';
  panel(b,f,null);
  // Logo upload
  var logoInp=document.getElementById('logo-file-inp');
  if(logoInp) logoInp.addEventListener('change',function(){
    if(this.files[0]) uploadLogo(DB[KI].name,this.files[0]);
  });
  var ec=document.getElementById("ec"),ee=document.getElementById("ee");
  if(ec) ec.addEventListener("click",function(){DB[KI].ext=false;doEdit(KI);});
  if(ee) ee.addEventListener("click",function(){DB[KI].ext=true;doEdit(KI);});
  document.querySelectorAll(".ds").forEach(function(b){b.addEventListener("click",function(){delete DB[KI].sezony[decodeURIComponent(this.dataset.s)];saveDB();doEdit(KI);draw();});});
  document.querySelectorAll(".dd2").forEach(function(b){b.addEventListener("click",function(){DB[KI].sezony[decodeURIComponent(this.dataset.s)].splice(parseInt(this.dataset.di),1);saveDB();doEdit(KI);draw();});});
  document.querySelectorAll(".ad").forEach(function(b){b.addEventListener("click",function(){var s=decodeURIComponent(this.dataset.s),si=parseInt(this.dataset.si);var v=document.getElementById("nd"+si);if(!v||!v.value.trim())return;DB[KI].sezony[s].push(v.value.trim());saveDB();doEdit(KI);draw();});});
  var as2=document.getElementById("as");
  if(as2) as2.addEventListener("click",function(){var v=document.getElementById("nsi");if(!v||!v.value.trim())return;if(!DB[KI].sezony[v.value.trim()])DB[KI].sezony[v.value.trim()]=[];saveDB();doEdit(KI);draw();});
  var _ekd=document.getElementById("ekdel");
  if(_ekd) _ekd.addEventListener("click",function(){askDel(KI);});
  document.getElementById("ec2").addEventListener("click",idle);
  document.getElementById("es").addEventListener("click",function(){
    var n=document.getElementById("ename");if(n&&n.value.trim())DB[KI].name=n.value.trim();
    NEW_PENDING=-1;   // zatwierdzony — nie kasujemy go przy wyjsciu
    var sk2=Object.keys(DB[KI].sezony||{}),ns={};
    for(var i=0;i<sk2.length;i++){var sn=document.getElementById("sn"+i);var k2=sn&&sn.value.trim()?sn.value.trim():sk2[i];ns[k2]=DB[KI].sezony[sk2[i]];}
    DB[KI].sezony=ns;saveDB();draw();idle();
  });
}

function goTeam(ki,s,d){
  TEAM_KI=ki;TEAM_S=s;TEAM_D=d;
  var k=String(ki);
  if(!PLAYERS[k])PLAYERS[k]={};
  if(!PLAYERS[k][s])PLAYERS[k][s]={};
  if(!PLAYERS[k][s][d])PLAYERS[k][s][d]=[];
  renderTeam();
}

function renderTeam(){
  var k=DB[TEAM_KI];
  var ps=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[];
  var ext=k?(k.ext||false):false;
  var color=ext?'#085041':'#1a2b4a';
  var safe=k?((k.name||'').replace(/ /g,'_').replace(/\//g,'-')):'';
  var logo=safe?LOGOS[safe]||'':'';

  // Wszystkie sezony tego klubu z drużyną TEAM_D, od najnowszego
  var allSeasons=Object.keys(k?k.sezony:{}).filter(function(s){
    return(k.sezony[s]||[]).indexOf(TEAM_D)>=0;
  }).sort(function(a,b){return b.localeCompare(a);});

  var initK=(k?k.name||'?':'?').split(' ').map(function(w){return w[0]||'';}).join('').substring(0,3).toUpperCase();

  var h='<div style="display:flex;flex-direction:column;height:100%">';

  // HEADER CARD
  h+='<div style="padding:12px 14px;border-bottom:1px solid #f0f0f0;display:flex;align-items:center;gap:10px;flex-shrink:0">';
  h+='<div style="width:42px;height:42px;border-radius:9px;background:'+color+';display:flex;align-items:center;justify-content:center;flex-shrink:0;overflow:hidden">';
  if(logo){
    h+='<img src="'+logo+'" style="width:38px;height:38px;object-fit:contain">';
  } else {
    h+='<span style="font-size:11px;font-weight:700;color:rgba(255,255,255,.85)">'+initK+'</span>';
  }
  h+='</div>';
  h+='<div style="flex:1;min-width:0">';
  h+='<div style="font-size:14px;font-weight:700;color:#1a2b4a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc(TEAM_D)+'</div>';
  h+='<div style="font-size:10px;color:#999;margin-top:1px">'+esc(k?k.name:'')+' &middot; '+esc(TEAM_S)+' &middot; '+ps.length+' zawodników</div>';
  h+='</div>';
  h+='<button id="rt-edit" style="background:#f3f4f6;border:1px solid #e5e7eb;color:#374151;font-size:11px;font-weight:500;padding:6px 11px;border-radius:7px;cursor:pointer;white-space:nowrap">Edytuj</button>';
  h+='<button id="apb" style="background:#EF9F27;color:#fff;border:none;font-size:11px;font-weight:600;padding:6px 11px;border-radius:7px;cursor:pointer;white-space:nowrap">+ Zawodnik</button>';
  h+='</div>';

  // SEASON PILLS
  if(allSeasons.length>1){
    h+='<div style="padding:8px 14px;border-bottom:1px solid #f0f0f0;display:flex;gap:6px;overflow-x:auto;flex-shrink:0">';
    allSeasons.forEach(function(s){
      var active=s===TEAM_S;
      h+='<button class="sp-pill" data-s="'+encodeURIComponent(s)+'" style="padding:4px 14px;border-radius:20px;border:1px solid '+(active?color:'#e5e7eb')+';background:'+(active?color:'#fff')+';color:'+(active?'#fff':'#555')+';font-size:11px;font-weight:'+(active?'600':'400')+';cursor:pointer;flex-shrink:0;white-space:nowrap">'+esc(s)+'</button>';
    });
    h+='</div>';
  }

  // PLAYER GRID
  h+='<div style="flex:1;overflow-y:auto;padding:12px 14px">';
  if(!ps.length){
    h+='<div style="text-align:center;padding:40px 20px;color:#ccc;font-size:12px">Brak zawodników. Kliknij + Zawodnik.</div>';
  } else {
    var pal=CARD_PALS[TEAM_KI%CARD_PALS.length];
    h+='<div id="pl-grid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:10px">';
    ps.forEach(function(p,pi){
      var akt=p.aktywny!==false;
      var cbg=akt?pal.bg:"#8a94a0";
      var cacc=akt?pal.acc:"rgba(255,255,255,.6)";
      var silh=SILHOUETTES[playerSilhIdx(p)];
      h+='<div class="pl-card" data-pi="'+pi+'" style="border-radius:14px;overflow:hidden;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.1);background:#fff">';
      h+='<div style="background:'+cbg+';padding:8px 8px 0;text-align:center;position:relative;min-height:86px">';
      h+='<div style="position:absolute;top:6px;left:7px;background:rgba(0,0,0,.25);border-radius:7px;padding:1px 7px;font-size:12px;font-weight:800;color:'+cacc+';letter-spacing:.5px;line-height:1.5">'+esc(String(p.num||'?'))+'</div>';
      h+=silh;
      h+='</div>';
      h+='<div style="padding:7px 8px 8px;background:#fff">';
      h+='<div style="font-size:11.5px;font-weight:700;color:'+(akt?'#1a2b4a':'#bbb')+';white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc(p.nazwisko||'')+'</div>';
      h+='<div style="font-size:10.5px;color:#666;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-bottom:4px">'+esc(p.imie||'')+'</div>';
      h+='<div style="display:flex;justify-content:space-between;align-items:center">';
      if(p.poz) h+='<span style="font-size:9px;color:#888;background:#f0f2f7;padding:1px 5px;border-radius:6px;overflow:hidden;text-overflow:ellipsis;max-width:55px;font-weight:600">'+esc(p.poz)+'</span>';
      else h+='<span></span>';
      h+='<span style="font-size:8px;font-weight:600;padding:1px 5px;border-radius:7px;background:'+(akt?'#E1F5EE':'#f5f5f5')+';color:'+(akt?'#0F6E56':'#bbb')+'">'+(akt?'Akt.':'Nier.')+'</span>';
      h+='</div></div></div>';
    });
    h+='</div>';
  }
  h+='</div>';

  // BOTTOM BAR
  h+='<div style="padding:8px 14px;border-top:1px solid #f0f0f0;display:flex;align-items:center;gap:6px;flex-shrink:0">';
  h+='<div style="flex:1;position:relative">';
  h+='<svg style="position:absolute;left:8px;top:50%;transform:translateY(-50%);width:11px;height:11px;stroke:#bbb;fill:none;stroke-width:2;pointer-events:none" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>';
  h+='<input id="pl-q" placeholder="Szukaj zawodnika..." style="width:100%;padding:5px 8px 5px 25px;border:1px solid #e8edf5;border-radius:8px;font-size:11px;box-sizing:border-box">';
  h+='</div>';
  h+='<button id="rt-import" style="background:none;border:1px solid #e5e7eb;color:#666;font-size:10px;padding:4px 10px;border-radius:7px;cursor:pointer;white-space:nowrap">&#8595; Import</button>';
  h+='</div>';

  // IMPORT PANEL (hidden)
  h+='<div id="import-panel" style="display:none;padding:10px 14px;border-top:1px solid #f0f0f0;flex-shrink:0">';
  h+='<div id="dz" style="border:2px dashed #e0e0e0;border-radius:10px;padding:16px;text-align:center;cursor:pointer;background:#fafafa;margin-bottom:8px">';
  h+='<div style="font-size:20px;margin-bottom:4px">&#128194;</div>';
  h+='<div style="font-size:12px;font-weight:500;color:#444">Przeciągnij lub kliknij</div>';
  h+='<div style="font-size:10px;color:#aaa">.xlsx, .csv</div></div>';
  h+='<div style="display:flex;gap:6px">';
  h+='<a href="/template/sklad" style="flex:1;background:#eef2ff;border:1px solid #c8d5f5;color:#1a2b4a;padding:7px;border-radius:8px;font-size:11px;text-align:center;text-decoration:none;display:block;font-weight:500">&#8595; Excel</a>';
  h+='<a href="/template/sklad?fmt=csv" style="flex:1;background:#eef2ff;border:1px solid #c8d5f5;color:#1a2b4a;padding:7px;border-radius:8px;font-size:11px;text-align:center;text-decoration:none;display:block;font-weight:500">&#8595; CSV</a>';
  h+='</div></div>';

  h+='</div>';

  var kname=esc(k?k.name:'');
  var sbar='<div style="display:flex;align-items:center;width:100%;gap:5px;overflow:hidden">'
    +'<span id="bc-home" style="font-size:11px;cursor:pointer;color:#EF9F27;flex-shrink:0">Drużyny</span>'
    +'<span style="font-size:11px;color:#ddd;flex-shrink:0">&rsaquo;</span>'
    +'<span style="font-size:11px;color:#888;flex-shrink:0">'+kname+'</span>'
    +'<span style="font-size:11px;color:#ddd;flex-shrink:0">&rsaquo;</span>'
    +'<span style="font-size:11px;color:#444;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc(TEAM_D)+'</span>'
    +'</div>';
  panel(h,"",EMBED?"":sbar);

  var _bch=document.getElementById("bc-home");if(_bch)_bch.addEventListener("click",idle);
  document.getElementById("rt-edit").addEventListener("click",function(){doEdit(TEAM_KI);});

  document.getElementById("apb").addEventListener("click",function(){
    EDIT_PI=-1;
    document.getElementById("pm-title").textContent="Nowy zawodnik";
    document.getElementById("pm-sub").textContent=TEAM_D+" · "+TEAM_S;
    document.getElementById("pm-imie").value="";
    document.getElementById("pm-nazw").value="";
    document.getElementById("pm-num").value="";
    document.getElementById("pm-poz").value="";
    var akt=document.getElementById("pm-aktywny");if(akt)akt.value="1";
    clearPersonFields();
    document.getElementById("ov-player").style.display="flex";
  });

  document.querySelectorAll('.sp-pill').forEach(function(btn){
    btn.addEventListener('click',function(){
      var ns=decodeURIComponent(this.dataset.s);
      if(ns===TEAM_S)return;
      TEAM_S=ns;
      var ki=String(TEAM_KI);
      if(!PLAYERS[ki])PLAYERS[ki]={};
      if(!PLAYERS[ki][ns])PLAYERS[ki][ns]={};
      if(!PLAYERS[ki][ns][TEAM_D])PLAYERS[ki][ns][TEAM_D]=[];
      renderTeam();
    });
  });

  document.querySelectorAll('.pl-card').forEach(function(card){
    card.addEventListener('click',function(){showPlayer(parseInt(this.dataset.pi));});
  });

  var qEl=document.getElementById("pl-q");
  if(qEl){
    qEl.addEventListener("input",function(){
      var q=this.value.trim().toLowerCase();
      document.querySelectorAll('.pl-card').forEach(function(card){
        var pi=parseInt(card.dataset.pi);
        var p=(PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[])[pi];
        if(!p){card.style.display='none';return;}
        card.style.display=(!q||((p.imie||'')+(p.nazwisko||'')).toLowerCase().indexOf(q)>=0)?'':'none';
      });
    });
  }

  document.getElementById("rt-import").addEventListener("click",function(){
    var ip=document.getElementById("import-panel");
    var vis=ip.style.display==='block';
    ip.style.display=vis?'none':'block';
    this.style.background=vis?'none':'#eef2ff';
    this.style.borderColor=vis?'#e5e7eb':'#c8d5f5';
  });

  var dz=document.getElementById("dz");
  if(dz){
    dz.addEventListener("click",function(){var inp=document.createElement("input");inp.type="file";inp.accept=".xlsx,.csv";inp.addEventListener("change",function(){if(this.files[0])importF(this.files[0]);});inp.click();});
    dz.addEventListener("dragover",function(e){e.preventDefault();this.style.borderColor="#EF9F27";this.style.background="#FAEEDA";});
    dz.addEventListener("dragleave",function(){this.style.borderColor="#e0e0e0";this.style.background="#fafafa";});
    dz.addEventListener("drop",function(e){e.preventDefault();this.style.borderColor="#e0e0e0";this.style.background="#fafafa";var f=e.dataTransfer.files[0];if(f)importF(f);});
  }

  window._plSortAsc=true;
  window._plFilterAkt=false;
}

function plSort(){
  window._plSortAsc=!window._plSortAsc;
  var lbl=document.getElementById("pl-sl");
  if(lbl)lbl.textContent=window._plSortAsc?"A→Z":"Z→A";
  renderList();
}

function plFilter(){
  window._plFilterAkt=!window._plFilterAkt;
  var btn=document.getElementById("pl-af");
  if(btn){btn.textContent=window._plFilterAkt?"Wszyscy":"Aktywni";btn.style.background=window._plFilterAkt?"#E1F5EE":"#f8f9fa";btn.style.color=window._plFilterAkt?"#0F6E56":"#666";btn.style.borderColor=window._plFilterAkt?"#5DCAA5":"#e0e0e0";}
  renderList();
}

function renderList(){
  var listEl=document.getElementById("pl-list");
  if(!listEl)return;
  var ps=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[];
  var qEl=document.getElementById("pl-q");
  var q=qEl?(qEl.value||"").trim().toLowerCase():"";
  var ext=DB[TEAM_KI]?(DB[TEAM_KI].ext||false):false;
  var color=ext?'#085041':'#1a2b4a';
  var filtered=ps.map(function(p,i){return{p:p,i:i};}).filter(function(x){
    if(window._plFilterAkt&&x.p.aktywny===false)return false;
    if(!q)return true;
    return((x.p.imie||"")+" "+(x.p.nazwisko||"")).toLowerCase().indexOf(q)>=0;
  });
  filtered.sort(function(a,b){
    var ca=((a.p.nazwisko||"")+" "+(a.p.imie||"")).toLowerCase();
    var cb=((b.p.nazwisko||"")+" "+(b.p.imie||"")).toLowerCase();
    return window._plSortAsc?ca.localeCompare(cb,"pl"):cb.localeCompare(ca,"pl");
  });
  var cnt=document.getElementById("pl-cnt");
  if(cnt)cnt.textContent=filtered.length+(filtered.length===1?" zawodnik":" zawodników");
  if(!filtered.length){
    listEl.innerHTML='<div style="text-align:center;padding:30px;color:#ccc;font-size:12px">'+(q?"Brak wyników dla \""+esc(q)+"\".":"Brak zawodników. Kliknij + Dodaj.")+'</div>';
    return;
  }
  var pal=CARD_PALS[TEAM_KI%CARD_PALS.length];
  var h='<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(138px,1fr));gap:12px;padding:14px">';
  filtered.forEach(function(x){
    var p=x.p,pi=x.i;
    var akt=p.aktywny!==false;
    var cbg=akt?pal.bg:"#8a94a0";
    var cacc=akt?pal.acc:"rgba(255,255,255,.6)";
    var silh=SILHOUETTES[playerSilhIdx(p)];
    h+='<div class="pl-row" data-pi="'+pi+'" style="border-radius:14px;overflow:hidden;cursor:pointer;box-shadow:0 2px 10px rgba(0,0,0,.12);background:#fff">';
    h+='<div style="background:'+cbg+';padding:10px 10px 0;text-align:center;position:relative;min-height:96px">';
    h+='<div style="position:absolute;top:8px;left:9px;background:rgba(0,0,0,.25);border-radius:8px;padding:2px 8px;font-size:13px;font-weight:800;color:'+cacc+';letter-spacing:.5px;line-height:1.5">'+esc(String(p.num||'?'))+'</div>';
    h+='<button class="pl-del" data-pi="'+pi+'" style="position:absolute;top:6px;right:7px;background:rgba(0,0,0,.2);border:none;cursor:pointer;padding:3px 6px;border-radius:7px;color:rgba(255,255,255,.75);font-size:12px;line-height:1">&#128465;</button>';
    h+=silh;
    h+='</div>';
    h+='<div style="padding:9px 10px 10px;background:#fff">';
    h+='<div style="font-size:12.5px;font-weight:700;color:'+(akt?'#1a2b4a':'#bbb')+';white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc(p.nazwisko||"")+'</div>';
    h+='<div style="font-size:11px;color:#666;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-bottom:6px">'+esc(p.imie||"")+'</div>';
    h+='<div style="display:flex;justify-content:space-between;align-items:center">';
    if(p.poz) h+='<span style="font-size:10px;font-weight:600;color:#888;background:#f0f2f7;padding:1px 6px;border-radius:7px;overflow:hidden;text-overflow:ellipsis;max-width:60px">'+esc(p.poz)+'</span>';
    else h+='<span></span>';
    h+='<span style="font-size:9px;font-weight:600;padding:2px 6px;border-radius:8px;background:'+(akt?'#E1F5EE':'#f5f5f5')+';color:'+(akt?'#0F6E56':'#bbb')+'">'+(akt?'Aktywny':'Nieaktywny')+'</span>';
    h+='</div></div></div>';
  });
  h+='</div>';
  listEl.innerHTML=h;

  listEl.querySelectorAll('.pl-row').forEach(function(card){
    card.addEventListener('click',function(e){
      if(e.target.classList.contains('pl-del')||(e.target.closest&&e.target.closest('.pl-del')))return;
      showPlayer(parseInt(this.dataset.pi));
    });
  });
  listEl.querySelectorAll('.pl-del').forEach(function(btn){
    btn.addEventListener('click',function(e){
      e.stopPropagation();
      pendDelPl=parseInt(this.dataset.pi);
      document.getElementById("ov-title").textContent="Usunąć zawodnika?";
      document.getElementById("ov-msg").textContent="Operacji nie można cofnąć.";
      document.getElementById("ov-ok").onclick=function(){PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D].splice(pendDelPl,1);closeOv();saveDB();renderTeam();};
      document.getElementById("ov-del").style.display="flex";
    });
  });
}

// Wzrost tego samego zawodnika we wszystkich sezonach jego klubu.
// Dane mamy już w przeglądarce — nie potrzeba osobnego zapytania.
// Trajektoria pomiaru jako mały wykres liniowy — bez bibliotek zewnętrznych.
function _sparkline(punkty, kolor, lepiejWiecej){
  if(punkty.length<2) return '';
  var W=190,Hh=44,pad=5;
  var v=punkty.map(function(p){return p.v;});
  var mn=Math.min.apply(null,v), mx=Math.max.apply(null,v);
  if(mx===mn){mx=mn+1;mn=mn-1;}
  var xs=function(i){return pad+i*(W-2*pad)/(punkty.length-1);};
  var ys=function(val){return Hh-pad-((val-mn)/(mx-mn))*(Hh-2*pad);};
  var d='',kolka='';
  punkty.forEach(function(p,i){
    d+=(i?' L':'M')+xs(i).toFixed(1)+' '+ys(p.v).toFixed(1);
    var ost=(i===punkty.length-1);
    kolka+='<circle cx="'+xs(i).toFixed(1)+'" cy="'+ys(p.v).toFixed(1)+'" r="'+(ost?3:2)+'" fill="'+(ost?kolor:'#c8cfda')+'"></circle>';
  });
  return '<svg width="'+W+'" height="'+Hh+'" style="display:block">'
    +'<path d="'+d+'" fill="none" stroke="'+kolor+'" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"></path>'
    +kolka+'</svg>';
}

function rysujRozwoj(d,kolor){
  if(!d||!d.ok) return '<div style="color:#8b1a1a;font-size:12px;padding:20px">'+((d&&d.error)||'Brak danych')+'</div>';
  if(!d.terminy.length) return '<div style="color:#9aa5b5;font-size:12.5px;text-align:center;padding:30px">'
    +'Brak pomiarów tego zawodnika.<br><span style="font-size:11.5px">Dodaj je w zakładce Zarządzanie → Pomiary.</span></div>';

  var h='';

  // ── tempo wzrostu — liczone z dat, nie z liczby pomiarów ──────────────────
  var wz=[];
  d.terminy.forEach(function(t){
    var v=d.pomiary[t.id+'|wzrost'];
    if(v!==undefined) wz.push({data:t.data,v:v});
  });
  if(wz.length>=2){
    var a=wz[wz.length-2], b=wz[wz.length-1];
    var dni=(new Date(b.data)-new Date(a.data))/86400000;
    if(dni>20){
      var tempo=(b.v-a.v)/(dni/365.25);
      var szybko=tempo>=7;
      h+='<div style="display:flex;align-items:center;gap:12px;background:'+(szybko?'#fff4d6':'#f4f7fb')
        +';border:1px solid '+(szybko?'#f0d9a0':'#e2e8f2')+';border-radius:10px;padding:11px 14px;margin-bottom:14px">'
        +'<div><div style="font-size:9px;font-weight:700;letter-spacing:.06em;color:#b0b8c8;text-transform:uppercase">Tempo wzrostu</div>'
        +'<div style="font-size:19px;font-weight:800;color:'+(szybko?'#7a5200':'#1a2b4a')+'">'
        +(tempo>0?'+':'')+tempo.toFixed(1)+' <span style="font-size:11px;font-weight:600">cm/rok</span></div></div>'
        +'<div style="font-size:11.5px;color:'+(szybko?'#7a5200':'#8a93a5')+';line-height:1.45">'
        +(szybko
          ? 'Faza szybkiego wzrostu. Zwróć uwagę na obciążenia — w tym okresie rośnie ryzyko przeciążeń, a koordynacja bywa przejściowo słabsza.'
          : 'Przyrost w normie dla tego okresu.')
        +'</div></div>';
    }
  }

  // ── karty parametrów, w dwóch grupach ────────────────────────────────────
  ['antropometria','motoryka'].forEach(function(gr){
    var lista=d.parametry.filter(function(p){
      if(p.grupa!==gr) return false;
      return d.terminy.some(function(t){return d.pomiary[t.id+'|'+p.kod]!==undefined;});
    });
    if(!lista.length) return;
    h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin:4px 0 8px">'
      +(gr==='antropometria'?'Antropometria':'Testy motoryczne')+'</div>'
      +'<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(215px,1fr));gap:10px;margin-bottom:16px">';

    lista.forEach(function(p){
      var pk=[];
      d.terminy.forEach(function(t){
        var v=d.pomiary[t.id+'|'+p.kod];
        if(v!==undefined) pk.push({data:t.data,v:v,opis:t.opis});
      });
      var ost=pk[pk.length-1], poprz=pk.length>1?pk[pk.length-2]:null;
      var delta=poprz?(ost.v-poprz.v):null;
      var lepiej=(delta===null)?null:((p.kier==='mniej')?(delta<0):(delta>0));
      var kd=(delta===null||delta===0)?'#9aa5b5':(lepiej?'#1a6b3c':'#8b1a1a');
      var r=d.ranking[p.kod];

      h+='<div style="border:1px solid #e8edf5;border-radius:11px;padding:11px 12px;background:#fff">'
        +'<div style="display:flex;justify-content:space-between;align-items:baseline;gap:6px">'
        +'<span style="font-size:11px;font-weight:700;color:#1a2b4a;line-height:1.25">'+esc(p.nazwa)+'</span>'
        +(p.liczony?'<span style="font-size:8px;color:#8a93a5;background:#eef2ff;padding:1px 5px;border-radius:10px">liczony</span>':'')
        +'</div>'
        +'<div style="display:flex;align-items:baseline;gap:7px;margin:5px 0 2px">'
        +'<span style="font-size:21px;font-weight:800;color:#1a2b4a;font-variant-numeric:tabular-nums">'+ost.v+'</span>'
        +'<span style="font-size:10px;color:#9aa5b5">'+esc(p.jed||'')+'</span>'
        +(delta!==null?'<span style="font-size:11.5px;font-weight:700;color:'+kd+'">'+(delta>0?'+':'')+(Math.round(delta*100)/100)+'</span>':'')
        +'</div>'
        +'<div style="font-size:9.5px;color:#b0b8c8;margin-bottom:6px">'+ost.data+(pk.length>1?(' · '+pk.length+' pomiary'):'')+'</div>'
        +_sparkline(pk,kolor,p.kier!=='mniej')
        +(r?('<div style="margin-top:7px;padding-top:7px;border-top:1px solid #f0f4ff;font-size:10.5px;color:#6b7c95">'
            +'<b style="color:#1a2b4a">'+r.pozycja+'.</b> z '+r.z+' w drużynie'
            +(r.pozycja===1?' <span style="color:#c47d12;font-weight:700">najlepszy</span>':'')
            +'</div>'):'')
        +'</div>';
    });
    h+='</div>';
  });

  // ── tabela terminów ──────────────────────────────────────────────────────
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin:2px 0 7px">Terminy pomiarowe</div>'
   +'<div style="display:flex;flex-wrap:wrap;gap:6px">';
  d.terminy.forEach(function(t){
    h+='<span style="display:inline-flex;align-items:baseline;gap:6px;border:1px solid #e8edf5;border-radius:20px;padding:3px 11px;background:#fbfcfe">'
      +'<b style="font-size:11px;color:#1a2b4a">'+t.data+'</b>'
      +(t.opis?'<span style="font-size:10px;color:#8a93a5">'+esc(t.opis)+'</span>':'')
      +'<span style="font-size:9px;color:#b0b8c8">'+esc(t.sezon)+'</span></span>';
  });
  h+='</div>';
  return h;
}

function wzrostHistoria(pl){
  var k=DB[TEAM_KI]; if(!k) return [];
  var pid=pl.person_id||null;
  var im=(pl.imie||'').toLowerCase().trim(), nz=(pl.nazwisko||'').toLowerCase().trim();
  var wynik=[];
  Object.keys(k.sezony||{}).forEach(function(s){
    (PLAYERS[String(TEAM_KI)][s]||{}) && Object.keys(PLAYERS[String(TEAM_KI)][s]||{}).forEach(function(d){
      (PLAYERS[String(TEAM_KI)][s][d]||[]).forEach(function(p){
        var ten=(pid&&p.person_id===pid) ||
                (!pid&&(p.imie||'').toLowerCase().trim()===im&&(p.nazwisko||'').toLowerCase().trim()===nz);
        if(ten && (p.wzrost||0)>0) wynik.push({sezon:s, druzyna:d, wzrost:p.wzrost});
      });
    });
  });
  wynik.sort(function(a,b){return a.sezon<b.sezon?-1:(a.sezon>b.sezon?1:0);});
  return wynik;
}

function showPlayer(pi,keepEdit){
  EDIT_PI=pi;
  if(!keepEdit){SP_EDIT=false;SP_DIRTY=false;}
  var ED=SP_EDIT;
  var pl=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D][pi];
  if(!pl)return;
  var k=DB[TEAM_KI];
  var ext=k?(k.ext||false):false;
  var color=ext?'#085041':'#1a2b4a';
  var logo='';if(k){var _safe=(k.name||'').replace(/ /g,'_').replace(/\//g,'-');logo=LOGOS[_safe]||'';}
  var akt=pl.aktywny!==false;
  var dob=pl.data_urodzenia||pl.dob||'';
  var lic=pl.numer_licencji_pzk||pl.lic||'';
  var wzrost=pl.wzrost||'';
  var sc=pl._scout||{};

  // Wiek
  var age='';
  if(dob){var _d=new Date(dob);if(!isNaN(_d.getTime())){var _t=new Date();var _a=_t.getFullYear()-_d.getFullYear();if(_t.getMonth()<_d.getMonth()||(_t.getMonth()===_d.getMonth()&&_t.getDate()<_d.getDate()))_a--;if(_a>0&&_a<35)age=_a+' lat';}}

  var archetypes=['Shot Creator','Floor General','High Usage Guard','3&D Wing','Versatile Forward','Modern Big','Energy Defender'];
  var spItems=[{k:'talent',l:'TALENT'},{k:'pracowitosci',l:'PRACOWITOŚĆ'},{k:'iq',l:'IQ'},{k:'charakter',l:'CHARAKTER'},{k:'potencjal',l:'POTENCJAŁ'}];
  var soItems=[
    {k:'archetyp',l:'ARCHETYP',d:'Główny archetyp i kierunek rozwoju'},
    {k:'off_role',l:'OFFENSIVE ROLE',d:'Rola w ataku, tworzenie przewag'},
    {k:'def_role',l:'DEFENSIVE ROLE',d:'Rola obronna, wpływ na obronę'},
    {k:'scalability',l:'SCALABILITY',d:'Skalowanie na wyższy poziom'},
    {k:'playoff',l:'PLAYOFF VIABILITY',d:'Skuteczność w meczach wysokiego ciśnienia'},
    {k:'versatility',l:'VERSATILITY',d:'Ile aspektów gry wnosi wartość?'},
    {k:'decision',l:'DECISION MAKING',d:'Podejmowanie decyzji pod presją'},
    {k:'winning',l:'WINNING IMPACT',d:'Wpływ na wygrywanie poza statystykami'},
    {k:'lineup',l:'LINEUP FIT',d:'Optymalne ustawienie i koledzy'}
  ];

  var h='<div style="display:flex;flex-direction:column;height:100%">';

  // HEADER
  h+='<div style="background:'+color+';position:relative;overflow:hidden;padding:12px 14px;flex-shrink:0">';
  if(logo)h+='<img src="'+logo+'" style="position:absolute;right:-5px;bottom:-5px;height:80px;object-fit:contain;opacity:.07;pointer-events:none">';
  h+='<div style="display:flex;align-items:center;gap:10px">';
  h+='<div style="width:46px;height:46px;border-radius:10px;background:rgba(255,255,255,.18);display:flex;align-items:center;justify-content:center;font-size:19px;font-weight:800;color:#fff;flex-shrink:0">'+esc(String(pl.num||'?'))+'</div>';
  h+='<div style="flex:1;min-width:0">';
  h+='<div style="font-size:16px;font-weight:700;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+esc((pl.imie||'')+' '+(pl.nazwisko||''))+'</div>';
  h+='<div style="font-size:10px;color:rgba(255,255,255,.65);margin-top:2px">#'+esc(String(pl.num||''))+' &middot; '+esc(pl.poz||'—')+' &middot; '+esc(k?k.name:'')+' &middot; '+esc(TEAM_S)+'</div>';
  h+='<div style="display:flex;gap:4px;margin-top:6px;flex-wrap:wrap">';
  if(sc.archetyp)h+='<span style="background:rgba(255,255,255,.22);color:#fff;font-size:9px;font-weight:700;padding:2px 7px;border-radius:4px;letter-spacing:.04em">'+esc(sc.archetyp.toUpperCase())+'</span>';
  if(pl.poz)h+='<span style="background:rgba(255,255,255,.14);color:rgba(255,255,255,.85);font-size:9px;font-weight:600;padding:2px 7px;border-radius:4px">'+esc(pl.poz)+'</span>';
  if(age)h+='<span style="background:rgba(255,255,255,.14);color:rgba(255,255,255,.85);font-size:9px;font-weight:600;padding:2px 7px;border-radius:4px">'+esc(age)+'</span>';
  if(sc.projekcja)h+='<span style="background:rgba(239,159,39,.35);color:#FFD788;font-size:9px;font-weight:700;padding:2px 7px;border-radius:4px">Proj: '+esc(sc.projekcja)+'</span>';
  h+='</div></div>';
  h+='<div style="display:flex;gap:5px;flex-shrink:0">';
  h+='<button id="sp-edit" title="'+(ED?'Zapisz zmiany':'Odblokuj edycję danych')+'" style="background:'+(ED?'#EF9F27':'rgba(255,255,255,.22)')+';border:1px solid '+(ED?'#EF9F27':'rgba(255,255,255,.3)')+';cursor:pointer;color:#fff;font-size:10px;font-weight:600;padding:5px 10px;border-radius:7px;white-space:nowrap">'+(ED?'&#10003; Zapisz':'&#9998; Edytuj')+'</button>';
  if(ED) h+='<button id="sp-cancel" title="Odrzuć zmiany" style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);cursor:pointer;color:#fff;font-size:10px;padding:5px 10px;border-radius:7px;white-space:nowrap">Anuluj</button>';
  h+='<button id="sp-back" style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);cursor:pointer;color:#fff;font-size:10px;padding:5px 10px;border-radius:7px;white-space:nowrap">&#8592; Drużyna</button>';
  h+='</div>';
  h+='</div></div>';

  // TABS
  h+='<div style="display:flex;border-bottom:1px solid #f0f0f0;background:#fff;flex-shrink:0;overflow-x:auto">';
  [['bio','Bio'],['profil','Profil'],['rozwoj','Rozwój'],['trener','Ocena Trenera'],['stats','Statystyki'],['plan','Plan Rozwoju']].forEach(function(t,i){
    var a=i===0;
    h+='<button class="sp-tab" data-tab="'+t[0]+'" style="background:none;border:none;border-bottom:2px solid '+(a?color:'transparent')+';padding:9px 13px;font-size:12px;font-weight:'+(a?'600':'400')+';color:'+(a?color:'#999')+';cursor:pointer;white-space:nowrap">'+t[1]+'</button>';
  });
  h+='</div>';

  // TAB CONTENT
  h+='<div style="flex:1;overflow-y:auto">';

  // BIO TAB — dane osobowe (globalne) i drużynowe
  var _fmtDob=function(d){
    if(!d) return '';
    var m=String(d).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? (m[3]+'.'+m[2]+'.'+m[1]) : String(d);
  };
  var _lbl=function(t){return '<div style="font-size:9px;font-weight:700;letter-spacing:.06em;color:#b0b8c8;text-transform:uppercase;margin-bottom:3px">'+t+'</div>';};
  var _inpS='width:100%;padding:6px 9px;border:1px solid #cfd8e8;border-radius:7px;font-size:12.5px;box-sizing:border-box;font-family:inherit;color:#1a2b4a;background:#fff';

  // W trybie podglądu — wartość; w trybie edycji — pole formularza.
  var _bioRow=function(label,val,id,suffix){
    return '<div style="min-width:0">'+_lbl(label)
      +'<div'+(id?' id="'+id+'"':'')+' style="font-size:13px;font-weight:600;color:#1a2b4a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'
      +(val?(esc(String(val))+(suffix||'')):'<span style="color:#c8cfda;font-weight:400">—</span>')
      +'</div></div>';
  };
  var _bioInp=function(label,val,id,type,ph){
    return '<div style="min-width:0">'+_lbl(label)
      +'<input id="'+id+'" type="'+(type||'text')+'" value="'+esc(String(val||''))+'"'
      +(ph?' placeholder="'+esc(ph)+'"':'')+' style="'+_inpS+'"></div>';
  };

  h+='<div id="spt-bio" style="padding:13px">';

  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:9px">Dane osobowe (globalne)</div>';
  h+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:11px 14px;margin-bottom:11px">';
  if(ED){
    h+=_bioInp('Imię *',pl.imie,'be-imie');
    h+=_bioInp('Nazwisko *',pl.nazwisko,'be-nazw');
    h+=_bioInp('Data urodzenia',dob,'be-dob','date');
    h+=_bioInp('Nr licencji PZKosz',lic,'be-lic','text','np. 99483');
    h+=_bioInp('Pseudonim','','be-psd','text','opcjonalny');
    h+=_bioInp('Wzrost w sezonie '+TEAM_S+' (cm)',(pl.wzrost||''),'be-wzrost','number','np. 185');
  } else {
    h+=_bioRow('Imię',pl.imie);
    h+=_bioRow('Nazwisko',pl.nazwisko);
    h+=_bioRow('Data urodzenia',_fmtDob(dob),'bio-v-dob');
    h+=_bioRow('Nr licencji PZKosz',lic,'bio-v-lic');
    h+=_bioRow('Pseudonim','','bio-v-psd');
    var _hw=wzrostHistoria(pl);
    var _biez=(pl.wzrost||0);
    var _delta='';
    if(_biez>0&&_hw.length>1){
      var _prev=null;
      for(var _i=0;_i<_hw.length;_i++){ if(_hw[_i].sezon<TEAM_S) _prev=_hw[_i]; }
      if(_prev){
        var _d=_biez-_prev.wzrost;
        var _kol=_d>0?'#1a6b3c':(_d<0?'#8b1a1a':'#9aa5b5');
        _delta=' <span style="font-size:11px;font-weight:700;color:'+_kol+'">'
              +(_d>0?'+':'')+_d+'</span>'
              +' <span style="font-size:10px;font-weight:400;color:#9aa5b5">od '+esc(_prev.sezon)+'</span>';
      }
    }
    h+=_bioRow('Wzrost ('+esc(TEAM_S)+')',_biez||'',  'bio-v-wzrost',
               ' <span style="font-size:11px;font-weight:400;color:#9aa5b5">cm</span>'+_delta);
  }
  h+='</div>';
  if(_hw.length){
    var _chipy='';
    _hw.forEach(function(x,ix){
      var _akt=x.sezon===TEAM_S;
      var _r=ix>0?(x.wzrost-_hw[ix-1].wzrost):null;
      _chipy+='<span style="display:inline-flex;align-items:baseline;gap:5px;padding:3px 10px;border-radius:20px;'
        +'border:1px solid '+(_akt?color:'#e8edf5')+';background:'+(_akt?'#f0f4ff':'#fff')+'">'
        +'<span style="font-size:9px;color:#9aa5b5">'+esc(x.sezon)+'</span>'
        +'<span style="font-size:12px;font-weight:700;color:#1a2b4a">'+x.wzrost+'</span>'
        +(_r!==null&&_r!==0?'<span style="font-size:10px;font-weight:700;color:'+(_r>0?'#1a6b3c':'#8b1a1a')+'">'+(_r>0?'+':'')+_r+'</span>':'')
        +'</span>';
    });
    h+='<div style="margin-bottom:13px">'
      +'<div style="font-size:9px;font-weight:700;letter-spacing:.06em;color:#b0b8c8;text-transform:uppercase;margin-bottom:6px">Wzrost w kolejnych sezonach</div>'
      +'<div style="display:flex;flex-wrap:wrap;gap:6px">'+_chipy+'</div></div>';
  }
  h+='<div id="bio-person" style="display:none;font-size:11px;line-height:1.45;padding:9px 11px;border-radius:8px;margin-bottom:15px"></div>';

  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:9px">Dane drużynowe</div>';
  h+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:11px 14px;margin-bottom:6px">';
  if(ED){
    h+=_bioInp('Numer',pl.num||'','be-num','number','np. 5');
    h+=_bioInp('Pozycja',pl.poz,'be-poz','text','np. Rzucający');
    h+='<div style="min-width:0">'+_lbl('Status')
      +'<select id="be-akt" style="'+_inpS+'">'
      +'<option value="1"'+(akt?' selected':'')+'>Aktywny</option>'
      +'<option value="0"'+(akt?'':' selected')+'>Nieaktywny</option>'
      +'</select></div>';
  } else {
    h+=_bioRow('Numer',pl.num?('#'+pl.num):'');
    h+=_bioRow('Pozycja',pl.poz);
    h+='<div style="min-width:0">'+_lbl('Status')
      +(akt
        ? '<span style="background:#e8f5ee;color:#1a6b3c;font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px">Aktywny</span>'
        : '<span style="background:#f5f5f5;color:#9ca3af;font-size:11px;padding:3px 10px;border-radius:20px">Nieaktywny</span>')
      +'</div>';
  }
  h+=_bioRow('Drużyna',TEAM_D+' · '+TEAM_S);
  h+='</div>';
  h+='<div id="bio-msg" style="display:none;font-size:11px;line-height:1.45;padding:9px 11px;border-radius:8px;margin-top:11px"></div>';

  h+='<div style="display:flex;justify-content:flex-end;border-top:1px solid #f0f4ff;margin-top:14px;padding-top:11px">';
  h+='<button id="sp-del" style="padding:7px 13px;background:none;border:1px solid #f0d5d5;border-radius:8px;cursor:pointer;font-size:11.5px;color:#C0392B">&#128465; Usuń zawodnika</button>';
  h+='</div>';

  h+='</div>'; // end bio

  // PROFIL TAB — ocena scoutingowa (dotychczasowa zawartość Bio)
  h+='<div id="spt-profil" style="display:none;padding:13px">';
  if(!ED) h+='<div style="background:#f4f7fb;border:1px solid #e2e8f2;color:#6b7c95;font-size:11px;padding:8px 11px;border-radius:8px;margin-bottom:12px">Podgląd oceny. Kliknij <b>&#9998; Edytuj</b> w nagłówku, aby zmieniać.</div>';
  h+='<div'+(ED?'':' style="pointer-events:none;opacity:.85" inert')+'>';

  // Archetyp
  h+='<div style="margin-bottom:13px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Archetyp</div>';
  h+='<div style="display:flex;flex-wrap:wrap;gap:5px">';
  archetypes.forEach(function(a){
    var sel=(sc.archetyp||'')===a;
    h+='<button class="arch-pill" data-arch="'+esc(a)+'" style="padding:5px 11px;border-radius:20px;border:1px solid '+(sel?color:'#e5e7eb')+';background:'+(sel?color:'#fff')+';color:'+(sel?'#fff':'#555')+';font-size:11px;cursor:pointer;font-weight:'+(sel?'600':'400')+'">'+esc(a)+'</button>';
  });
  h+='</div></div>';

  // Szybki podgląd
  h+='<div style="margin-bottom:13px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Szybki podgląd</div>';
  h+='<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:6px">';
  spItems.forEach(function(item){
    var val=(sc.sp||{})[item.k]||0;
    h+='<div style="background:#f8f9fc;border-radius:8px;padding:8px 5px;text-align:center">';
    h+='<div style="font-size:7px;font-weight:700;color:#bbb;margin-bottom:5px;letter-spacing:.04em;line-height:1.2">'+esc(item.l)+'</div>';
    h+='<div class="star-grp" data-prefix="sp" data-key="'+item.k+'" style="display:flex;justify-content:center;gap:1px">';
    for(var i=1;i<=5;i++){h+='<span class="star-b" data-prefix="sp" data-key="'+item.k+'" data-val="'+i+'" style="cursor:pointer;font-size:12px;color:'+(i<=val?'#EF9F27':'#ddd')+'">&#9733;</span>';}
    h+='</div></div>';
  });
  h+='</div></div>';

  // Opis
  h+='<div style="margin-bottom:13px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Opis zawodnika</div>';
  h+='<textarea id="sp-opis" style="width:100%;min-height:56px;padding:8px;border:1px solid #e8edf5;border-radius:8px;font-size:12px;resize:vertical;box-sizing:border-box;color:#333;line-height:1.4;font-family:inherit" placeholder="Brak opisu...">'+esc(sc.opis||'')+'</textarea>';
  h+='</div>';

  // Ocena Scoutingowa
  h+='<div style="margin-bottom:13px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Ocena Scoutingowa &mdash; 9 Sekcji</div>';
  h+='<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:7px">';
  soItems.forEach(function(item,idx){
    var val=(sc.so||{})[item.k]||0;
    h+='<div style="background:#fff;border:1px solid #e8edf5;border-radius:9px;padding:9px">';
    h+='<div style="display:flex;justify-content:space-between;margin-bottom:4px">';
    h+='<div style="font-size:8px;font-weight:700;color:#1a2b4a;letter-spacing:.03em;line-height:1.3">'+esc(item.l)+'</div>';
    h+='<div style="font-size:9px;color:#ddd;font-weight:700">'+(idx+1)+'</div>';
    h+='</div>';
    h+='<div class="star-grp" data-prefix="so" data-key="'+item.k+'" style="display:flex;gap:1px;margin-bottom:4px">';
    for(var i=1;i<=5;i++){h+='<span class="star-b" data-prefix="so" data-key="'+item.k+'" data-val="'+i+'" style="cursor:pointer;font-size:12px;color:'+(i<=val?'#EF9F27':'#ddd')+'">&#9733;</span>';}
    h+='</div>';
    h+='<div style="font-size:8px;color:#aaa;line-height:1.3">'+esc(item.d)+'</div>';
    h+='</div>';
  });
  h+='</div></div>';

  // Czy pomaga wygrywać
  h+='<div style="margin-bottom:13px">';
  h+='<div style="font-size:11px;font-weight:600;color:#374151;margin-bottom:7px">Czy jego gra pomaga wygrywać?</div>';
  h+='<div style="display:flex;gap:5px;flex-wrap:wrap">';
  [['tak','TAK'],['raczej_tak','RACZEJ TAK'],['raczej_nie','RACZEJ NIE'],['nie','NIE']].forEach(function(o){
    var sel=(sc.pomaga||'')===o[0];
    h+='<button class="pw-btn" data-val="'+o[0]+'" style="padding:6px 11px;border-radius:7px;border:1px solid '+(sel?color:'#e5e7eb')+';background:'+(sel?color:'#fff')+';color:'+(sel?'#fff':'#555')+';font-size:11px;font-weight:'+(sel?'600':'400')+';cursor:pointer">'+o[1]+'</button>';
  });
  h+='</div></div>';

  // Projekcja + Pewność
  h+='<div style="display:flex;gap:14px;margin-bottom:14px;flex-wrap:wrap">';
  h+='<div style="flex:1;min-width:150px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Projekcja</div>';
  h+='<div style="display:flex;gap:5px">';
  ['A','B','C','D','F'].forEach(function(p){
    var sel=(sc.projekcja||'')===p;
    h+='<button class="proj-btn" data-val="'+p+'" style="width:32px;height:32px;border-radius:7px;border:1px solid '+(sel?color:'#e5e7eb')+';background:'+(sel?color:'#fff')+';color:'+(sel?'#fff':'#555')+';font-size:13px;font-weight:'+(sel?'700':'400')+';cursor:pointer">'+p+'</button>';
  });
  h+='</div></div>';
  h+='<div style="flex:1;min-width:100px">';
  h+='<div style="font-size:9px;font-weight:700;letter-spacing:.07em;color:#b0b8c8;text-transform:uppercase;margin-bottom:7px">Pewność</div>';
  h+='<div class="star-grp" data-prefix="pew" data-key="pewnosc" style="display:flex;gap:3px">';
  var pewVal=sc.pewnosc||0;
  for(var _i=1;_i<=5;_i++){h+='<span class="star-b" data-prefix="pew" data-key="pewnosc" data-val="'+_i+'" style="cursor:pointer;font-size:18px;color:'+(_i<=pewVal?'#EF9F27':'#ddd')+'">&#9733;</span>';}
  h+='</div></div>';
  h+='</div>';

  h+='</div>'; // koniec obszaru kontrolek
  h+='</div>'; // end profil

  // Inne taby (stub)
  h+='<div id="spt-rozwoj" style="display:none;padding:13px"><div style="color:#bbb;font-size:12px;text-align:center;padding:26px">Wczytuję pomiary…</div></div>';
  h+='<div id="spt-trener" style="display:none;padding:14px"><div style="background:#f8f9fc;border-radius:10px;padding:20px;text-align:center;color:#bbb;font-size:12px"><div style="font-size:22px;margin-bottom:8px">&#128203;</div>Ocena trenera &mdash; wkrótce</div></div>';
  h+='<div id="spt-stats" style="display:none;padding:14px"><div style="background:#f8f9fc;border-radius:10px;padding:20px;text-align:center;color:#bbb;font-size:12px"><div style="font-size:22px;margin-bottom:8px">&#128202;</div>Statystyki z meczów &mdash; wkrótce</div></div>';
  h+='<div id="spt-plan" style="display:none;padding:14px"><div style="background:#f8f9fc;border-radius:10px;padding:20px;text-align:center;color:#bbb;font-size:12px"><div style="font-size:22px;margin-bottom:8px">&#127919;</div>Plan Rozwoju &mdash; wkrótce</div></div>';

  h+='</div>'; // end tab content
  h+='</div>'; // end main

  // sbar breadcrumb
  var _kn=esc(k?k.name:'');
  var _pn=esc((pl.imie||'')+' '+(pl.nazwisko||''));
  var sbar='<div style="display:flex;align-items:center;width:100%;gap:5px;overflow:hidden">'
    +'<span id="bc-home" style="font-size:11px;cursor:pointer;color:#EF9F27;flex-shrink:0">Drużyny</span>'
    +'<span style="font-size:11px;color:#ddd;flex-shrink:0">&rsaquo;</span>'
    +'<span id="bc-team" style="font-size:11px;cursor:pointer;color:#888;flex-shrink:0;white-space:nowrap">'+_kn+' &middot; '+esc(TEAM_D)+'</span>'
    +'<span style="font-size:11px;color:#ddd;flex-shrink:0">&rsaquo;</span>'
    +'<span style="font-size:11px;color:#444;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+_pn+'</span>'
    +'</div>';
  panel(h,"",EMBED?"":sbar);

  // Breadcrumb + back
  var _bch=document.getElementById("bc-home");if(_bch)_bch.addEventListener("click",idle);
  var _bct=document.getElementById("bc-team");if(_bct)_bct.addEventListener("click",renderTeam);
  document.getElementById("sp-back").addEventListener('click',renderTeam);

  // Tab switcher
  document.querySelectorAll('.sp-tab').forEach(function(btn){
    btn.addEventListener('click',function(){
      var tab=this.dataset.tab;
      document.querySelectorAll('.sp-tab').forEach(function(b){b.style.borderBottomColor='transparent';b.style.color='#999';b.style.fontWeight='400';});
      this.style.borderBottomColor=color;this.style.color=color;this.style.fontWeight='600';
      ['bio','profil','rozwoj','trener','stats','plan'].forEach(function(t){
        var el=document.getElementById('spt-'+t);
        if(el)el.style.display=(t===tab)?'block':'none';
      });
    });
  });

  // Kontrolki oceny — obsługa tylko w trybie edycji
  // Archetyp pills
  if(ED) document.querySelectorAll('.arch-pill').forEach(function(btn){
    btn.addEventListener('click',function(){
      var a=this.dataset.arch;
      if(!pl._scout)pl._scout={};
      pl._scout.archetyp=(pl._scout.archetyp===a?'':a);
      SP_DIRTY=true;showPlayer(pi,true);
    });
  });

  // Star buttons (without full re-render)
  if(ED) document.querySelectorAll('.star-b').forEach(function(btn){
    btn.addEventListener('click',function(){
      var prefix=this.dataset.prefix,key=this.dataset.key,val=parseInt(this.dataset.val);
      if(!pl._scout)pl._scout={};
      var sc2=pl._scout;
      if(prefix==='sp'){if(!sc2.sp)sc2.sp={};sc2.sp[key]=(sc2.sp[key]===val?0:val);}
      else if(prefix==='so'){if(!sc2.so)sc2.so={};sc2.so[key]=(sc2.so[key]===val?0:val);}
      else if(prefix==='pew'){sc2.pewnosc=(sc2.pewnosc===val?0:val);}
      SP_DIRTY=true;
      var grp=this.closest?this.closest('.star-grp'):this.parentElement;
      var nv=prefix==='sp'?(sc2.sp[key]||0):(prefix==='so'?(sc2.so[key]||0):(sc2.pewnosc||0));
      if(grp)grp.querySelectorAll('.star-b').forEach(function(s){s.style.color=parseInt(s.dataset.val)<=nv?'#EF9F27':'#ddd';});
    });
  });

  // Pomaga wygrywać
  if(ED) document.querySelectorAll('.pw-btn').forEach(function(btn){
    btn.addEventListener('click',function(){
      var val=this.dataset.val;
      if(!pl._scout)pl._scout={};
      pl._scout.pomaga=(pl._scout.pomaga===val?'':val);
      SP_DIRTY=true;
      document.querySelectorAll('.pw-btn').forEach(function(b){
        var sel=(pl._scout.pomaga||'')===b.dataset.val;
        b.style.borderColor=sel?color:'#e5e7eb';b.style.background=sel?color:'#fff';
        b.style.color=sel?'#fff':'#555';b.style.fontWeight=sel?'600':'400';
      });
    });
  });

  // Projekcja
  if(ED) document.querySelectorAll('.proj-btn').forEach(function(btn){
    btn.addEventListener('click',function(){
      var val=this.dataset.val;
      if(!pl._scout)pl._scout={};
      pl._scout.projekcja=(pl._scout.projekcja===val?'':val);
      SP_DIRTY=true;
      document.querySelectorAll('.proj-btn').forEach(function(b){
        var sel=(pl._scout.projekcja||'')===b.dataset.val;
        b.style.borderColor=sel?color:'#e5e7eb';b.style.background=sel?color:'#fff';
        b.style.color=sel?'#fff':'#555';b.style.fontWeight=sel?'700':'400';
      });
    });
  });

  // Opis auto-save
  var opisEl=document.getElementById("sp-opis");
  if(opisEl&&!ED) opisEl.readOnly=true;
  if(opisEl&&ED)opisEl.addEventListener("change",function(){if(!pl._scout)pl._scout={};pl._scout.opis=this.value;SP_DIRTY=true;});

  // Edytuj / Zapisz — edycja odbywa się na miejscu, bez osobnego okna
  document.getElementById("sp-edit").addEventListener('click',function(){
    if(!SP_EDIT){ SP_EDIT=true; showPlayer(pi,true); return; }

    var v=function(id){var e=document.getElementById(id);return e?e.value:'';};
    var im=(v('be-imie')||'').trim(), nz=(v('be-nazw')||'').trim();
    var msg=document.getElementById('bio-msg');
    var say=function(txt,bg,fg){ if(msg){msg.textContent=txt;msg.style.background=bg;msg.style.color=fg;msg.style.display='block';} };
    if(!im||!nz){
      var e1=document.getElementById('be-imie'); if(e1&&!im) e1.style.borderColor='#E24B4A';
      var e2=document.getElementById('be-nazw'); if(e2&&!nz) e2.style.borderColor='#E24B4A';
      say('Imię i nazwisko są wymagane.','#ffe1e1','#a30000');
      return;
    }

    // 1) dane drużynowe — lokalnie
    pl.imie=im; pl.nazwisko=nz;
    pl.num=parseInt(v('be-num'))||0;
    pl.poz=v('be-poz')||'';
    pl.aktywny=(v('be-akt')!=='0');
    pl.data_urodzenia=v('be-dob')||'';
    pl.numer_licencji_pzk=(v('be-lic')||'').trim();
    pl.wzrost=parseInt(v('be-wzrost'))||0;   // pomiar tego sezonu

    // 2) dane osobowe — wspólne dla wszystkich drużyn tego zawodnika
    var btn=this, lbl=btn.innerHTML;
    btn.disabled=true; btn.innerHTML='&#8987; Zapisuję…';
    fetch('/api/person/save',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({imie:im,nazwisko:nz,
        data_urodzenia:v('be-dob'),
        numer_licencji_pzk:(v('be-lic')||'').trim(),
        pseudonim:(v('be-psd')||'').trim(),
        aktywny:(v('be-akt')!=='0')})})
      .then(function(r){return r.json();})
      .then(function(d){
        if(!d||!d.ok) throw new Error(d&&d.error?d.error:'Błąd zapisu danych osobowych');
        saveDB();                       // jeden zapis dla wszystkiego
        SP_EDIT=false; SP_DIRTY=false;
        showPlayer(pi,true);
      })
      .catch(function(err){
        btn.disabled=false; btn.innerHTML=lbl;
        say('Nie zapisano: '+(err.message||err),'#ffe1e1','#a30000');
      });
  });

  // Anuluj — porzuć zmiany i wczytaj dane od nowa
  var _spc=document.getElementById("sp-cancel");
  if(_spc) _spc.addEventListener('click',function(){
    SP_EDIT=false;
    if(SP_DIRTY){ SP_DIRTY=false; renderTeam(); return; }   // przeładuj z zapisanego stanu
    showPlayer(pi,true);
  });

  // Rozwój — historia pomiarów tego zawodnika
  (function(){
    var el=document.getElementById('spt-rozwoj');
    if(!el) return;
    if(!pl.person_id){
      el.innerHTML='<div style="color:#9aa5b5;font-size:12px;text-align:center;padding:26px">'
        +'Zawodnik nie ma jeszcze kartoteki osoby — pomiary pojawią się po pierwszym zapisie danych.</div>';
      return;
    }
    var tid=null;
    try{ var mm=location.search.match(/team_id=(\d+)/); if(mm) tid=mm[1]; }catch(e){}
    fetch('/klub/api/pomiar/osoba/'+pl.person_id+(tid?('?team_id='+tid):''))
      .then(function(r){return r.json();})
      .then(function(d){ el.innerHTML=rysujRozwoj(d,color); })
      .catch(function(e){ el.innerHTML='<div style="color:#8b1a1a;font-size:12px;padding:20px">Nie udało się wczytać pomiarów: '+e+'</div>'; });
  })();

  // Bio — dociągnięcie danych globalnych (pseudonim, potwierdzenie profilu osoby)
  (function(){
    var _im=(pl.imie||'').trim(), _nz=(pl.nazwisko||'').trim();
    var box=document.getElementById('bio-person');
    if(!_im||!_nz||!box) return;
    fetch('/api/person/lookup?imie='+encodeURIComponent(_im)+'&nazwisko='+encodeURIComponent(_nz))
      .then(function(r){return r.json();})
      .then(function(d){
        if(!d||!d.found){
          box.textContent='ℹ Brak profilu globalnego — powstanie przy pierwszym zapisie danych.';
          box.style.background='#fff8e1'; box.style.color='#7a5200'; box.style.display='block';
          return;
        }
        var p=d.person;
        // uzupełnia podgląd albo pole formularza — zależnie od trybu
        var set=function(viewId,inpId,v,suf,raw){
          if(!v) return;
          var el=document.getElementById(viewId);
          if(el){ el.innerHTML=esc(String(v))+(suf||''); return; }
          var ie=document.getElementById(inpId);
          if(ie&&!ie.value) ie.value=(raw!==undefined?raw:v);
        };
        set('bio-v-dob','be-dob',_fmtDob(p.data_urodzenia),'',p.data_urodzenia);
        set('bio-v-lic','be-lic',p.numer_licencji_pzk);
        set('bio-v-psd','be-psd',p.pseudonim);
        // wzrost NIE jest dociągany z kartoteki — to pomiar konkretnego sezonu
        var miss=[];
        if(!p.data_urodzenia) miss.push('data urodzenia');
        if(!p.numer_licencji_pzk) miss.push('nr licencji PZKosz');
        if(miss.length){
          box.textContent='⚠ Profil globalny istnieje (id='+p.id+', w '+(d.teams_count||'?')+' drużynach) — brakuje: '+miss.join(', ')+'.';
          box.style.background='#fff4d6'; box.style.color='#7a5200';
        } else {
          box.textContent='✓ Profil globalny w pełni uzupełniony (id='+p.id+') — dane wspólne dla wszystkich drużyn tego zawodnika.';
          box.style.background='#e1f5e1'; box.style.color='#2c662c';
        }
        box.style.display='block';
      })
      .catch(function(){});
  })();

  // Usuń
  document.getElementById("sp-del").addEventListener('click',function(){
    pendDelPl=pi;
    document.getElementById("ov-title").textContent="Usunąć zawodnika?";
    document.getElementById("ov-msg").textContent="Operacji nie można cofnąć.";
    document.getElementById("ov-ok").onclick=function(){PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D].splice(pendDelPl,1);closeOv();saveDB();renderTeam();};
    document.getElementById("ov-del").style.display="flex";
  });
}

function clearPersonFields(){
  ["pm-dob","pm-lic","pm-psd","pm-wzrost"].forEach(function(id){
    var el=document.getElementById(id);
    if(el){
      if(el.tagName==="TEXTAREA"||el.tagName==="INPUT") el.value="";
    }
  });
  var info=document.getElementById("pm-person-info");
  if(info){info.style.display="none";info.textContent="";}
}

function lookupPersonByName(){
  var im=document.getElementById("pm-imie").value.trim();
  var nz=document.getElementById("pm-nazw").value.trim();
  if(!im||!nz){clearPersonFields();return;}
  fetch("/api/person/lookup?imie="+encodeURIComponent(im)+"&nazwisko="+encodeURIComponent(nz))
    .then(function(r){return r.json();})
    .then(function(d){
      var info=document.getElementById("pm-person-info");
      if(d.found){
        var p=d.person;
        document.getElementById("pm-dob").value=p.data_urodzenia||"";
        document.getElementById("pm-lic").value=p.numer_licencji_pzk||"";
        document.getElementById("pm-psd").value=p.pseudonim||"";
        document.getElementById("pm-wzrost").value=p.wzrost||"";
        if(info){
          var hasDob = !!(p.data_urodzenia);
          var hasLic = !!(p.numer_licencji_pzk);
          var miss = [];
          if(!hasDob) miss.push("data urodzenia");
          if(!hasLic) miss.push("nr licencji PZKosz");
          if(miss.length){
            info.textContent="⚠ Profil globalny istnieje (id="+p.id+", w "+(d.teams_count||"?")+" drużynach) — BRAK: "+miss.join(", ")+". Uzupełnij teraz aby dane propagowały do wszystkich drużyn.";
            info.style.background="#fff4d6"; info.style.color="#7a5200";
          } else {
            info.textContent="✓ Profil globalny w pełni uzupełniony (id="+p.id+") — dane załadowane. Zmiany zapiszą się dla wszystkich drużyn tego zawodnika.";
            info.style.background="#e1f5e1"; info.style.color="#2c662c";
          }
          info.style.display="block";
        }
      } else {
        if(info){info.textContent="ℹ Nowa osoba — po zapisie utworzy się globalny profil. Możesz uzupełnić DOB/Licencję teraz lub później.";info.style.background="#fff8e1";info.style.color="#7a5200";info.style.display="block";}
      }
    })
    .catch(function(){});
}
document.getElementById("pm-imie").addEventListener("blur",lookupPersonByName);
document.getElementById("pm-nazw").addEventListener("blur",lookupPersonByName);

document.getElementById("pm-cancel").addEventListener("click",function(){document.getElementById("ov-player").style.display="none";});
document.getElementById("pm-save").addEventListener("click",function(){
  var im=document.getElementById("pm-imie").value.trim();
  var nz=document.getElementById("pm-nazw").value.trim();
  if(!im||!nz){document.getElementById("pm-imie").style.borderColor="#E24B4A";return;}
  var poz=document.getElementById("pm-poz").value;
  var num=parseInt(document.getElementById("pm-num").value)||0;
  var aktEl=document.getElementById("pm-aktywny");
  var aktywny=aktEl?aktEl.value!=="0":true;
  var _dobV=document.getElementById("pm-dob").value;
  var _licV=document.getElementById("pm-lic").value.trim();
  var _wzV=parseInt(document.getElementById("pm-wzrost").value)||0;
  var _existing=(EDIT_PI>=0&&PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D][EDIT_PI])||{};
  var _plRec={imie:im,nazwisko:nz,num:num,poz:poz,aktywny:aktywny,
              data_urodzenia:_dobV||"", numer_licencji_pzk:_licV||"", wzrost:_wzV||0};
  if(_existing._scout)_plRec._scout=_existing._scout;
  if(EDIT_PI>=0){
    PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D][EDIT_PI]=_plRec;
  } else {
    PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D].push(_plRec);
  }
  EDIT_PI=-1;
  var personPayload={
    imie:im, nazwisko:nz,
    data_urodzenia:document.getElementById("pm-dob").value,
    numer_licencji_pzk:document.getElementById("pm-lic").value.trim(),
    pseudonim:document.getElementById("pm-psd").value.trim(),
    wzrost:document.getElementById("pm-wzrost").value.trim(),
    aktywny:aktywny
  };
  var btn=document.getElementById("pm-save");
  var origLbl=btn.textContent;
  btn.disabled=true; btn.textContent="⏳ Zapisuję...";
  fetch("/api/person/save",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(personPayload)})
    .then(function(r){return r.json();})
    .then(function(d){
      if(!d||!d.ok){throw new Error(d&&d.error?d.error:"Błąd zapisu persons");}
      saveDB();
      var info=document.getElementById("pm-person-info");
      if(info){info.textContent="✓ Zapisano globalnie (person_id="+d.person_id+"). Dane będą widoczne we wszystkich drużynach tego zawodnika.";info.style.background="#e1f5e1";info.style.color="#2c662c";info.style.display="block";}
      setTimeout(function(){
        document.getElementById("ov-player").style.display="none";
        btn.disabled=false; btn.textContent=origLbl;
        renderTeam();
      }, 700);
    })
    .catch(function(err){
      btn.disabled=false; btn.textContent=origLbl;
      var info=document.getElementById("pm-person-info");
      if(info){info.textContent="❌ BŁĄD zapisu: "+(err.message||err)+". Sprawdź konsolę (F12).";info.style.background="#ffe1e1";info.style.color="#a30000";info.style.display="block";}
      console.error("[pm-save] error:",err);
    });
  return;
});

function _normalizeDob(v){
  if(!v) return "";
  v=String(v).trim();
  if(!v) return "";
  if(/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  var m=v.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if(m) return m[3]+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");
  var n=parseFloat(v);
  if(!isNaN(n) && n>20000 && n<60000){
    var d=new Date(Math.round((n-25569)*86400000));
    return d.toISOString().substring(0,10);
  }
  return "";
}

var IMPORT_PENDING=null;
// ── Krok potwierdzenia: wzrost do poprawienia przed zapisem ─────────────────
function _importPodglad(strategy){
  if(!IMPORT_PENDING) return;
  var np=IMPORT_PENDING.newPlayers||[];
  var current=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[];
  var byKey={};
  current.forEach(function(p){
    byKey[(p.nazwisko||'').toLowerCase().trim()+'|'+(p.imie||'').toLowerCase().trim()]=p;
  });
  var h='<div style="font-size:15px;font-weight:600;color:#1a2b4a;margin-bottom:4px">Sprawdź wzrost przed zapisem</div>'
   +'<div style="font-size:11.5px;color:#666;margin-bottom:12px;line-height:1.5">'
   +'Pomiar zapisze się dla sezonu <b>'+esc(TEAM_S)+'</b>. Puste pole = brak pomiaru, '
   +'poprzednia wartość zostanie zachowana.</div>'
   +'<div style="max-height:330px;overflow-y:auto;border:1px solid #eef1f7;border-radius:9px">'
   +'<table style="width:100%;border-collapse:collapse;font-size:12px">';
  np.forEach(function(p,i){
    var k=(p.nazwisko||'').toLowerCase().trim()+'|'+(p.imie||'').toLowerCase().trim();
    var ist=byKey[k];
    var stan=ist
      ? '<span style="font-size:9px;font-weight:700;color:#1a6b3c;background:#e8f5ee;padding:1px 7px;border-radius:20px">w składzie</span>'
      : '<span style="font-size:9px;font-weight:700;color:#7a5200;background:#fff8e1;padding:1px 7px;border-radius:20px">nowy</span>';
    var poprz=(ist&&ist.wzrost)?ist.wzrost:0;
    h+='<tr style="border-bottom:1px solid #f4f6fa">'
      +'<td style="padding:6px 10px;color:#1a2b4a;font-weight:600;white-space:nowrap">'+esc(p.nazwisko+' '+p.imie)+'</td>'
      +'<td style="padding:6px 6px">'+stan+'</td>'
      +'<td style="padding:6px 6px;color:#9aa5b5;font-size:10.5px;white-space:nowrap">'
        +(poprz?('było '+poprz):'')+'</td>'
      +'<td style="padding:6px 10px;text-align:right;white-space:nowrap">'
      +'<input type="number" class="imp-wz" data-i="'+i+'" value="'+(p._wzrost||'')+'" min="0" max="260" '
      +'style="width:66px;border:1px solid #dde3ee;border-radius:6px;padding:4px 7px;font-size:12px;text-align:right"> cm</td>'
      +'</tr>';
  });
  h+='</table></div>'
   +'<div style="display:flex;gap:8px;margin-top:14px">'
   +'<button id="imp-wroc" style="flex:1;background:none;border:1px solid #ddd;color:#888;padding:9px;border-radius:8px;cursor:pointer;font-size:12px">Wstecz</button>'
   +'<button id="imp-ok" style="flex:2;background:#EF9F27;color:#fff;border:none;padding:9px;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600">Zapisz import</button>'
   +'</div>';
  var box=document.querySelector('#ov-import > div');
  box.innerHTML=h;
  document.getElementById('imp-wroc').addEventListener('click',function(){
    document.getElementById('ov-import').style.display='none';
    _showImportDialog(np);
  });
  document.getElementById('imp-ok').addEventListener('click',function(){
    document.querySelectorAll('.imp-wz').forEach(function(inp){
      var i=parseInt(inp.dataset.i);
      np[i]._wzrost=parseInt(inp.value)||0;
    });
    _applyImport(strategy);
  });
}

function _applyImport(strategy){
  if(!IMPORT_PENDING) return;
  var newPlayers=IMPORT_PENDING.newPlayers||[];
  var current=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[];
  var added=0,updated=0,replaced=0;
  function _mkRec(p){
    return {imie:p.imie, nazwisko:p.nazwisko, num:p.num, poz:p.poz, aktywny:p.aktywny,
            data_urodzenia:p._dob||"", numer_licencji_pzk:p._lic||"", wzrost:p._wzrost||0};
  }
  if(strategy==="replace"){
    PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]=newPlayers.map(_mkRec);
    replaced=newPlayers.length;
  } else if(strategy==="merge"){
    var byKey={};
    current.forEach(function(p,i){
      var k=(p.nazwisko||"").toLowerCase().trim()+"|"+(p.imie||"").toLowerCase().trim();
      byKey[k]=i;
    });
    newPlayers.forEach(function(np){
      var k=(np.nazwisko||"").toLowerCase().trim()+"|"+(np.imie||"").toLowerCase().trim();
      if(byKey[k]!==undefined){
        var idx=byKey[k];
        if(np.num) current[idx].num=np.num;
        if(np.poz) current[idx].poz=np.poz;
        if(np._dob) current[idx].data_urodzenia=np._dob;
        if(np._lic) current[idx].numer_licencji_pzk=np._lic;
        if(np._wzrost) current[idx].wzrost=np._wzrost;   // 0 = brak pomiaru, zostaje poprzedni
        current[idx].aktywny=true;
        updated++;
      } else {
        current.push(_mkRec(np));
        added++;
      }
    });
    PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]=current;
  } else if(strategy==="append"){
    newPlayers.forEach(function(np){ current.push(_mkRec(np)); });
    PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]=current;
    added=newPlayers.length;
  }
  IMPORT_PENDING=null;
  document.getElementById("ov-import").style.display="none";
  var parts=[];
  if(replaced) parts.push("zastąpiono "+replaced);
  if(updated)  parts.push("zaktualizowano "+updated);
  if(added)    parts.push("dodano "+added);
  fetch("/druzyny/save",{method:"POST",headers:{"Content-Type":"application/json"},
         body:JSON.stringify({kluby:DB,players:PLAYERS})})
    .then(function(r){return r.json();})
    .then(function(d){
      if(d&&d.ok===false) throw new Error(d.error||"Błąd zapisu");
      alert("Import zakończony: "+parts.join(", ")+" zawodników.\n"
        +"Wzrost zapisany dla sezonu "+TEAM_S+". Data urodzenia i licencja trafiły do kartoteki globalnej "
        +"(tylko tam, gdzie były puste).");
      renderTeam();
    })
    .catch(function(err){
      alert("❌ Błąd zapisu: "+(err.message||err));
      console.error("[import] save error:",err);
    });
}
function _parseImportRows(rows,isCsv){
  var parsed=[];
  for(var i=1;i<rows.length;i++){
    var row=rows[i];
    if(!row) continue;
    if(isCsv && row.length<3) continue;
    var nz=String(row[1]||"").trim(),im=String(row[2]||"").trim();
    if(!im&&!nz)continue;
    if(nz.toLowerCase()==="nazwisko"||im.toLowerCase()==="imie"||im.toLowerCase()==="imię")continue;
    parsed.push({
      imie:im, nazwisko:nz, num:0, poz:String(row[6]||"").trim(), aktywny:true,
      _lic:String(row[3]||"").trim(),
      _dob:_normalizeDob(String(row[4]||"").trim()),
      _wzrost:parseInt(row[5])||0
    });
  }
  return parsed;
}
function _showImportDialog(parsed){
  if(!parsed.length){ alert("Brak danych w pliku."); return; }
  var current=PLAYERS[String(TEAM_KI)][TEAM_S][TEAM_D]||[];
  IMPORT_PENDING={newPlayers:parsed};
  if(current.length===0){
    _applyImport("replace");
    return;
  }
  document.getElementById("im-info").innerHTML=
    "W drużynie <b>"+esc(TEAM_D)+"</b> jest już <b>"+current.length+"</b> zawodników. "+
    "Z pliku wczytano <b>"+parsed.length+"</b> zawodników.<br><br>"+
    "<b>Wybierz strategię:</b>";
  document.getElementById("ov-import").style.display="flex";
}
function importF(file){
  var ext=file.name.split(".").pop().toLowerCase();
  if(ext==="csv"){
    var r=new FileReader();r.onload=function(e){
      var lines=e.target.result.split(/\r?\n/);
      var rows=lines.map(function(l){return l.split(",");});
      _showImportDialog(_parseImportRows(rows,true));
    };r.readAsText(file,"UTF-8");
  } else if(ext==="xlsx"){
    var r=new FileReader();r.onload=function(e){
      try{
        var wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});
        var rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:"",raw:false});
        _showImportDialog(_parseImportRows(rows,false));
      }catch(err){alert("Błąd: "+err.message);}
    };r.readAsArrayBuffer(file);
  } else {
    alert("Obsługiwane: .xlsx, .csv");
  }
}
document.querySelectorAll(".im-opt").forEach(function(b){
  b.addEventListener("click",function(){_importPodglad(this.dataset.strategy);});
});
var imCancelBtn=document.getElementById("im-cancel");
if(imCancelBtn) imCancelBtn.addEventListener("click",function(){
  IMPORT_PENDING=null;
  document.getElementById("ov-import").style.display="none";
});

function _findKlub(nazwa){
  for(var i=0;i<DB.length;i++){ if((DB[i].name||"")===nazwa) return i; }
  return -1;
}

function _deepLink(){
  // 1) sklad konkretnej druzyny (uzywane przez zakladke Roster)
  var qk=_Q.get("klub"),qs=_Q.get("sezon"),qd=_Q.get("druzyna");
  if(qk&&qs&&qd){
    var i=_findKlub(qk);
    if(i>=0){
      var sez=DB[i].sezony||{};
      if(sez[qs]&&(sez[qs]||[]).indexOf(qd)>=0){ goTeam(i,qs,qd); return true; }
    }
  }
  // 2) edycja wskazanego klubu
  var qe=_Q.get("edit");
  if(qe){
    var j=_findKlub(qe);
    if(j>=0){ doEdit(j); return true; }
  }
  // 3) nowy klub — tworzymy szkielet i od razu otwieramy edycje
  if(_Q.get("new")==="1"){
    // Szkielet tylko w pamieci — na dysk trafi dopiero po kliknieciu Zapisz.
    DB.push({name:"Nowy klub",ext:false,sezony:{}});
    KI=DB.length-1; NEW_PENDING=KI; doEdit(KI); return true;
  }
  return false;
}

draw();
_deepLink();   // brak parametrow -> serwer przekierowuje na /klub
})();
