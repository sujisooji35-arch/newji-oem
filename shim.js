/* ───────── 깃허브(Supabase) 실행용 연결 계층 ─────────
   claude.ai 아티팩트 기능(window.claude.use)을 Supabase·브라우저 기능으로 대신합니다.
   설정값은 config 객체만 바꾸면 됩니다. */
(function(){
'use strict';
var config={
  url:'https://etalcfamjysspyygmibn.supabase.co',
  anonKey:'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImV0YWxjZmFtanlzc3B5eWdtaWJuIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTM2ODM0OTUsImV4cCI6MjEwOTI1OTQ5NX0.LnKWi9kp0NPKQy-eS02JQR8jNTUMAGugIe8Nak-ZqMQ',
  sheetId:'1-bgJkq_OGNVOyocskDKZXpINjzhLOK5kxSuKfGbOEEo',
  sheetTabs:['필름_에스쇼콜라','필름_부산','필름_에티오피아','필름_콜롬비아 디카페인','필름_코스타리카 디카페인','필름_에티오피아 디카페인','필름_므쵸베리','생분해필터']
};
var SB=window.supabase.createClient(config.url,config.anonKey,{auth:{persistSession:false}});
window.__NEWJI_SB=SB;

/* ── 표 이름·열 이름 대응 ── */
var TBL={batches:'newji_batches',events:'newji_events',materials:'newji_materials',orders:'newji_orders',receipts:'newji_receipts',beans:'newji_beans',products:'newji_products',settings:'newji_settings',logs:'newji_logs'};
var MAP={batches:{order:'order_id',lastGrind:'last_grind',start:'start_date',end:'end_date'},events:{endDate:'end_date'},
  materials:{group:'grp',order:'ord',perBean:'per_bean',sheetTab:'sheet_tab',lossNotes:'loss_notes'},orders:{},receipts:{order:'order_id'},beans:{order:'ord'},products:{order:'ord'},logs:{user:'usr',docId:'doc_id'}};
var COLS={batches:['order_id','item','bean','batches','kg','qty','roast','recv','grind','last_grind','start_date','end_date','status','memo','daily','actual'],
  logs:['at','usr','uid','action','coll','doc_id','title','fields','ctx'],
  events:['type','title','date','end_date','time','contact','memo','done'],
  materials:['name','code','grp','unit','basis','scope','loss','ord','memo','link','sheet_tab','per','per_bean','log','loss_notes'],
  orders:['title','no','date','due','status','memo','lines'],receipts:['date','order_id','memo','rows'],
  beans:['name','code','decaf','active','ord','memo'],products:['name','pack','comp','active','ord','memo']};
var DATES={roast:1,recv:1,grind:1,last_grind:1,start_date:1,end_date:1,date:1,due:1};
var NUMS={batches:1,kg:1,qty:1,loss:1,ord:1,pack:1};
var INTS={qty:1,ord:1,pack:1};
function inv(c){var m=MAP[c]||{},r={};Object.keys(m).forEach(function(k){r[m[k]]=k});return r}
function toRow(c,doc){
  if(c==='settings')return {data:doc};
  var m=MAP[c]||{},cols=COLS[c]||[],row={};
  Object.keys(doc||{}).forEach(function(k){
    var col=m[k]||k;if(cols.indexOf(col)<0)return;var v=doc[k];
    if(DATES[col])v=v?String(v).slice(0,10):null;
    else if(NUMS[col]){if(v===''||v==null)v=null;else{v=Number(v);if(!isFinite(v))v=null;else if(INTS[col])v=Math.round(v)}}
    else if(v===undefined)v=null;
    row[col]=v;
  });
  return row;
}
function fromRow(c,row){
  if(c==='settings')return row.data||{};
  var iv=inv(c),o={};
  Object.keys(row).forEach(function(col){if(col==='id'||col==='updated_at')return;var v=row[col];if(v===null||v===undefined)return;o[iv[col]||col]=v});
  return o;
}
var loaders={};
function errCode(e){return {code:'unavailable',message:(e&&e.message)||String(e)}}
function collection(c){
  var T=TBL[c];
  return {
    onSnapshot:function(next,error){
      var busy=false,again=false;
      function load(){
        if(busy){again=true;return}busy=true;
        (c==='logs'?SB.from(T).select('*').order('at',{ascending:false}).limit(500):SB.from(T).select('*')).then(function(res){
          busy=false;
          if(res.error){if(error)error(errCode(res.error));}
          else next({docs:(res.data||[]).map(function(r){var d=fromRow(c,r);return {id:r.id,exists:true,data:function(){return d}}})});
          if(again){again=false;load()}
        },function(e){busy=false;if(error)error(errCode(e))});
      }
      (loaders[c]=loaders[c]||[]).push(load);
      load();
      var t=null;
      SB.channel('newji-'+c+'-'+Math.random().toString(36).slice(2,7)).on('postgres_changes',{event:'*',schema:'public',table:T},function(){clearTimeout(t);t=setTimeout(load,250)}).subscribe();
      setInterval(function(){if(document.visibilityState==='visible')load()},60000);
      return function(){};
    },
    doc:function(id){
      function done(res){if(res.error)throw {code:'invalid_argument',message:res.error.message};(loaders[c]||[]).forEach(function(f){f()})}
      return {
        set:function(d){var row=toRow(c,d);row.id=id;row.updated_at=new Date().toISOString();return SB.from(T).upsert(row).then(done)},
        update:function(d){var row=toRow(c,d);row.updated_at=new Date().toISOString();return SB.from(T).update(row).eq('id',id).then(done)},
        delete:function(){return SB.from(T).delete().eq('id',id).then(done)}
      };
    }
  };
}
var db={collection:collection,doc:function(path){var p=path.split('/');return collection(p[0]).doc(p[1])}};

/* ── 엑셀 다운로드 ── */
var downloads={save:function(req){
  return new Promise(function(res){
    var blob=req.data instanceof Blob?req.data:new Blob([req.data],{type:'application/octet-stream'});
    var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=req.filename||'download';document.body.appendChild(a);a.click();
    setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},1500);res({status:'saved'});
  });
}};

/* ── 구글 시트: '링크가 있는 모든 사용자 보기'로 공유된 시트를 CSV로 읽음 ── */
function parseCSV(t){
  var rows=[],row=[],f='',q=false;
  for(var i=0;i<t.length;i++){var ch=t[i];
    if(q){if(ch==='"'){if(t[i+1]==='"'){f+='"';i++}else q=false}else f+=ch}
    else if(ch==='"')q=true;else if(ch===','){row.push(f);f=''}else if(ch==='\n'){row.push(f);rows.push(row);row=[];f=''}else if(ch!=='\r')f+=ch;}
  if(f!==''||row.length){row.push(f);rows.push(row)}
  return rows;
}
function knownTabs(){
  var s={};config.sheetTabs.forEach(function(t){s[t]=1});
  try{(window.__NEWJI_MATS||[]).forEach(function(m){if(m.sheetTab)s[m.sheetTab]=1})}catch(e){}
  return Object.keys(s);
}
var mcp={
  callTool:function(server,tool,input){
    if(tool==='get_spreadsheet')return Promise.resolve({payload:{sheets:knownTabs().map(function(t){return {properties:{title:t}}})}});
    if(tool==='get_values'){
      var tab=String(input.range||'').split('!')[0].replace(/^'|'$/g,'');
      var url='https://docs.google.com/spreadsheets/d/'+config.sheetId+'/gviz/tq?tqx=out:csv&headers=0&range=A1:J1000&sheet='+encodeURIComponent(tab);
      return fetch(url,{credentials:'omit'}).then(function(r){
        if(!r.ok)throw {code:'tool_error'};
        return r.text();
      }).then(function(txt){
        if(/<html/i.test(txt.slice(0,200)))throw {code:'tool_error'};
        var vals=parseCSV(txt),head=vals.slice(0,3).map(function(r){return r.join(' ')}).join(' ');
        // 없는 탭을 요청하면 구글이 첫 탭을 돌려줄 수 있으므로, 제목줄에 탭 이름이 있을 때만 사용
        if(head.replace(/\s/g,'').indexOf(tab.replace(/\s/g,''))<0)return {payload:{range:input.range,values:[]}};
        return {payload:{range:input.range,values:vals}};
      },function(e){throw (e&&e.code)?e:{code:'tool_error'}});
    }
    return Promise.reject({code:'bad_request'});
  }
};

/* ── 로그인 ── */
var SESSION=null;
try{SESSION=JSON.parse(localStorage.getItem('newji-session')||'null')}catch(e){}
if(SESSION&&(!SESSION.at||Date.now()-SESSION.at>30*864e5))SESSION=null;
var loginResolve;var loggedIn=new Promise(function(r){loginResolve=r});
function saveSession(s){SESSION=s;try{localStorage.setItem('newji-session',JSON.stringify(s))}catch(e){}}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})}
function el(html){var d=document.createElement('div');d.innerHTML=html;return d.firstElementChild}
function showLogin(first){
  var box=el('<div class="lgwrap" id="lgwrap"><form class="lgbox" id="lgform"><h1>뉴지코리아 OEM 일정</h1><p class="lgsub">'+(first?'처음 사용합니다. 관리자 이름과 비밀번호를 정해 주세요.':'이름과 비밀번호를 입력하세요.')+'</p>'+
    '<label for="lg_name">이름</label><input id="lg_name" autocomplete="username" required>'+
    '<label for="lg_pw">비밀번호</label><input id="lg_pw" type="password" autocomplete="'+(first?'new-password':'current-password')+'" required minlength="4">'+
    '<p class="lgerr" id="lgerr" hidden></p><button class="btn primary" type="submit">'+(first?'관리자 만들고 시작':'로그인')+'</button>'+
    '<p class="lghint">'+(first?'이 관리자가 다른 사용자를 추가합니다.':'계정이 없으면 관리자에게 등록을 요청하세요.')+'</p></form></div>');
  document.body.appendChild(box);
  setTimeout(function(){var n=document.getElementById('lg_name');if(n)n.focus()},50);
  document.getElementById('lgform').addEventListener('submit',function(e){
    e.preventDefault();
    var n=document.getElementById('lg_name').value.trim(),p=document.getElementById('lg_pw').value,er=document.getElementById('lgerr'),b=e.target.querySelector('button');
    b.disabled=true;er.hidden=true;
    SB.rpc(first?'newji_bootstrap':'newji_login',{p_name:n,p_pw:p}).then(function(r){
      b.disabled=false;
      if(r.error){er.textContent='서버에 연결하지 못했습니다. 잠시 후 다시 시도해 주세요. ('+r.error.message+')';er.hidden=false;return}
      var d=r.data||{};
      if(!d.ok){er.textContent=d.msg||'이름 또는 비밀번호가 맞지 않습니다.';er.hidden=false;return}
      saveSession({name:d.name,role:d.role,at:Date.now()});box.remove();loginResolve(SESSION);userChip();
    });
  });
}
function logout(){try{localStorage.removeItem('newji-session')}catch(e){}location.reload()}
function userChip(){
  var c=document.getElementById('uchip');if(c)c.remove();
  var roleTxt={admin:'관리자',editor:'편집',viewer:'보기 전용'}[SESSION.role]||SESSION.role;
  c=el('<div class="uchip" id="uchip"><span><b>'+esc(SESSION.name)+'</b> · '+roleTxt+'</span><button type="button" data-u="pw">비밀번호 변경</button>'+(SESSION.role==='admin'?'<button type="button" data-u="users">사용자 관리</button>':'')+'<button type="button" data-u="out">로그아웃</button></div>');
  document.body.appendChild(c);
  c.addEventListener('click',function(e){var b=e.target.closest('[data-u]');if(!b)return;var u=b.dataset.u;if(u==='out')logout();else if(u==='pw')pwModal();else if(u==='users')usersModal()});
}
function modal(title,body){
  var m=el('<div class="umodal"><div class="uscrim"></div><div class="ubox" role="dialog" aria-modal="true" aria-label="'+esc(title)+'"><div class="uhd"><h2>'+esc(title)+'</h2><button type="button" class="xbtn" data-x="1" aria-label="닫기">×</button></div><div class="ubd">'+body+'</div></div></div>');
  document.body.appendChild(m);m.addEventListener('click',function(e){if(e.target.dataset.x||e.target.classList.contains('uscrim'))m.remove()});return m;
}
function pwModal(){
  var m=modal('비밀번호 변경','<form id="pwf" class="uform"><label>현재 비밀번호<input type="password" id="pw_old" required></label><label>새 비밀번호 (4자 이상)<input type="password" id="pw_new" required minlength="4"></label><p class="lgerr" id="pwerr" hidden></p><button class="btn primary">변경</button></form>');
  m.querySelector('#pwf').addEventListener('submit',function(e){e.preventDefault();
    SB.rpc('newji_change_pw',{p_name:SESSION.name,p_old:m.querySelector('#pw_old').value,p_new:m.querySelector('#pw_new').value}).then(function(r){
      var d=r.data||{},er=m.querySelector('#pwerr');if(r.error||!d.ok){er.textContent=(d&&d.msg)||'변경하지 못했습니다.';er.hidden=false;return}m.remove();alertMsg('비밀번호를 바꿨습니다.')})});
}
var ADMIN_PW=null;
function usersModal(){
  var m=modal('사용자 관리','<div id="umain"></div>'),main=m.querySelector('#umain');
  function ask(){main.innerHTML='<form id="apw" class="uform"><label>관리자 비밀번호 확인<input type="password" id="a_pw" required></label><p class="lgerr" id="aerr" hidden></p><button class="btn primary">확인</button></form>';
    main.querySelector('#apw').addEventListener('submit',function(e){e.preventDefault();ADMIN_PW=main.querySelector('#a_pw').value;list()})}
  function call(action,extra){return SB.rpc('newji_admin',Object.assign({a_name:SESSION.name,a_pw:ADMIN_PW,p_action:action},extra||{})).then(function(r){return r.error?{ok:false,msg:r.error.message}:(r.data||{})})}
  function list(){call('list').then(function(d){
    if(!d.ok){ADMIN_PW=null;ask();var er=main.querySelector('#aerr');er.textContent=d.msg||'확인 실패';er.hidden=false;return}
    var h='<table class="utable"><thead><tr><th>이름</th><th>권한</th><th>상태</th><th>최근 로그인</th><th></th></tr></thead><tbody>';
    d.users.forEach(function(u){var me=u.name===SESSION.name;
      h+='<tr><td><b>'+esc(u.name)+'</b></td><td>'+(me?'관리자':'<select data-role="'+esc(u.name)+'">'+[['admin','관리자'],['editor','편집'],['viewer','보기 전용']].map(function(r){return '<option value="'+r[0]+'"'+(u.role===r[0]?' selected':'')+'>'+r[1]+'</option>'}).join('')+'</select>')+'</td>'+
        '<td>'+(u.active?'사용 중':'<span class="lgerr">중지</span>')+'</td><td class="num">'+(u.last_login?String(u.last_login).slice(0,16).replace('T',' '):'-')+'</td>'+
        '<td>'+(me?'':'<button type="button" class="btn sm" data-setpw="'+esc(u.name)+'">비밀번호 재설정</button> <button type="button" class="btn sm" data-act="'+esc(u.name)+'" data-v="'+(!u.active)+'">'+(u.active?'사용 중지':'다시 사용')+'</button>')+'</td></tr>'});
    h+='</tbody></table><form id="uadd" class="uform row"><label>새 사용자 이름<input id="u_name" required></label><label>초기 비밀번호<input id="u_pw" required minlength="4"></label><label>권한<select id="u_role"><option value="editor">편집</option><option value="viewer">보기 전용</option><option value="admin">관리자</option></select></label><button class="btn primary">추가</button></form><p class="lgerr" id="uerr" hidden></p>';
    main.innerHTML=h;
    function fail(d){var er=main.querySelector('#uerr');er.textContent=d.msg||'처리하지 못했습니다.';er.hidden=false}
    main.querySelector('#uadd').addEventListener('submit',function(e){e.preventDefault();call('add',{p_name:main.querySelector('#u_name').value,p_pw:main.querySelector('#u_pw').value,p_role:main.querySelector('#u_role').value}).then(function(d){if(d.ok)list();else fail(d)})});
    main.addEventListener('change',function(e){var n=e.target.dataset.role;if(n)call('role',{p_name:n,p_role:e.target.value}).then(function(d){if(!d.ok)fail(d)})});
    main.addEventListener('click',function(e){var b=e.target.closest('button');if(!b)return;
      if(b.dataset.act)call('active',{p_name:b.dataset.act,p_active:b.dataset.v==='true'}).then(function(d){if(d.ok)list();else fail(d)});
      if(b.dataset.setpw){var row=b.closest('td');row.innerHTML='<input type="password" placeholder="새 비밀번호 4자 이상" class="inpw"> <button type="button" class="btn sm" data-dopw="'+esc(b.dataset.setpw)+'">저장</button>'}
      if(b.dataset.dopw){var v=b.parentNode.querySelector('.inpw').value;call('setpw',{p_name:b.dataset.dopw,p_pw:v}).then(function(d){if(d.ok)list();else fail(d)})}
    });
  })}
  if(ADMIN_PW)list();else ask();
}
function alertMsg(t){var x=document.getElementById('toast');if(x){x.textContent=t;x.hidden=false;setTimeout(function(){x.hidden=true},2600)}}

/* ── 사용자·권한 ── */
var user={
  isOwner:function(){return Promise.resolve(SESSION&&SESSION.role==='admin')},
  canEdit:function(){return Promise.resolve(SESSION&&SESSION.role!=='viewer')},
  can:function(){return Promise.resolve(SESSION?SESSION.role!=='viewer':null)},
  id:function(){return Promise.resolve(SESSION?SESSION.name:null)},
  me:function(){return Promise.resolve({id:SESSION&&SESSION.name,name:SESSION&&SESSION.name||''})}
};

/* ── 앱이 부르는 window.claude.use 를 대신함 (로그인 후에 응답) ── */
window.claude={use:function(name){
  return loggedIn.then(function(){
    if(name==='db')return db;
    if(name==='downloads')return downloads;
    if(name==='user')return user;
    if(name==='mcp')return mcp;
    return null;
  });
}};

function start(){
  if(SESSION){loginResolve(SESSION);userChip();return}
  SB.rpc('newji_has_users').then(function(r){
    if(r.error){showLogin(false);var er=document.getElementById('lgerr');if(er){er.textContent='서버 설정이 아직 안 된 것 같습니다. supabase_login_setup.sql 을 실행했는지 확인해 주세요. ('+r.error.message+')';er.hidden=false}return}
    showLogin(r.data===false);
  });
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else start();
})();
