/* Google Drive direct OAuth browser read-only application. Never persist access tokens. */
(function(){'use strict';
const $=id=>document.getElementById(id);
const targets=['М. Жалолиддинов','П. Мўминов','А. Маҳкамов'];
const norm=s=>String(s??'').replace(/\s+/g,'').toLowerCase().replaceAll('махкамов','маҳкамов').replaceAll('муминов','мўминов');
const DEFAULT_FILE_ID='1Cj8HFSQbqZ9ArFENYuT51sp1pb1xU8tDrOMu7vaPb-4';
let tokenClient=null, token='',expires=0,lastModified='',busy=false,refreshTimer=null;
const section=document.createElement('section');section.className='panel';section.style.border='2px solid #b8d4ef';
section.innerHTML=`<h2 style="margin-top:0">Google Drive — тўғридан-тўғри уланиш</h2>
<p class="muted">Ушбу версия Drive for desktop ва Python талаб қилмайди. Файл Google Drive API орқали фақат ўқилади; асл Excel ўзгартирилмайди. Ҳисобот ҳар 60 сонияда текширилади.</p>
<div class="filters" style="grid-template-columns:repeat(2,minmax(220px,1fr))">
<div class="field"><label>Google OAuth Client ID (бир марта созланади)</label><input style="width:100%" id="oauth-client" placeholder="...apps.googleusercontent.com" autocomplete="off"></div>
<div class="field"><label>Google Drive файл ID</label><input style="width:100%" id="drive-file-id" autocomplete="off"></div>
</div>
<div class="flex" style="margin-top:12px"><button class="primary" id="google-login">Google аккаунтга кириш</button><button id="google-update">↻ Ҳозир янгилаш</button><button id="google-logout">Уланишни узиш</button><span id="google-state" class="muted">Ҳали уланмаган</span></div>
<p class="muted">Дастурни <b>HTTPS</b> сайтда очинг (масалан, GitHub Pages). OAuth Client ID яратиш ва сайтнинг аниқ URL манбасини <i>Authorized JavaScript origins</i> га киритиш шарт. Файл идентификатори сақланади; Google токени сақланмайди.</p>
<p class="muted" id="google-file-meta"></p>`;
document.querySelector('main.wrap').insertBefore(section,document.querySelector('main.wrap').firstChild);
$('drive-file-id').value=localStorage.getItem('bolim_drive_fileid')||DEFAULT_FILE_ID;
$('oauth-client').value=localStorage.getItem('bolim_oauth_client')||'';
const msg=(s,error=false)=>{$('google-state').textContent=s;$('google-state').style.color=error?'#b42339':'#18835c'};
function strings(root,tag){return [...root.getElementsByTagNameNS('*',tag)];}
function text(el,tag){const n=strings(el,tag)[0];return n?n.textContent:'';}
function parseXML(s){let d=new DOMParser().parseFromString(s,'application/xml');if(d.querySelector('parsererror'))throw new Error('Excel XML ни ўқиш хатоси');return d;}
const dateFrom=(n)=>{if(typeof n==='number'&&n>40000&&n<65000){const d=new Date(Date.UTC(1899,11,30)+n*86400000);return d.toISOString().slice(0,10);}const m=String(n??'').match(/(20\d\d)[-. /](\d\d?)[-. /](\d\d?)/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:''};
function weekDate(name,first){let m=name.match(/(\d{1,2})[.-](\d{1,2})[.-](20\d{2})/);if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;m=name.match(/(\d{1,2})[-.](\d{1,2})[.]?(\d{1,2})[.]?(20\d{2})/);if(m)return `${m[4]}-${m[3].padStart(2,'0')}-${m[1].padStart(2,'0')}`;m=first.toLowerCase().match(/(20\d{2})\s*йил\s*(\d{1,2})\s*([а-яёҳқғў]+)/);if(m){const months={'январ':1,'феврал':2,'март':3,'апрел':4,'май':5,'июн':6,'июл':7,'август':8,'сентябр':9,'октябр':10,'ноябр':11,'декабр':12};for(const [k,v] of Object.entries(months))if(m[3].startsWith(k))return `${m[1]}-${String(v).padStart(2,'0')}-${m[2].padStart(2,'0')}`;}return ''}
function pathJoin(raw){let p=raw.replace(/^\//,'');if(!p.startsWith('xl/'))p='xl/'+p;const stack=[];for(const v of p.split('/')){if(v==='..')stack.pop();else if(v!=='.')stack.push(v);}return stack.join('/')}
async function readExcel(blob){
 const zip=await JSZip.loadAsync(blob);const file=async p=>{if(!zip.file(p))throw Error('Excel ичида '+p+' йўқ');return zip.file(p).async('string')};
 let ss=[];if(zip.file('xl/sharedStrings.xml')){const d=parseXML(await file('xl/sharedStrings.xml'));ss=strings(d,'si').map(si=>strings(si,'t').map(t=>t.textContent||'').join(''))}
 const book=parseXML(await file('xl/workbook.xml'));
 const rels=parseXML(await file('xl/_rels/workbook.xml.rels'));
 const relMap={};for(const e of strings(rels,'Relationship'))relMap[e.getAttribute('Id')]=e.getAttribute('Target');
 const output=[];
 for(const sh of strings(book,'sheet')){
  const name=sh.getAttribute('name');const rid=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');const target=pathJoin(relMap[rid]||'');
  if(!zip.file(target))continue;
  const doc=parseXML(await file(target));const rows=[];
  for(const row of strings(doc,'sheetData').flatMap(n=>[...n.children]).filter(n=>n.localName==='row')){
   const cols={};let max=-1;
   for(const c of [...row.children].filter(n=>n.localName==='c')){
    const m=(c.getAttribute('r')||'').match(/^[A-Z]+/);if(!m)continue;let index=0;for(const l of m[0])index=index*26+l.charCodeAt(0)-64;index--;
    let value='';const type=c.getAttribute('t');
    if(type==='inlineStr')value=strings(c,'t').map(t=>t.textContent||'').join('');
    else {const v=strings(c,'v')[0];if(!v)continue;value=v.textContent||'';if(type==='s')value=ss[Number(value)]??'';else if(!['str','e','b'].includes(type)&&value.trim()!==''&&!isNaN(Number(value)))value=Number(value)}
    cols[index]=value;max=Math.max(max,index);
   }
   if(max>=0)rows.push({index:Number(row.getAttribute('r'))||rows.length+1,cells:Array.from({length:max+1},(_,i)=>cols[i]??'')});
  }
  const header=rows.slice(0,12).findIndex(r=>r.cells.some(v=>['Белгиланган вазифалар','Мазмуни','Амалга ошириладиган чора-тадбирлар'].includes(String(v).trim())));
  if(header<0)continue;
  const isExec=name.trim().toUpperCase()==='ИЖРО',isPlan=name.includes('Иш режа');const week=weekDate(name,String(rows[0]?.cells[0]??''));
  for(const row of rows.slice(header+1)){
   const r=row.cells,at=i=>r.length>i?String(r[i]??'').trim():'';
   const title=at(isExec?4:1),responsible=at(isExec?5:isPlan?4:2);
   const mem=targets.filter(t=>norm(responsible).includes(norm(t)));if(!title||!mem.length)continue;
   if(!isExec&&!isPlan&&!/^\d+(\.0)?$/.test(at(0)))continue;
   let raw=isPlan?'':at(isExec?10:4);if(isExec&&!raw)raw=at(9);const st=norm(raw);
   const status=isPlan?'Маълум эмас':st.includes('бажарилмади')||st.includes('қолдирилди')?'Бажарилмаган':st.includes('бажарилди')||st.includes('бажарилган')||st.includes('жавобкиритилди')?'Бажарилди':st.includes('бажарилмоқда')||st.includes('жараён')||st.includes('ишланмоқда')?'Жараёнда':'Маълум эмас';
   const term=at(isExec||isPlan?3:5),deadline=dateFrom(r[isExec||isPlan?3:5]);
   output.push({id:name+':'+row.index,sheet:name,row:row.index,title,responsible,members:mem,detail:at(isExec?7:isPlan?2:3),status,rawStatus:raw,deadline,termText:term,week,priority:!isExec&&!isPlan?at(6):'',kind:isExec?'Ижро':isPlan?'Иш режа':'Ҳафталик'});
  }
 }
 return output;
}

function fileId(){const v=$('drive-file-id').value.trim();return (v.match(/\/d\/([a-zA-Z0-9_-]+)/)||[])[1]||(v.match(/[?&]id=([a-zA-Z0-9_-]+)/)||[])[1]||v;}
function savedSettings(){localStorage.setItem('bolim_drive_fileid',fileId());localStorage.setItem('bolim_oauth_client',$('oauth-client').value.trim());}
async function api(url){const res=await fetch(url,{headers:{'Authorization':'Bearer '+token},cache:'no-store'});if(!res.ok){let m='HTTP '+res.status;try{const d=await res.json();m=d.error?.message||m}catch{};if(res.status===401){token='';expires=0;}throw Error(m)}return res;}
async function poll(force=false){if(!token||busy)return;if(Date.now()>expires){msg('Google рухсати муддати тугади. Қайта киринг.',true);token='';return;}
busy=true;try{const id=fileId();if(!/^[a-zA-Z0-9_-]{15,}$/.test(id))throw Error('Файл ID нотўғри');
const metadata=await (await api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=id,name,mimeType,modifiedTime,size,capabilities(canDownload)&supportsAllDrives=true`)).json();
if(metadata.capabilities?.canDownload===false)throw Error('Бу файлни ўқиш/юклаш ҳуқуқи берилмаган');
$('google-file-meta').textContent='Манба: '+metadata.name+' · Drive да охирги ўзгариш: '+(metadata.modifiedTime?new Date(metadata.modifiedTime).toLocaleString('ru-RU'):'—');
if(!force&&metadata.modifiedTime===lastModified){msg('● Ўзгариш йўқ · '+new Date().toLocaleTimeString('ru-RU'));return;}
msg('Google Drive дан Excel ўқилмоқда...');
const isSheet=metadata.mimeType==='application/vnd.google-apps.spreadsheet';
const url=isSheet?`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}/export?mimeType=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet`:`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?alt=media&supportsAllDrives=true`;
const blob=await (await api(url)).blob();const rows=await readExcel(blob);if(!rows.length)throw Error('Танланган файлда уч ходимга тегишли топшириқ аниқланмади');
tasks=rows;lastModified=metadata.modifiedTime||'';updateSheets();redraw();
$('sync').textContent='● Google Drive · '+new Date().toLocaleTimeString('ru-RU');
$('note').textContent='Бевосита Google Drive дан олинган маълумотлар. Охирги янгиланиш: '+new Date().toLocaleString('ru-RU')+'. Файлнинг асли дастур томонидан ўзгартирилмайди.';
msg('● Уланган · '+rows.length+' та ёзув · '+new Date().toLocaleTimeString('ru-RU'));
}catch(err){console.error(err);msg('⚠ '+err.message,true)}finally{busy=false}}
function initToken(){const client=$('oauth-client').value.trim();if(!/^\d+-[a-zA-Z0-9_-]+\.apps\.googleusercontent\.com$/.test(client))throw Error('OAuth Client ID формати нотўғри');
if(location.protocol!=='https:'&&location.hostname!=='localhost'&&location.hostname!=='127.0.0.1')throw Error('Google кириш учун HTTPS сайт талаб қилинади');
if(!window.google?.accounts?.oauth2)throw Error('Google Identity Services юкланмади. Интернетни текширинг');
tokenClient=google.accounts.oauth2.initTokenClient({client_id:client,scope:'https://www.googleapis.com/auth/drive.readonly',callback:async(response)=>{if(response.error){msg('Google рухсат хатоси: '+response.error,true);return}token=response.access_token;expires=Date.now()+(Number(response.expires_in||3600)-90)*1000;lastModified='';msg('Google уланди, файл ўқилмоқда...');await poll(true)},error_callback:e=>msg('Google ойнаси хатоси: '+(e.message||e.type||'номаълум'),true)});
}
$('google-login').onclick=()=>{try{savedSettings();initToken();tokenClient.requestAccessToken({prompt:'consent'})}catch(e){msg(e.message,true)}};
$('google-update').onclick=()=>poll(true);
$('google-logout').onclick=()=>{if(token&&window.google?.accounts?.oauth2)google.accounts.oauth2.revoke(token,()=>{});token='';expires=0;msg('Уланиш узилди')};
$('drive-file-id').addEventListener('change',()=>{savedSettings();lastModified=''});
$('oauth-client').addEventListener('change',savedSettings);
refreshTimer=setInterval(()=>poll(false),60000);
})();
