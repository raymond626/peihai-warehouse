// ═══════════════════════════════════════════════════════════════════
// PEIHAI WAREHOUSE MANAGEMENT SYSTEM — v2.0
// ═══════════════════════════════════════════════════════════════════

// ─── Material taxonomy ──────────────────────────────────────────────
const MAT = {
  L:{label:'皮類 Leather',      types:{C:'牛 Cow',S:'羊 Sheep',P:'豬 Pig',X:'特 Special'},     specs:{1:'面皮 Full Grain',2:'反毛 Suede',3:'巴戈 Nubuck',4:'狼皮 Action',5:'三層 Split',6:'帶毛 Hair-on',7:'貼膜 Laminated',8:'後處理 Treated',9:'內裡 Lining'}},
  S:{label:'人造皮 Synthetic',  types:{M:'超纖 Microfiber',N:'不織布 Non-woven',W:'布底 Woven',T:'TPU',V:'PVC'},       specs:{1:'面皮紋 Smooth',2:'絨面 Suede-like',3:'亮面 Patent',4:'壓紋 Embossed',5:'毛 Plush'}},
  T:{label:'布類 Textile',      types:{M:'網布 Mesh',C:'帆布 Canvas',K:'針織 Knit',J:'提花 Jacquard',G:'佳績布 Jiaji',L:'萊卡布 Lycra',S:'特織 Special'},specs:{1:'單層 Single',2:'雙層 Double',3:'三明治 Sandwich',4:'彈性 Elastic',5:'防水 WP'}},
  W:{label:'條狀類 Webbing',    types:{W:'織帶 Webbing',E:'鬆緊帶 Elastic',L:'鞋帶 Lace',P:'滾邊 Binding'}},
  A:{label:'副料 Auxiliary',    types:{C:'港寶 Counter',S:'泡棉 Foam',E:'EVA',H:'高發泡 Hi-Foam',I:'中底 Insole',F:'補強 Reinforcement'}},
  O:{label:'大底 Outsole',      types:{},specs:{}},
  H:{label:'五金 Hardware',     types:{E:'鞋眼 Eyelet',H:'鉤 Hook',B:'扣 Buckle',Z:'拉鍊 Zipper',S:'亮片 Sequin'}}
};

const CAT_LABELS = {L:'L 皮類',S:'S 人造皮',T:'T 布類',W:'W 條狀',A:'A 副料',O:'O 大底',H:'H 五金'};
const OUTSOLE_SIZES = {20:{uk:'4',us:'4.5C',jp:'10.5'},21:{uk:'4.5',us:'5C',jp:'11'},22:{uk:'5.5',us:'6C',jp:'12'},23:{uk:'6.5',us:'7C',jp:'13'},24:{uk:'7',us:'7.5C',jp:'13.5'},25:{uk:'7.5',us:'8C',jp:'14'},26:{uk:'8.5',us:'9C',jp:'15'},27:{uk:'9.5',us:'10C',jp:'16'},28:{uk:'10.5',us:'11C',jp:'17'},29:{uk:'11.5',us:'12C',jp:'18'},30:{uk:'12',us:'12.5C',jp:'18.5'},31:{uk:'12.5',us:'13C',jp:'19'},32:{uk:'13.5',us:'1Y',jp:'20'},33:{uk:'1',us:'1.5Y',jp:'20.5'},34:{uk:'2',us:'2.5Y',jp:'21.5'},35:{uk:'2.5',us:'3Y',jp:'22'},36:{uk:'3.5',us:'4',jp:'23'},37:{uk:'4.5',us:'5',jp:'23.5'},38:{uk:'5',us:'5.5',jp:'24'},39:{uk:'6',us:'6.5',jp:'24.5'},40:{uk:'6',us:'7',jp:'25'},41:{uk:'7',us:'8',jp:'26'},42:{uk:'7.5',us:'8.5',jp:'26.5'},43:{uk:'8.5',us:'9.5',jp:'27.5'},44:{uk:'9',us:'10',jp:'28'},45:{uk:'10',us:'11',jp:'29'},46:{uk:'11',us:'12',jp:'30'},47:{uk:'11.5',us:'12.5',jp:'30.5'},48:{uk:'12.5',us:'13.5',jp:'31.5'},49:{uk:'13.5',us:'14.5',jp:'32.5'},50:{uk:'14.5',us:'15.5',jp:'33.5'}};

// ─── Global state ───────────────────────────────────────────────────
let inventory = [];
let customBrands = [];
let batchCart = [];
let db = null;
let supabaseClient = null;
let cloudEnabled = false;
let currentUser = null;
let currentRole = 'local';
let realtimeChannel = null;
let suppressCloudSave = false;
let suppressRealtimeUntil = 0;
let cloudSaveTimer = null;
let cloudLoadTimer = null;
const pendingCloudItems = new Map();
let dataDirectoryHandle = null;
let scannerStream = null;
let scannerActive = false;
let scannerTargetId = '';
let currentDetailCode = null;
let pendingDeleteCode = null;

const DB_NAME = 'Peihai_WMS_DB_V26';
const BACKUP_FILE = 'peihai-inventory-backup.json';
const LIBS = {
  qr: 'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',
  xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
};
const scriptLoads = {};

// ─── Helpers ────────────────────────────────────────────────────────
function now() {
  const d = new Date();
  return d.getFullYear()+'/'+pad(d.getMonth()+1)+'/'+pad(d.getDate())+' '+pad(d.getHours())+':'+pad(d.getMinutes());
}
function pad(n){ return String(n).padStart(2,'0'); }
function esc(v){ return String(v||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]); }
function escapeRegExp(v){ return String(v||'').replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function splitBulkValues(v){
  return String(v||'').split(/[\n\r,，;；]+/).map(s=>s.trim()).filter(Boolean);
}
function parseBulkCodeInput(v, defaultQty=1){
  const raw=String(v||'').trim();
  if(!raw) return [];
  const rows=[];
  raw.split(/\r?\n/).forEach(line=>{
    const text=line.trim();
    if(!text) return;
    const tabParts=text.split(/\t+/).map(s=>s.trim()).filter(Boolean);
    if(tabParts.length>=2){
      const qty=parseFloat(tabParts[1]);
      rows.push({code:tabParts[0].toUpperCase(), qty:qty>0?qty:defaultQty});
      return;
    }
    const commaParts=text.split(/[，,;；]+/).map(s=>s.trim()).filter(Boolean);
    if(commaParts.length===2 && !isNaN(parseFloat(commaParts[1]))){
      const qty=parseFloat(commaParts[1]);
      rows.push({code:commaParts[0].toUpperCase(), qty:qty>0?qty:defaultQty});
      return;
    }
    commaParts.forEach(code=>rows.push({code:code.toUpperCase(), qty:defaultQty}));
  });
  return rows;
}
function nextSeqForPrefix(prefix, reservedCodes=[]){
  const re=new RegExp('^'+escapeRegExp(prefix)+'-(\\d+)(?:-|$)','i');
  const codes=inventory.map(i=>i.code||'').concat(reservedCodes||[]);
  const maxSeq=codes.reduce((mx,code)=>{
    const m=String(code||'').match(re);
    if(!m) return mx;
    const n=parseInt(m[1],10);
    return n>mx?n:mx;
  },0);
  return String(maxSeq+1).padStart(4,'0');
}
function nextMaterialCode(prefix, colorCode, reservedCodes=[]){
  return prefix+'-'+nextSeqForPrefix(prefix,reservedCodes)+'-'+String(colorCode||'').trim();
}
function touchCloudSyncWindow(ms=5000){
  suppressRealtimeUntil=Math.max(suppressRealtimeUntil,Date.now()+ms);
}
function safePhotoSrc(p){
  const v=String(p||'').trim();
  if(!v||v.startsWith('sample-images/')) return '';
  if(/^data:image\/(?:png|jpe?g|webp|gif);base64,[a-z0-9+/=\s]+$/i.test(v)) return v.replace(/\s/g,'');
  try{
    const u=new URL(v, location.href);
    if(u.protocol==='https:'||u.origin===location.origin) return u.href;
  }catch(e){}
  return '';
}
function photoOk(p){ return !!safePhotoSrc(p); }
function loadScript(src, globalName){
  if(globalName&&window[globalName]) return Promise.resolve();
  if(scriptLoads[src]) return scriptLoads[src];
  scriptLoads[src]=new Promise((resolve,reject)=>{
    const s=document.createElement('script');
    s.src=src;
    s.async=true;
    s.onload=()=>{
      if(globalName&&!window[globalName]){
        delete scriptLoads[src];
        reject(new Error(globalName+' 載入失敗'));
        return;
      }
      resolve();
    };
    s.onerror=()=>{
      delete scriptLoads[src];
      reject(new Error(src+' 載入失敗'));
    };
    document.head.appendChild(s);
  });
  return scriptLoads[src];
}
function ensureQRCode(){ return loadScript(LIBS.qr,'QRCode'); }
function ensureXLSX(){ return loadScript(LIBS.xlsx,'XLSX'); }
function getStatus(item){
  if(item.status&&item.status.includes('Pending')) return 'pending';
  if(item.status&&(item.status.includes('耗盡')||item.status.includes('Empty'))) return 'empty';
  return 'in';
}
function statusChip(item){
  const s=getStatus(item);
  if(s==='pending') return '<span class="chip chip-pending">待收料</span>';
  if(s==='empty')   return '<span class="chip chip-empty">已耗盡</span>';
  return '<span class="chip chip-in">在庫</span>';
}
function catChip(cat){ return `<span class="chip chip-${cat}">${CAT_LABELS[cat]||cat}</span>`; }
function thumbCell(item, canClick=true){
  const src=safePhotoSrc(item.photo);
  if(src){
    const click = canClick ? 'onclick="showImgPreview(this.src)"' : '';
    return `<img class="thumb" src="${esc(src)}" alt="照片" ${click}>`;
  }
  return `<div class="thumb-empty" title="無照片">📷</div>`;
}
function getItemCat(item){ return item.categoryCode||(item.code||'').charAt(0); }
function getOutsoleSize(item){
  if(!item.sizeEU) return '';
  return `EU ${item.sizeEU} / UK ${item.sizeUK||'-'} / US ${item.sizeUS||'-'} / JP ${item.sizeJP||'-'}`;
}
function searchText(item){
  return [item.productName||'',item.code||'',item.catName||'',item.typeName||'',
          item.specName||'',item.vendor||'',item.brand||'',item.colorCode||'',item.colorPantone||'',
          item.locationCode||'',item.locationName||'',item.thickness||'',
          item.supplier?.name||''].join(' ').toLowerCase();
}
function filterItems(q, cat, type, status, arr){
  return (arr||inventory).filter(item=>{
    if(q && !searchText(item).includes(q.toLowerCase())) return false;
    if(cat && getItemCat(item)!==cat) return false;
    if(type && item.typeName!==type) return false;
    if(status && getStatus(item)!==status) return false;
    return true;
  });
}

// ─── Cloud / Auth ────────────────────────────────────────────────────
function getSupaCfg(){
  const c=window.PEIHAI_SUPABASE_CONFIG||{};
  const url=(c.url||'').trim(), key=(c.anonKey||'').trim();
  if(!url||!key||url.includes('YOUR_')) return null;
  return {url,anonKey:key,allowSignUp:c.allowSignUp===true};
}
function canEdit(){ if(!cloudEnabled) return true; if(!currentUser) return false; return currentRole==='admin'||currentRole==='staff'||currentRole==='local'; }
function isAdmin(){ if(!cloudEnabled) return true; if(!currentUser) return false; return currentRole==='admin'||currentRole==='local'; }
function setCloud(msg, state){
  const b=document.getElementById('cloudBadge');
  const t=document.getElementById('cloudText');
  if(t) t.textContent=msg;
  if(b){ b.classList.remove('cloud-ok','cloud-err'); if(state) b.classList.add(state); }
}
function setAuthMsg(m){ const el=document.getElementById('authMsg'); if(el) el.textContent=m||''; }
function isUsefulAuthErrorText(v){
  const text=String(v||'').trim();
  return text!==''&&text!=='{}'&&text!=='[]'&&text!=='[object Object]';
}
function describeAuthError(error){
  const parts=[
    error?.message,
    error?.error_description,
    error?.msg,
    error?.code,
    error?.error,
    error?.details,
    error?.hint,
    error?.reason,
    error?.cause?.message
  ].filter(isUsefulAuthErrorText).map(v=>String(v).trim());
  const text=parts.join(' / ');
  if(/invalid login credentials|invalid_credentials/i.test(text)) return 'Email 或密碼不正確，請重新確認。';
  if(/email not confirmed|email_not_confirmed/i.test(text)) return '此帳號信箱尚未驗證，請先完成信箱驗證。';
  if(/too many requests|rate limit/i.test(text)) return '登入嘗試太頻繁，請稍後再試。';
  if(/failed to fetch|network|fetch/i.test(text)) return '連線 Supabase 失敗，請檢查網路或稍後再試。';
  if(text) return text;
  if(error?.status) return '帳號不存在、密碼不正確，或此帳號尚未被管理員加入權限名單。請也確認 Email 拼字，例如 warehouse / wearhouse 是否打錯。';
  return 'Supabase 沒有回傳詳細錯誤，請確認帳號/密碼，或重新整理後再試。';
}
function updateAuthUI(){
  const userBadge=document.getElementById('userBadge');
  const logoutBtn=document.getElementById('logoutBtn');
  const clearBtn=document.getElementById('clearAllBtn');
  if(cloudEnabled && !currentUser){
    document.body.classList.add('auth-required');
    setCloud('請登入','cloud-err');
    if(logoutBtn) logoutBtn.style.display='none';
    if(userBadge) userBadge.style.display='none';
    return;
  }
  document.body.classList.remove('auth-required');
  if(currentUser){
    setCloud('雲端同步中','cloud-ok');
    if(userBadge){
      userBadge.style.display='inline-flex';
      userBadge.textContent=(currentUser.email||'user')+' · '+currentRole;
    }
    if(logoutBtn){
      logoutBtn.style.display='inline-flex';
      logoutBtn.textContent='登出';
    }
  } else {
    setCloud('本機模式','');
    if(userBadge) userBadge.style.display='none';
    if(logoutBtn) logoutBtn.style.display='none';
  }
  applyRoleUI();
}
function applyRoleUI(){
  const canE=canEdit(), isAd=isAdmin();
  ['navManage','navBatch'].forEach(id=>{ const el=document.getElementById(id); if(el) el.style.display=canE?'':'none'; });
  const cb=document.getElementById('clearAllBtn');
  if(cb) cb.style.display=isAd?'':'none';
}
async function initSupabase(){
  try{
    const cfg=getSupaCfg();
    if(!cfg){ cloudEnabled=false; updateAuthUI(); return; }
    if(!window.supabase){
      cloudEnabled=false;
      updateAuthUI();
      setCloud('Supabase 載入失敗','cloud-err');
      setAuthMsg('Supabase 程式庫尚未載入，請檢查網路後重新整理。');
      return;
    }
    cloudEnabled=true;
    supabaseClient=window.supabase.createClient(cfg.url,cfg.anonKey);
    const rb=document.getElementById('registerBtn');
    if(rb) rb.style.display=cfg.allowSignUp?'':'none';
    const {data,error}=await supabaseClient.auth.getSession();
    if(error) console.warn('Supabase session check failed',error);
    await applySession(data?.session||null);
    supabaseClient.auth.onAuthStateChange((_,s)=>applySession(s||null));
  }catch(error){
    console.error('Supabase init failed',error);
    cloudEnabled=false;
    updateAuthUI();
    setCloud('Supabase 初始化失敗','cloud-err');
    setAuthMsg('Supabase 初始化失敗：'+describeAuthError(error));
  }
}
async function applySession(session){
  if(!cloudEnabled) return;
  currentUser=session?.user||null;
  currentRole=currentUser?'viewer':'local';
  if(!currentUser){ if(realtimeChannel){supabaseClient.removeChannel(realtimeChannel);realtimeChannel=null;} updateAuthUI(); return; }
  await loadProfile();
  updateAuthUI();
  await loadCloud();
  // Auto-load Excel data if cloud is empty or has old IMP codes
  const hasOldCloud=inventory.some(i=>(i.code||'').startsWith('IMP-'));
  if(inventory.length===0 || hasOldCloud){
    try{
      const excelItems=EXCEL_IMPORT_DATA;
      inventory=excelItems;
      if(canEdit()){ await saveAllCloud(); }
      inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
      renderAll();
    }catch(e){ console.error(e); renderAll(); }
  }
  subscribeCloud();
}
async function loadProfile(){
  if(!supabaseClient||!currentUser) return;
  const {data,error}=await supabaseClient.from('profiles').select('role').eq('id',currentUser.id).maybeSingle();
  if(!error && data?.role){ currentRole=data.role; return; }
  await supabaseClient.from('profiles').insert({id:currentUser.id,email:currentUser.email,role:'viewer'});
  currentRole='viewer';
}
async function loginUser(){
  if(!supabaseClient){
    await initSupabase();
    if(!supabaseClient){ setAuthMsg('Supabase 尚未設定或尚未載入，請重新整理後再試。'); return; }
  }
  const email=document.getElementById('authEmail').value.trim();
  const pwd=document.getElementById('authPassword').value;
  if(!email||!pwd){ setAuthMsg('請輸入 Email 與密碼'); return; }
  setAuthMsg('登入中…');
  try{
    const {data,error}=await supabaseClient.auth.signInWithPassword({email,password:pwd});
    if(error){
      console.warn('Login failed',error);
      setAuthMsg('登入失敗：'+describeAuthError(error));
      return;
    }
    if(!data?.session&&!data?.user){
      setAuthMsg('登入失敗：Supabase 沒有回傳登入階段，請稍後再試。');
      return;
    }
    setAuthMsg('登入成功，同步資料中…');
  }catch(error){
    console.error('Login crashed',error);
    setAuthMsg('登入失敗：'+describeAuthError(error));
  }
}
async function registerUser(){
  if(!supabaseClient) return;
  const email=document.getElementById('authEmail').value.trim();
  const pwd=document.getElementById('authPassword').value;
  if(!email||!pwd){ setAuthMsg('請輸入 Email 與密碼'); return; }
  if(pwd.length<6){ setAuthMsg('密碼至少 6 碼'); return; }
  setAuthMsg('建立中…');
  try{
    const {error}=await supabaseClient.auth.signUp({email,password:pwd});
    if(error){ setAuthMsg('建立失敗：'+describeAuthError(error)); return; }
    setAuthMsg('帳號已建立，請確認信箱後登入。');
  }catch(error){
    console.error('Register crashed',error);
    setAuthMsg('建立失敗：'+describeAuthError(error));
  }
}
async function logoutUser(){
  if(supabaseClient) await supabaseClient.auth.signOut();
  currentUser=null; currentRole='local'; updateAuthUI();
}
async function loadCloud(showAlert){
  if(!supabaseClient||!currentUser) return;
  if(cloudLoadTimer){ clearTimeout(cloudLoadTimer); cloudLoadTimer=null; }
  setCloud('雲端同步中','cloud-ok');
  const {data,error}=await supabaseClient.from('materials').select('code,data,updated_at').order('updated_at',{ascending:false});
  if(error){ setCloud('雲端讀取失敗','cloud-err'); if(showAlert) alert('雲端讀取失敗'); return; }
  inventory=(data||[]).map(r=>r.data).filter(Boolean);
  inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
  if(db) replaceInDB(inventory,()=>renderAll(),false);
  else renderAll();
  setCloud('雲端同步完成','cloud-ok');
}
function scheduleCloudReload(){
  if(Date.now()<suppressRealtimeUntil) return;
  if(cloudLoadTimer) clearTimeout(cloudLoadTimer);
  cloudLoadTimer=setTimeout(()=>loadCloud(false),700);
}
function subscribeCloud(){
  if(!supabaseClient||!currentUser||realtimeChannel) return;
  realtimeChannel=supabaseClient.channel('mat-rt')
    .on('postgres_changes',{event:'*',schema:'public',table:'materials'},()=>scheduleCloudReload())
    .subscribe();
}
function queueCloudSave(item){
  if(!cloudEnabled||!supabaseClient||!currentUser||suppressCloudSave||!item?.code) return true;
  pendingCloudItems.set(item.code,item);
  if(cloudSaveTimer) clearTimeout(cloudSaveTimer);
  cloudSaveTimer=setTimeout(()=>flushCloudSaves(),900);
  setCloud('雲端同步排程中','cloud-ok');
  return true;
}
async function flushCloudSaves(){
  if(cloudSaveTimer){ clearTimeout(cloudSaveTimer); cloudSaveTimer=null; }
  if(!pendingCloudItems.size) return true;
  const items=Array.from(pendingCloudItems.values());
  pendingCloudItems.clear();
  return saveItemsCloud(items);
}
async function saveItemsCloud(items){
  if(!cloudEnabled||!supabaseClient||!currentUser||suppressCloudSave) return true;
  const clean=(Array.isArray(items)?items:[items]).filter(item=>item&&item.code);
  if(!clean.length) return true;
  // permission checked by Supabase RLS
  setCloud('雲端同步中','cloud-ok');
  touchCloudSyncWindow(Math.max(5000,clean.length*120));
  const rows=clean.map(item=>({code:item.code,data:item,updated_by:currentUser.id,updated_at:new Date().toISOString()}));
  const {error}=await supabaseClient.from('materials').upsert(rows,{onConflict:'code'});
  if(error){
    clean.forEach(item=>pendingCloudItems.set(item.code,item));
    setCloud('雲端儲存失敗','cloud-err');
    return false;
  }
  setCloud('雲端同步完成','cloud-ok');
  return true;
}
async function saveItemCloud(item, immediate=false){
  if(immediate) return saveItemsCloud([item]);
  return queueCloudSave(item);
}
async function saveAllCloud(){
  if(!cloudEnabled||!supabaseClient||!currentUser||!isAdmin()) return;
  if(cloudSaveTimer){ clearTimeout(cloudSaveTimer); cloudSaveTimer=null; }
  pendingCloudItems.clear();
  const rows=inventory.map(item=>({code:item.code,data:item,updated_by:currentUser.id,updated_at:new Date().toISOString()}));
  if(!rows.length) return;
  setCloud('雲端整批同步中','cloud-ok');
  touchCloudSyncWindow(Math.max(8000,rows.length*160));
  const chunkSize=100;
  for(let i=0;i<rows.length;i+=chunkSize){
    const {error}=await supabaseClient.from('materials').upsert(rows.slice(i,i+chunkSize),{onConflict:'code'});
    if(error){ setCloud('雲端儲存失敗','cloud-err'); alert('匯入雲端失敗'); return; }
  }
  setCloud('雲端同步完成','cloud-ok');
}
async function deleteMaterialsCloud(codes){
  if(!cloudEnabled||!supabaseClient||!currentUser) return true;
  const list=Array.isArray(codes)?codes.filter(Boolean):[codes].filter(Boolean);
  if(!list.length) return true;
  touchCloudSyncWindow(Math.max(5000,list.length*120));
  const {error}=await supabaseClient.from('materials').delete().in('code',list);
  if(error){
    setCloud('雲端刪除失敗','cloud-err');
    alert('雲端刪除失敗：'+error.message);
    return false;
  }
  return true;
}

// ─── IndexedDB ──────────────────────────────────────────────────────
function initDB(){
  const req=indexedDB.open(DB_NAME,1);
  req.onupgradeneeded=e=>{ db=e.target.result; if(!db.objectStoreNames.contains('inventory')) db.createObjectStore('inventory',{keyPath:'code'}); };
  req.onsuccess=e=>{ db=e.target.result; loadFromDB(); };
  req.onerror=()=>alert('資料庫啟動失敗');
}
function loadFromDB(){
  try{
    const tx=db.transaction(['inventory'],'readonly');
    tx.objectStore('inventory').getAll().onsuccess=e=>{
      if(getSupaCfg()){
        setCloud('雲端登入確認中','cloud-ok');
        return;
      }
      if(cloudEnabled&&currentUser) return;
      inventory=e.target.result||[];
      // Force reload if empty OR if old IMP- codes detected
      const hasOldCodes = inventory.some(i=>(i.code||'').startsWith('IMP-'));
      if(inventory.length===0 || hasOldCodes){
        autoLoadExcelData();
      } else {
        inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
        renderAll();
      }
    };
  }catch(e){ console.error(e); }
}
function autoLoadExcelData(){
  try{
    const excelItems = EXCEL_IMPORT_DATA;
    inventory = excelItems;
    inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
    replaceInDB(inventory, ()=>{ renderAll(); }, false);
    localStorage.setItem('peihai_ex','1');
  }catch(e){ console.error('Auto load failed',e); renderAll(); }
}
function saveItemDB(item){
  try{
    db.transaction(['inventory'],'readwrite').objectStore('inventory').put(item);
    saveItemCloud(item);
  }catch(e){ console.error(e); }
}
function deleteItemDB(code){
  try{ db.transaction(['inventory'],'readwrite').objectStore('inventory').delete(code); }catch(e){}
}
function replaceInDB(items, onDone, syncCloud=true){
  try{
    suppressCloudSave=!syncCloud;
    const tx=db.transaction(['inventory'],'readwrite');
    const st=tx.objectStore('inventory');
    st.clear();
    items.forEach(item=>st.put(item));
    tx.oncomplete=()=>{
      suppressCloudSave=false;
      if(syncCloud&&cloudEnabled&&isAdmin()) saveAllCloud();
      if(onDone) onDone();
    };
    tx.onerror=()=>{ suppressCloudSave=false; alert('寫入資料庫失敗'); };
  }catch(e){ suppressCloudSave=false; alert('資料庫未準備好，請重新整理後再試'); }
}

// ─── Brands ─────────────────────────────────────────────────────────
function initBrands(){
  const saved=localStorage.getItem('peihai_brands');
  customBrands=saved?JSON.parse(saved):['FitFlop','Onitsuka Tiger','Alpinestars','Trail Pro','Urban Flex','通用開發 General'];
  localStorage.setItem('peihai_brands',JSON.stringify(customBrands));
  renderBrands();
}
function renderBrands(){
  const sel=document.getElementById('fBrand');
  if(!sel) return;
  sel.innerHTML='<option value="">選擇…</option>'+customBrands.map(b=>`<option>${esc(b)}</option>`).join('')+'<option value="_NEW_">＋ 新增品牌…</option>';
}
function handleBrandChange(){
  const sel=document.getElementById('fBrand');
  if(sel.value!=='_NEW_') return;
  const n=(prompt('輸入新品牌名稱：')||'').trim();
  if(n&&!customBrands.includes(n)){ customBrands.push(n); localStorage.setItem('peihai_brands',JSON.stringify(customBrands)); }
  renderBrands();
  if(n) sel.value=n; else sel.value='';
}

// ─── Outsole size selects ────────────────────────────────────────────
function buildOutsoleSels(){
  const eu=document.getElementById('sEU');
  if(!eu||eu.options.length>1) return;
  eu.innerHTML='<option value="">EU…</option>';
  for(let i=20;i<=50;i++) eu.innerHTML+=`<option value="${i}">${i}</option>`;
  const vals=k=>Array.from(new Set(Object.values(OUTSOLE_SIZES).map(s=>s[k])));
  const fill=(id,label,arr)=>{ const s=document.getElementById(id); s.innerHTML=`<option value="">${label}…</option>`+arr.map(v=>`<option>${v}</option>`).join(''); };
  fill('sUK','UK',vals('uk')); fill('sUS','US',vals('us')); fill('sJP','JP',vals('jp'));
}
function syncOutsoleSizes(){
  const eu=document.getElementById('sEU').value;
  const m=OUTSOLE_SIZES[eu]; if(!m) return;
  document.getElementById('sUK').value=m.uk;
  document.getElementById('sUS').value=m.us;
  document.getElementById('sJP').value=m.jp;
}

// ─── Form logic ──────────────────────────────────────────────────────
function onCatChange(){
  const cat=document.getElementById('fCat').value;
  const show=(id,v)=>{ const el=document.getElementById(id); if(el) el.style.display=v?'':'none'; };
  show('fTypeGroup',false); show('fSpecDropGroup',false); show('fSpecNumGroup',false);
  show('fHardnessGroup',false); show('fOutsoleSizeGroup',false);
  if(!cat){ buildSKUPreview(); return; }
  if(cat==='O'){
    buildOutsoleSels();
    show('fOutsoleSizeGroup',true);
    document.getElementById('fUnit').value='PRS';
  } else {
    show('fTypeGroup',true);
    const ts=document.getElementById('fType');
    ts.innerHTML='<option value="">選擇…</option>';
    const types=MAT[cat]?.types||{};
    Object.entries(types).forEach(([k,v])=>{ ts.innerHTML+=`<option value="${k}">${k} — ${v}</option>`; });
    if(['L','S','T'].includes(cat)){
      show('fSpecDropGroup',true);
      const ss=document.getElementById('fSpecDrop');
      ss.innerHTML='<option value="">選擇…</option>';
      Object.entries(MAT[cat].specs||{}).forEach(([k,v])=>{ ss.innerHTML+=`<option value="${k}">${k} — ${v}</option>`; });
    } else {
      show('fSpecNumGroup',true);
      const lbl={W:'寬度 Width (mm)',A:'厚度 Thickness',H:'尺寸 Size (mm)'}[cat]||'規格';
      document.getElementById('fSpecNumLabel').textContent=lbl+' *';
    }
    if(cat==='L') document.getElementById('fUnit').value='SF';
    else if(['S','T'].includes(cat)) document.getElementById('fUnit').value='YD';
    else if(cat==='H') document.getElementById('fUnit').value='PCS';
  }
  buildSKUPreview();
}
function onTypeChange(){
  const cat=document.getElementById('fCat').value;
  const type=document.getElementById('fType').value;
  const show=(id,v)=>{ const el=document.getElementById(id); if(el) el.style.display=v?'':'none'; };
  if(cat==='A' && (type==='E'||type==='H')) show('fHardnessGroup',true);
  else show('fHardnessGroup',false);
}
function buildSKUPreview(){
  const cat=document.getElementById('fCat').value;
  const color=(document.getElementById('fColor').value||'').toUpperCase().trim();
  if(!cat){ document.getElementById('skuPreview').textContent='—'; return; }
  if(cat==='O'){
    const name=(document.getElementById('fName').value||'').trim();
    document.getElementById('skuPreview').textContent=name||'—';
    return;
  }
  const type=document.getElementById('fType').value||'?';
  let spec='?';
  if(['L','S','T'].includes(cat)) spec=document.getElementById('fSpecDrop')?.value||'?';
  else { const v=document.getElementById('fSpecNum')?.value; spec=v?String(parseInt(v)).padStart(2,'0'):'??'; }
  const prefix=cat+type+spec;
  const seq=nextSeqForPrefix(prefix);
  const colorCodes=splitBulkValues(color);
  const firstCode=colorCodes[0]||'???';
  const colorCount=colorCodes.length;
  document.getElementById('skuPreview').textContent=`${prefix}-${seq}-${firstCode}`+(colorCount>1?' (+其他 '+(colorCount-1)+' 色)':'');
}

// ─── Photo upload ─────────────────────────────────────────────────────
function readFileAsDataURL(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>resolve(e.target.result);
    reader.onerror=()=>reject(new Error('圖片讀取失敗'));
    reader.readAsDataURL(file);
  });
}
function loadImageElement(src){
  return new Promise((resolve,reject)=>{
    const img=new Image();
    img.onload=()=>resolve(img);
    img.onerror=()=>reject(new Error('圖片載入失敗'));
    img.src=src;
  });
}
const PHOTO_MAX_SIDE=800;
const PHOTO_TARGET_BYTES=180000;
const PHOTO_MIN_SAVING_BYTES=2048;
const PHOTO_SYNC_CHUNK=10;
const PHOTO_QUALITY_STEPS=[0.62,0.56,0.50,0.44,0.38,0.32];
function dataUrlBytes(dataUrl){
  const body=String(dataUrl||'').split(',')[1]||'';
  return Math.ceil(body.replace(/\s/g,'').length*3/4);
}
function formatPhotoBytes(bytes){
  if(bytes>=1048576) return (bytes/1048576).toFixed(1)+' MB';
  return Math.max(1,Math.round(bytes/1024))+' KB';
}
function isStoredDataPhoto(src){
  return /^data:image\/(?:png|jpe?g|webp|gif);base64,/i.test(String(src||'').trim());
}
function isDataPhoto(src){
  return /^data:image\/(?!svg\+xml)[a-z0-9.+-]+;base64,/i.test(String(src||'').trim());
}
function drawPhotoToCanvas(img){
  const width=img.naturalWidth||img.width||1;
  const height=img.naturalHeight||img.height||1;
  const scale=Math.min(1,PHOTO_MAX_SIDE/Math.max(width,height));
  const canvas=document.createElement('canvas');
  canvas.width=Math.max(1,Math.round(width*scale));
  canvas.height=Math.max(1,Math.round(height*scale));
  const ctx=canvas.getContext('2d');
  if(!ctx) throw new Error('圖片處理失敗');
  ctx.fillStyle='#fff';
  ctx.fillRect(0,0,canvas.width,canvas.height);
  ctx.drawImage(img,0,0,canvas.width,canvas.height);
  return canvas;
}
async function compressPhotoDataUrl(src){
  const safeSrc=String(src||'').trim().replace(/\s/g,'');
  if(!isDataPhoto(safeSrc)) throw new Error('圖片格式不支援壓縮');
  const originalBytes=dataUrlBytes(safeSrc);
  const img=await loadImageElement(safeSrc);
  const width=img.naturalWidth||img.width||1;
  const height=img.naturalHeight||img.height||1;
  if(Math.max(width,height)<=PHOTO_MAX_SIDE && originalBytes<=PHOTO_TARGET_BYTES) return safeSrc;
  const canvas=drawPhotoToCanvas(img);
  let best='';
  let bestBytes=Infinity;
  for(const quality of PHOTO_QUALITY_STEPS){
    const next=canvas.toDataURL('image/jpeg',quality);
    const nextBytes=dataUrlBytes(next);
    if(nextBytes<bestBytes){ best=next; bestBytes=nextBytes; }
    if(nextBytes<=PHOTO_TARGET_BYTES) return next;
  }
  return best||safeSrc;
}
async function compressPhoto(file){
  if(!file.type.startsWith('image/')) throw new Error('請選擇圖片檔');
  if(file.type==='image/svg+xml') throw new Error('不支援 SVG 圖片');
  const original=await readFileAsDataURL(file);
  return compressPhotoDataUrl(original);
}
async function handlePhotoUpload(event, previewId, hiddenId){
  const file=event.target.files[0]; if(!file) return;
  try{
    const dataUrl=await compressPhoto(file);
    const img=document.getElementById(previewId);
    const hid=document.getElementById(hiddenId);
    if(img){ img.src=dataUrl; img.style.display='block'; }
    if(hid) hid.value=dataUrl;
  }catch(e){
    alert(e.message||'圖片處理失敗，請換一張照片再試');
  }
}
async function syncCompressedPhotoItems(changedItems){
  if(!cloudEnabled||!supabaseClient||!currentUser||!isAdmin()) return true;
  setCloud('照片同步中','cloud-ok');
  for(let i=0;i<changedItems.length;i+=PHOTO_SYNC_CHUNK){
    const ok=await saveItemsCloud(changedItems.slice(i,i+PHOTO_SYNC_CHUNK));
    if(!ok) return false;
  }
  setCloud('照片同步完成','cloud-ok');
  return true;
}
async function compressExistingPhotos(){
  if(!isAdmin()){ alert('只有管理員可以壓縮照片'); return; }
  const candidates=inventory.filter(item=>isStoredDataPhoto(safePhotoSrc(item.photo)));
  if(!candidates.length){ alert('目前沒有需要壓縮的照片'); return; }
  if(!confirm('會壓縮已儲存的 '+candidates.length+' 張照片，壓縮後同步到雲端。要開始嗎？')) return;
  await flushCloudSaves();
  setStorageStatus('照片壓縮中…');
  const changedItems=[];
  let beforeBytes=0, afterBytes=0, failed=0;
  for(const item of candidates){
    const original=safePhotoSrc(item.photo);
    const originalBytes=dataUrlBytes(original);
    beforeBytes+=originalBytes;
    try{
      const compressed=await compressPhotoDataUrl(original);
      const compressedBytes=dataUrlBytes(compressed);
      if(compressedBytes+PHOTO_MIN_SAVING_BYTES<originalBytes){
        item.photo=compressed;
        changedItems.push(item);
        afterBytes+=compressedBytes;
      }else{
        afterBytes+=originalBytes;
      }
    }catch(e){
      failed++;
      afterBytes+=originalBytes;
      console.warn('Photo compression skipped',item.code,e);
    }
  }
  if(!changedItems.length){
    setStorageStatus(failed?'照片壓縮完成，部分照片略過':'照片已經夠小');
    alert(failed?'沒有照片變更，'+failed+' 張照片無法壓縮。':'照片已經夠小，不需要再壓縮。');
    return;
  }
  const savedBytes=Math.max(0,beforeBytes-afterBytes);
  const finish=async()=>{
    renderAll();
    const cloudOk=await syncCompressedPhotoItems(changedItems);
    const msg='已壓縮 '+changedItems.length+' 張照片，約減少 '+formatPhotoBytes(savedBytes)+(failed?'，'+failed+' 張略過。':'。');
    setStorageStatus(msg);
    alert(cloudOk?msg:msg+' 但雲端同步失敗，稍後請再試一次。');
  };
  if(db) replaceInDB(inventory,finish,false);
  else finish();
}
function showImgPreview(src){
  const safeSrc=safePhotoSrc(src);
  if(!safeSrc) return;
  document.getElementById('imgPreviewImg').src=safeSrc;
  document.getElementById('imgPreviewBackdrop').classList.add('open');
}
function closeImgPreview(){ document.getElementById('imgPreviewBackdrop').classList.remove('open'); }

// ─── Camera scanner ───────────────────────────────────────────────────
async function startCamera(targetId){
  if(!('BarcodeDetector' in window)){ alert('此瀏覽器不支援相機掃碼，請用掃碼槍或手動輸入'); return; }
  try{
    scannerTargetId=targetId; scannerActive=true;
    document.getElementById('cameraModal').classList.add('open');
    const video=document.getElementById('scannerVideo');
    scannerStream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}});
    video.srcObject=scannerStream;
    await video.play();
    scanFrame();
  }catch(e){ stopCamera(); alert('無法開啟相機，請確認權限'); }
}
async function scanFrame(){
  if(!scannerActive) return;
  const video=document.getElementById('scannerVideo');
  if(video.readyState>=2){
    try{
      const codes=await new BarcodeDetector({formats:['qr_code']}).detect(video);
      if(codes.length>0){
        const val=(codes[0].rawValue||'').trim();
        const target=document.getElementById(scannerTargetId);
        if(target&&val){ target.value=val.toUpperCase(); stopCamera(); if(scannerTargetId==='batchInput') addToBatch(); return; }
      }
    }catch(e){ stopCamera(); return; }
  }
  requestAnimationFrame(scanFrame);
}
function stopCamera(){
  scannerActive=false;
  document.getElementById('cameraModal').classList.remove('open');
  const v=document.getElementById('scannerVideo');
  if(v) v.srcObject=null;
  if(scannerStream){ scannerStream.getTracks().forEach(t=>t.stop()); scannerStream=null; }
}

// ─── Create material ──────────────────────────────────────────────────
function createMaterial(){
  // permission checked by Supabase RLS
  // Clear previous errors
  document.querySelectorAll('#pageManage input, #pageManage select, #pageManage textarea').forEach(el=>el.classList.remove('field-err'));

  const cat=document.getElementById('fCat').value;
  const color=document.getElementById('fColor').value.toUpperCase().trim();
  const name=document.getElementById('fName').value.trim();
  const vendor=document.getElementById('fVendor').value.trim();
  const brand=document.getElementById('fBrand').value;
  const unit=document.getElementById('fUnit').value;
  const price=document.getElementById('fPrice').value;
  const currency=document.getElementById('fCurrency').value;
  const photo=document.getElementById('photoBase64').value;
  const thickness=document.getElementById('fThickness').value.trim();
  const contact=document.getElementById('fContact').value.trim();
  const phone=document.getElementById('fPhone').value.trim();
  const origin=document.getElementById('fOrigin').value.trim();
  const lead='0';
  const moq=document.getElementById('fMoq').value.trim();

  const err=(id)=>{ const el=document.getElementById(id); if(el) el.classList.add('field-err'); };
  let valid=true;
  if(!cat){ err('fCat'); valid=false; }
  if(!name){ err('fName'); valid=false; }
  if(!color){ err('fColor'); valid=false; }
  // Parse colors from two fields (comma, semicolon, or newline separated)
  const colorCodes=splitBulkValues(color);
  const colorNameField=(document.getElementById('fColorName').value||'');
  const colorNames=splitBulkValues(colorNameField);
  const colors=colorCodes.map((code,i)=>({code:code, name:colorNames[i]||code}));
  if(colors.length===0){ err('fColor'); valid=false; }
  if(!vendor){ err('fVendor'); valid=false; }
  if(!brand||brand==='_NEW_'){ err('fBrand'); valid=false; }

  let type='', specCode='', specName='', finalCode='';
  const sizeEU=document.getElementById('sEU')?.value||'';

  if(cat!=='O'){
    type=document.getElementById('fType').value;
    if(!type){ err('fType'); valid=false; }
  }

  if(cat==='O'){
    if(!sizeEU){ err('sEU'); valid=false; }
    specName=`大底尺寸 EU ${sizeEU} / UK ${document.getElementById('sUK').value} / US ${document.getElementById('sUS').value} / JP ${document.getElementById('sJP').value}`;
  } else if(['L','S','T'].includes(cat)){
    specCode=document.getElementById('fSpecDrop').value;
    if(!specCode){ err('fSpecDrop'); valid=false; }
    else specName=MAT[cat].specs[specCode]||specCode;
  } else if(['W','A','H'].includes(cat)){
    const sv=document.getElementById('fSpecNum').value;
    if(!sv||sv<=0){ err('fSpecNum'); valid=false; }
    else{
      specCode=String(parseInt(sv)).padStart(2,'0').substring(0,2);
      specName=cat==='A'?(sv/10).toFixed(1)+' mm':sv+' mm';
      if(cat==='A'&&(type==='E'||type==='H')){
        const hd=document.getElementById('fHardness').value;
        if(!hd||hd<=0){ err('fHardness'); valid=false; }
        else specName+=` (硬度 ${hd})`;
      }
    }
  }

  if(!valid){ alert('請填寫必填欄位'); return; }

  const catMap={L:'皮類 Leather',S:'人造皮 Synthetic',T:'布類 Textile',W:'條狀類 Webbing',A:'副料 Auxiliary',O:'OUTSOLE 大底',H:'五金 Hardware'};
  const createdCodes=[];
  colors.forEach((clr,ci)=>{
    let finalCode;
    if(cat==='O'){
      finalCode=ci===0?name:name+'-'+clr.code;
      if(inventory.some(i=>i.code&&i.code.toLowerCase()===finalCode.toLowerCase())) return;
    } else {
      const prefix=cat+type+specCode;
      finalCode=nextMaterialCode(prefix,clr.code,createdCodes);
    }
    const item={
      code:finalCode, categoryCode:cat, catName:catMap[cat]||cat,
      typeName:cat==='O'?'大底 Outsole':(MAT[cat]?.types[type]||type),
      specName, colorCode:clr.name, colorPantone:clr.code, currency, price:price||'0',
      sizeEU:cat==='O'?sizeEU:'', sizeUK:cat==='O'?document.getElementById('sUK').value:'',
      sizeUS:cat==='O'?document.getElementById('sUS').value:'',
      sizeJP:cat==='O'?document.getElementById('sJP').value:'',
      locationCode:'', locationName:'', qty:'0.0', unit,
      photo:ci===0?photo:'', productName:name, vendor, brand, thickness,
      supplier:{name:vendor, contact, phone, origin, lead:0, moq},
      date:'', status:'待到貨 Pending',
      timestamp:Date.now()-ci,
      logs:[{time:now(), action:'建檔 Created', amount:'+0', reason:'建立材料 '+clr.code+' / '+clr.name+(colors.length>1?' (批量建色 '+(ci+1)+'/'+colors.length+')':''), balance:'0', type:'create'}]
    };
    inventory.unshift(item);
    saveItemDB(item);
    createdCodes.push(finalCode);
  });
  flushCloudSaves();
  renderAll();
  ['fColor','fColorName','fName','photoBase64','fThickness','fContact','fPhone','fOrigin','fMoq','fPrice'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('photoPreview').style.display='none';
  document.getElementById('photoInput').value='';
  document.getElementById('skuPreview').textContent='—';
  if(cat==='O'){ ['sEU','sUK','sUS','sJP'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; }); }
  document.getElementById('createBackdrop').classList.remove('open');
  if(createdCodes.length>0) alert('✅ 建檔成功！共 '+createdCodes.length+' 筆\n\n'+createdCodes.join('\n')+'\n\n到貨後點入庫按鈕');
}

// ─── Inbound ──────────────────────────────────────────────────────────
function openInbound(code){
  // permission checked by Supabase RLS
  const item=inventory.find(i=>i.code===code); if(!item) return;
  document.getElementById('inboundTitle').innerHTML=`<span class="alert-icon">📦</span>${esc(item.productName||'')} — <code>${esc(item.code)}</code>`;
  document.getElementById('ibDate').valueAsDate=new Date();
  document.getElementById('ibQty').value='';
  document.getElementById('ibItemCode').value=code;
  document.getElementById('ibPhotoPreview').style.display='none';
  document.getElementById('ibPhotoBase64').value='';
  // Smart warehouse default
  if(item.locationCode){
    const p=item.locationCode.split('-');
    if(p.length>=2){ document.getElementById('ibWH').value=p[0]; document.getElementById('ibRack').value=p[1]; }
  } else { document.getElementById('ibWH').value=''; document.getElementById('ibRack').value=''; }
  document.getElementById('inboundBackdrop').classList.add('open');
}
function closeInbound(){ document.getElementById('inboundBackdrop').classList.remove('open'); }
function submitInbound(){
  const code=document.getElementById('ibItemCode').value;
  const item=inventory.find(i=>i.code===code); if(!item) return;
  const qty=parseFloat(document.getElementById('ibQty').value);
  const date=document.getElementById('ibDate').value;
  const wh=document.getElementById('ibWH').value;
  const rack=document.getElementById('ibRack').value;
  const photo=document.getElementById('ibPhotoBase64').value;
  if(!qty||qty<=0){ alert('請輸入正確的入庫數量'); return; }
  if(!wh||!rack){ alert('請選擇主倉與格'); return; }
  if(!date){ alert('請選擇入庫日期'); return; }
  const newQty=(parseFloat(item.qty)+qty).toFixed(1);
  item.locationCode=`${wh}-${rack}`;
  item.locationName=`主倉 ${wh} — 第 ${rack} 格`;
  item.qty=newQty;
  item.status='在庫 In';
  if(!item.date) item.date=date;
  if(photo) item.photo=photo;
  item.logs.push({time:now(),action:'收料入庫 Receive',amount:'+'+qty.toFixed(1),reason:'倉庫點收進貨',balance:newQty,type:'in'});
  saveItemDB(item); renderAll(); closeInbound();
  alert(`✅ ${item.code} 收料入庫成功！\n入庫：${qty.toFixed(1)} ${item.unit}\n儲位：${item.locationCode}`);
}

// ─── Outbound / Return ────────────────────────────────────────────────
function openOutbound(code, type){
  // permission checked by Supabase RLS
  const item=inventory.find(i=>i.code===code); if(!item) return;
  document.getElementById('outboundTitle').textContent=type==='out'?'出庫 Outbound':'還料 Return';
  document.getElementById('outboundInfo').innerHTML=`<span class="alert-icon">${type==='out'?'📤':'↩️'}</span>${esc(item.productName||'')} — <code>${esc(code)}</code>　庫存：<strong>${item.qty} ${item.unit}</strong>`;
  document.getElementById('obQty').value='';
  document.getElementById('obReason').value='';
  document.getElementById('obItemCode').value=code;
  document.getElementById('obType').value=type;
  document.getElementById('obSubmitBtn').textContent=type==='out'?'確認出庫':'確認還料';
  document.getElementById('obSubmitBtn').className='btn btn-lg '+(type==='out'?'btn-danger':'btn-purple');
  document.getElementById('outboundBackdrop').classList.add('open');
}
function closeOutbound(){ document.getElementById('outboundBackdrop').classList.remove('open'); }
function submitOutbound(){
  const code=document.getElementById('obItemCode').value;
  const type=document.getElementById('obType').value;
  const item=inventory.find(i=>i.code===code); if(!item) return;
  const qty=parseFloat(document.getElementById('obQty').value);
  const reason=document.getElementById('obReason').value.trim();
  if(!qty||qty<=0){ alert('請輸入正確數量'); return; }
  if(!reason){ alert('請輸入原因'); return; }
  if(type==='out'&&qty>parseFloat(item.qty)){ alert(`數量不可超過現有庫存 ${item.qty} ${item.unit}`); return; }
  const newQty=type==='out'
    ?(parseFloat(item.qty)-qty).toFixed(1)
    :(parseFloat(item.qty)+qty).toFixed(1);
  item.qty=newQty;
  if(type==='out' && parseFloat(item.qty)<=0) item.status='已耗盡 Empty';
  if(type==='return' && (item.status.includes('耗盡')||item.status.includes('Empty'))) item.status='在庫 In';
  item.logs.push({time:now(),action:type==='out'?'出庫 Out':'還料 Return',amount:(type==='out'?'-':'+')+qty.toFixed(1),reason,balance:newQty,type});
  saveItemDB(item); renderAll(); closeOutbound();
}

// ─── Quick outbound ───────────────────────────────────────────────────
function submitQuickOut(){
  // permission checked by Supabase RLS
  const code=(document.getElementById('qoCode').value||'').trim().toUpperCase();
  const qty=parseFloat(document.getElementById('qoQty').value);
  const reason=(document.getElementById('qoReason').value||'').trim();
  if(!code){ alert('請輸入材料編碼'); return; }
  if(!qty||qty<=0){ alert('請輸入正確數量'); return; }
  if(!reason){ alert('請輸入出庫原因'); return; }
  const item=inventory.find(i=>(i.code||'').toUpperCase()===code);
  if(!item){ alert('找不到材料編碼：'+code); document.getElementById('qoCode').value=''; return; }
  if(getStatus(item)==='pending'){ alert('此材料尚未入庫，無法出庫'); return; }
  if(parseFloat(item.qty)<=0){ alert('此材料已耗盡'); return; }
  if(qty>parseFloat(item.qty)){ alert(`數量不可超過現有庫存 ${item.qty} ${item.unit}`); return; }
  const newQty=(parseFloat(item.qty)-qty).toFixed(1);
  item.qty=newQty;
  if(parseFloat(item.qty)<=0) item.status='已耗盡 Empty';
  item.logs.push({time:now(),action:'掃碼出庫 Scan Out',amount:'-'+qty.toFixed(1),reason,balance:newQty,type:'out'});
  saveItemDB(item); renderAll();
  document.getElementById('qoCode').value=''; document.getElementById('qoQty').value='1'; document.getElementById('qoCode').focus();
  alert(`✅ 出庫成功：${item.code}\n數量：${qty.toFixed(1)} ${item.unit}`);
}

// ─── Detail modal ─────────────────────────────────────────────────────
function openDetail(code){
  const item=inventory.find(i=>i.code===code); if(!item) return;
  currentDetailCode=code;
  document.getElementById('detailTitle').textContent=item.code;
  // Photo
  const photoEl=document.getElementById('detailPhoto');
  const photoSrc=safePhotoSrc(item.photo);
  photoEl.innerHTML='';
  if(photoSrc){
    photoEl.className='detail-main-photo';
    const img=new Image();
    img.src=photoSrc;
    img.alt='照片';
    img.style.cssText='width:100%;height:100%;object-fit:cover;border-radius:inherit;cursor:zoom-in;';
    img.onclick=()=>showImgPreview(img.src);
    photoEl.appendChild(img);
  } else {
    photoEl.className='detail-main-photo empty';
    photoEl.textContent='📷';
  }
  // QR
  const qrEl=document.getElementById('detailQR');
  qrEl.textContent='QR 載入中…';
  ensureQRCode().then(()=>{
    if(currentDetailCode!==code) return;
    qrEl.innerHTML='';
    new QRCode(qrEl,{text:item.code,width:120,height:120,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.H});
  }).catch(()=>{
    qrEl.textContent='QR 載入失敗';
  });
  // Info
  const sizeDisp=getOutsoleSize(item);
  const locDisp=item.locationCode?`${item.locationName} (${item.locationCode})`:'未入庫';
  document.getElementById('detailInfo').innerHTML=`
    <div class="di-row"><span class="di-label">材料名稱</span><span class="di-val">${esc(item.productName||'')}</span></div>
    <div class="di-row"><span class="di-label">材料編碼</span><span class="di-val" style="font-family:var(--font-mono);color:var(--accent);">${esc(item.code)}</span></div>
    <div class="di-row"><span class="di-label">大類</span><span class="di-val">${catChip(getItemCat(item))}</span></div>
    <div class="di-row"><span class="di-label">種類 / 規格</span><span class="di-val">${esc(item.typeName||'')} ${item.specName?'/ '+esc(item.specName):''}</span></div>
    ${sizeDisp?`<div class="di-row span2"><span class="di-label">尺寸</span><span class="di-val">${esc(sizeDisp)}</span></div>`:''}
    <div class="di-row"><span class="di-label">顏色</span><span class="di-val">${esc(item.colorCode||'—')}</span></div>
    <div class="di-row"><span class="di-label">厚度</span><span class="di-val">${esc(item.thickness||'—')}</span></div>
    <div class="di-row"><span class="di-label">單位</span><span class="di-val">${esc(item.unit||'')}</span></div>
    <div class="di-row"><span class="di-label">單價</span><span class="di-val" style="color:var(--danger);font-weight:600;">${esc(item.price||'0')} ${esc(item.currency||'')}</span></div>
    <div class="di-row span2"><span class="di-label">儲位</span><span class="di-val" style="color:var(--teal);font-weight:600;">${esc(locDisp)}</span></div>
    <div class="di-row"><span class="di-label">庫存</span><span class="di-val" style="font-size:18px;font-weight:700;">${esc(item.qty)} ${esc(item.unit||'')}</span></div>
    <div class="di-row"><span class="di-label">狀態</span><span class="di-val">${statusChip(item)}</span></div>
    <div class="di-row"><span class="di-label">品牌</span><span class="di-val">${esc(item.brand||'—')}</span></div>
    <div class="di-row"><span class="di-label">建檔日期</span><span class="di-val">${esc(item.date||'—')}</span></div>
  `;
  // Supplier
  const s=item.supplier||{name:item.vendor||''};
  document.getElementById('detailSupplier').innerHTML=`
    <h3>廠商資料</h3>
    <div class="supplier-grid">
      <span class="sg-lbl">廠商</span><span class="sg-val">${esc(s.name||item.vendor||'—')}</span>
      <span class="sg-lbl">聯絡人</span><span class="sg-val">${esc(s.contact||'—')}</span>
      <span class="sg-lbl">電話</span><span class="sg-val">${esc(s.phone||'—')}</span>
      <span class="sg-lbl">產地</span><span class="sg-val">${esc(s.origin||'—')}</span>
      <span class="sg-lbl">交期</span><span class="sg-val">${s.lead?s.lead+'天':'—'}</span>
      <span class="sg-lbl">最小訂量</span><span class="sg-val">${esc(s.moq||'—')}</span>
    </div>
  `;
  // Logs
  const logs=Array.isArray(item.logs)?[...item.logs].reverse():[];
  const logBody=document.querySelector('#logTable tbody');
  logBody.innerHTML=logs.length===0?'<tr><td colspan="5" style="text-align:center;color:var(--text3);">尚無記錄</td></tr>':
    logs.map(l=>{
      const colors={in:'var(--success)',out:'var(--danger)',create:'var(--text3)',return:'var(--purple)','batch-out':'var(--danger)','scan-out':'var(--danger)'};
      const c=colors[l.type]||'var(--text)';
      return `<tr><td style="white-space:nowrap;color:var(--text3);font-size:12px;">${esc(l.time)}</td><td style="font-weight:600;color:${c};">${esc(l.action)}</td><td style="font-weight:700;color:${c};">${esc(l.amount)}</td><td>${esc(l.reason)}</td><td style="font-family:var(--font-mono);font-weight:600;">${esc(l.balance)}</td></tr>`;
    }).join('');
  document.getElementById('detailBackdrop').classList.add('open');
}
function closeDetail(){ document.getElementById('detailBackdrop').classList.remove('open'); currentDetailCode=null; }
function openEditFromDetail(){ closeDetail(); switchPage('manage'); }






// ─── Batch select & delete ──────────────────────────────────────────
function toggleMgSel(checked){
  document.querySelectorAll('.mg-chk').forEach(c=>c.checked=checked);
  const sa=document.getElementById('mgSelectAll'); if(sa) sa.checked=checked;
  updateBatchDeleteBar();
}
function updateBatchDeleteBar(){
  const selected=document.querySelectorAll('.mg-chk:checked').length;
  const bar=document.getElementById('batchDeleteBar');
  if(bar){
    bar.style.display=selected>0?'flex':'none';
    document.getElementById('batchDeleteCount').textContent='已選 '+selected+' 筆';
  }
}
async function batchDelete(){
  const codes=Array.from(document.querySelectorAll('.mg-chk:checked')).map(c=>c.value);
  if(!codes.length){ alert('請先勾選要刪除的材料'); return; }
  if(!confirm('⚠️ 確定要刪除 '+codes.length+' 筆材料嗎？\n\n此操作無法復原！')) return;
  const confirmText=prompt('請輸入 DELETE 確認刪除 '+codes.length+' 筆：');
  if(confirmText!=='DELETE'){ alert('已取消'); return; }
  saveSafetyBackup('批量刪除前');
  const cloudOk=await deleteMaterialsCloud(codes);
  if(!cloudOk) return;
  codes.forEach(code=>{
    inventory=inventory.filter(i=>i.code!==code);
    deleteItemDB(code);
  });
  renderAll();
  alert('✅ 已刪除 '+codes.length+' 筆材料');
}
// Listen for checkbox changes
document.addEventListener('change',function(e){
  if(e.target.classList.contains('mg-chk')) updateBatchDeleteBar();
});

// ─── Excel batch import ──────────────────────────────────────────────
async function importExcel(event){
  const file=event.target.files?.[0]; event.target.value='';
  if(!file) return;
  try{
    await ensureXLSX();
  }catch(err){
    alert('Excel 功能載入失敗，請檢查網路後再試');
    return;
  }
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length){ alert('Excel 檔案是空的'); return; }
      if(!confirm('找到 '+rows.length+' 筆資料，確定匯入嗎？')) return;
      const t=now();
      let created=0, skipped=0;
      const catMap={L:'皮類 Leather',S:'人造皮 Synthetic',T:'布類 Textile',W:'條狀類 Webbing',A:'副料 Auxiliary',O:'OUTSOLE 大底',H:'五金 Hardware'};

      rows.forEach((row,idx)=>{
        const brand=String(row['品牌']||'').trim();
        const vendor=String(row['廠商']||'').trim();
        const season=String(row['季別']||'').trim();
        const matType=String(row['材料類型']||'').trim();
        const name=String(row['材料名稱']||'').trim();
        const colorCode=String(row['顏色編碼 (Pantone)']||row['顏色編碼']||'').trim();
        const colorName=String(row['顏色名稱']||'').trim();
        const qty=parseFloat(row['數量'])||0;
        const unit=String(row['單位 (SF/YD/PCS/PRS)']||row['單位']||'YD').trim().toUpperCase();
        const thickness=String(row['厚度']||'').trim();
        const wh=String(row['主倉 (A-T / TEMP)']||row['主倉']||'').trim().toUpperCase();
        const grid=String(row['格 (1-32)']||row['格']||'').trim();
        const price=String(row['單價']||'0');
        const currency=String(row['幣別 (TWD/USD/RMB)']||row['幣別']||'TWD').trim();
        const note=String(row['備注']||'').trim();

        if(!name||!colorCode){ skipped++; return; }

        // Auto classify
        const mt=matType.toLowerCase();
        let cat='S',type='M',spec='1';
        if(mt.includes('反毛')||mt.includes('suede')){ cat='L';type='C';spec='2'; }
        else if(mt.includes('椰皮')||mt.includes('面皮')){ cat='L';type='C';spec='1'; }
        else if(mt.includes('皮')&&!mt.includes('人造')&&!mt.includes('pu')){ cat='L';type='C';spec='1'; }
        else if(mt.includes('超纖')||mt.includes('超件')||mt.includes('仿超')){ cat='S';type='M';spec='1'; }
        else if(mt.includes('pu')&&mt.includes('濕式')){ cat='S';type='M';spec='2'; }
        else if(mt.includes('pu')&&mt.includes('霧面')){ cat='S';type='M';spec='3'; }
        else if(mt.includes('pu')){ cat='S';type='M';spec='1'; }
        else if(mt.includes('網布')||mt.includes('mesh')){ cat='T';type='M';spec='1'; }
        else if(mt.includes('萊卡')||mt.includes('lycra')){ cat='T';type='K';spec='4'; }
        else if(mt.includes('帆布')){ cat='T';type='C';spec='1'; }
        else if(mt.includes('提花')){ cat='T';type='J';spec='1'; }
        else if(mt.includes('布')){ cat='T';type='K';spec='1'; }
        else if(mt.includes('防水膜')){ cat='A';type='F';spec='10'; }

        const prefix=cat+type+spec;
        const code=nextMaterialCode(prefix,colorCode);

        if(inventory.some(i=>i.code===code)){ skipped++; return; }

        const locCode=wh&&grid?wh+'-'+grid:'';
        const locName=wh&&grid?(wh==='TEMP'?'TEMP — 暫放區':'主倉 '+wh+' — 第 '+grid+' 格'):'';

        const item={
          code, categoryCode:cat, catName:catMap[cat]||'',
          typeName:matType, specName:spec,
          colorCode:colorName||colorCode, colorPantone:colorCode,
          currency, price,
          sizeEU:'',sizeUK:'',sizeUS:'',sizeJP:'',
          locationCode:locCode, locationName:locName,
          qty:qty.toFixed(1), unit:unit||'YD',
          photo:'', productName:name, vendor, brand, thickness,
          supplier:{name:vendor,contact:'',phone:'',origin:'',lead:0,moq:''},
          date:qty>0?new Date().toISOString().split('T')[0]:'',
          status:qty>0?'在庫 In':'待到貨 Pending',
          timestamp:Date.now()-(idx*100),
          note:(season?'季別: '+season+' ':'')+(note||''),
          logs:[{time:t,action:'Excel 匯入 Import',amount:'+'+qty.toFixed(1),reason:'批量 Excel 匯入 / '+brand+' / '+name,balance:qty.toFixed(1),type:'in'}]
        };
        inventory.unshift(item);
        saveItemDB(item);
        created++;
      });
      renderAll();
      flushCloudSaves();
      alert('✅ Excel 匯入完成！\n\n成功：'+created+' 筆\n跳過：'+skipped+' 筆（缺少名稱或顏色編碼）');
    }catch(err){
      console.error(err);
      alert('Excel 讀取失敗，請確認格式是否正確');
    }
  };
  reader.readAsArrayBuffer(file);
}

function downloadTemplate(){
  window.open('https://raymond626.github.io/peihai-warehouse/材料批量建檔模板.xlsx','_blank');
  alert('如果下載沒有開始，請到 GitHub 倉庫下載模板檔案。\n\n或聯繫管理員取得最新模板。');
}

// ─── Batch inbound ──────────────────────────────────────────────────
let batchInboundCart = [];
function addBatchInboundCode(code, qty=1, mergeExisting=false){
  const normalized=String(code||'').trim().toUpperCase();
  const amount=parseFloat(qty)||1;
  if(!normalized) return {ok:false,msg:'空白編碼'};
  const item=inventory.find(i=>(i.code||'').toUpperCase()===normalized);
  if(!item) return {ok:false,msg:'找不到：'+normalized};
  const existing=batchInboundCart.find(c=>c.code===item.code);
  if(existing){
    if(mergeExisting){
      existing.qty=(parseFloat(existing.qty)||0)+amount;
      return {ok:true,merged:true,code:item.code};
    }
    return {ok:false,msg:'已在入庫車中：'+item.code};
  }
  batchInboundCart.push({code:item.code, qty:amount});
  return {ok:true,merged:false,code:item.code};
}
function addToBatchInbound(){
  // permission checked by Supabase RLS
  const code=(document.getElementById('biCode').value||'').trim().toUpperCase();
  const qty=parseFloat(document.getElementById('biQty').value)||1;
  if(!code) return;
  const result=addBatchInboundCode(code,qty,false);
  if(!result.ok){ alert(result.msg); document.getElementById('biCode').value=''; return; }
  document.getElementById('biCode').value='';
  document.getElementById('biQty').value='1';
  renderBatchInbound();
}
function addBulkBatchInbound(){
  const input=document.getElementById('biBulkCodes');
  const defaultQty=parseFloat(document.getElementById('biQty')?.value)||1;
  const rows=parseBulkCodeInput(input?.value,defaultQty);
  if(!rows.length){ alert('請先貼上要入庫的材料編碼'); return; }
  let added=0, merged=0;
  const failed=[];
  rows.forEach(row=>{
    const result=addBatchInboundCode(row.code,row.qty,true);
    if(result.ok){
      if(result.merged) merged++;
      else added++;
    } else failed.push(result.msg);
  });
  if(input) input.value='';
  renderBatchInbound();
  const msg='批量加入完成：新增 '+added+' 筆，合併 '+merged+' 筆'+(failed.length?'\n\n未加入：\n'+failed.slice(0,10).join('\n')+(failed.length>10?'\n...':''):'');
  alert(msg);
}
function removeBatchInbound(code){ batchInboundCart=batchInboundCart.filter(c=>c.code!==code); renderBatchInbound(); }
function renderBatchInbound(){
  const el=document.getElementById('batchInboundList');
  const actions=document.getElementById('batchInboundActions');
  if(!batchInboundCart.length){ el.innerHTML=''; actions.style.display='none'; return; }
  actions.style.display='block';
  if(!document.getElementById('biDate').value) document.getElementById('biDate').valueAsDate=new Date();
  el.innerHTML='<table class="wh-table"><thead><tr><th>材料名稱</th><th>編碼</th><th>顏色</th><th>入庫數量</th><th>移除</th></tr></thead><tbody>'+
    batchInboundCart.map(c=>{
      const item=inventory.find(i=>i.code===c.code)||{};
      return '<tr><td style="font-weight:500;">'+esc(item.productName||'')+'</td><td class="code-cell">'+esc(c.code)+'</td><td>'+esc(item.colorCode||'')+'</td><td><input type="number" value="'+c.qty+'" min="0.1" step="0.1" style="width:80px;font-weight:700;color:var(--success);text-align:center;" onchange="updateBatchInboundQty(\''+esc(c.code)+'\',this.value)"></td><td><button class="btn btn-sm btn-danger" onclick="removeBatchInbound(\''+esc(c.code)+'\')">移除</button></td></tr>';
    }).join('')+'</tbody></table>';
}
function updateBatchInboundQty(code,val){ const c=batchInboundCart.find(x=>x.code===code); if(c) c.qty=parseFloat(val)||0; }
function clearBatchInbound(){ if(confirm('清空入庫車？')){ batchInboundCart=[]; renderBatchInbound(); } }
function submitBatchInbound(){
  // permission checked by Supabase RLS
  if(!batchInboundCart.length){ alert('入庫車為空'); return; }
  const wh=document.getElementById('biWH').value;
  const startGrid=parseInt(document.getElementById('biStartGrid').value)||0;
  const date=document.getElementById('biDate').value;
  if(!wh){ alert('請選擇主倉'); return; }
  if(!startGrid){ alert('請輸入起始格號'); return; }
  if(!date){ alert('請選擇日期'); return; }
  if(!confirm('確定將 '+batchInboundCart.length+' 筆材料入庫到 主倉 '+wh+'？')) return;
  const t=now();
  batchInboundCart.forEach((c,idx)=>{
    const item=inventory.find(i=>i.code===c.code); if(!item) return;
    const grid=startGrid+idx;
    const newQty=(parseFloat(item.qty)+c.qty).toFixed(1);
    item.locationCode=wh+'-'+grid;
    item.locationName='主倉 '+wh+' — 第 '+grid+' 格';
    item.qty=newQty;
    item.status='在庫 In';
    if(!item.date) item.date=date;
    item.logs.push({time:t,action:'批量入庫 Batch In',amount:'+'+c.qty.toFixed(1),reason:'批量入庫 主倉'+wh+' 第'+grid+'格',balance:newQty,type:'in'});
    saveItemDB(item);
  });
  flushCloudSaves();
  alert('✅ 批量入庫成功！共 '+batchInboundCart.length+' 筆');
  batchInboundCart=[];
  renderBatchInbound();
  renderAll();
}

// ─── Copy material (change color only) ──────────────────────────────
function copyMaterial(code){
  // permission checked by Supabase RLS
  const src=inventory.find(i=>i.code===code); if(!src) return;
  const newColor=prompt('輸入新顏色代碼：\n\n原始材料：'+src.productName+'\n原始顏色：'+src.colorCode);
  if(!newColor||!newColor.trim()) return;
  const color=newColor.trim();
  // Build new code with same prefix but new color
  const parts=src.code.split('-');
  let prefix=parts.length>=2?parts.slice(0,-1).join('-'):src.code;
  // Find next sequence for this prefix
  const samePrefix=inventory.filter(i=>(i.code||'').startsWith(prefix));
  // If code format is PREFIX-NNNN-COLOR, increment the NNNN
  let newCode;
  if(parts.length>=3){
    const basePrefix=parts[0];
    const allSame=inventory.filter(i=>(i.code||'').startsWith(basePrefix+'-'));
    const maxSeq=allSame.reduce((max,i)=>{
      const p=i.code.split('-');
      const n=parseInt(p[1])||0;
      return n>max?n:max;
    },0);
    newCode=basePrefix+'-'+String(maxSeq+1).padStart(4,'0')+'-'+color;
  } else {
    newCode=src.code+'-COPY-'+color;
  }
  // Check duplicate
  if(inventory.some(i=>i.code===newCode)){ alert('編碼已存在：'+newCode); return; }
  const newItem=JSON.parse(JSON.stringify(src));
  newItem.code=newCode;
  newItem.colorCode=color;
  newItem.photo='';
  newItem.qty='0.0';
  newItem.status='待到貨 Pending';
  newItem.locationCode='';
  newItem.locationName='';
  newItem.timestamp=Date.now();
  newItem.logs=[{time:now(),action:'複製建檔 Copy',amount:'+0',reason:'從 '+src.code+' 複製，改色 '+color,balance:'0',type:'create'}];
  inventory.unshift(newItem);
  saveItemDB(newItem);
  renderAll();
  alert('✅ 複製成功！\n\n新編碼：'+newCode+'\n顏色：'+color+'\n\n請到貨後再入庫。');
}

// ─── Edit material ──────────────────────────────────────────────────
function openEdit(code){
  // permission checked by Supabase RLS
  const item=inventory.find(i=>i.code===code); if(!item) return;
  document.getElementById('editCodeBanner').innerHTML='編輯材料：'+esc(item.code);
  document.getElementById('editItemCode').value=code;
  document.getElementById('editCode').value=item.code;
  document.getElementById('editName').value=item.productName||'';
  document.getElementById('editColor').value=item.colorCode||'';
  document.getElementById('editThickness').value=item.thickness||'';
  document.getElementById('editVendor').value=item.vendor||'';
  document.getElementById('editBrand').value=item.brand||'';
  document.getElementById('editContact').value=(item.supplier&&item.supplier.contact)||'';
  document.getElementById('editPhone').value=(item.supplier&&item.supplier.phone)||'';
  document.getElementById('editOrigin').value=(item.supplier&&item.supplier.origin)||'';
  document.getElementById('editLead').value=(item.supplier&&item.supplier.lead)||'';
  document.getElementById('editCurrency').value=item.currency||'TWD';
  document.getElementById('editPrice').value=item.price||'';
  document.getElementById('editNote').value=item.note||'';
  document.getElementById('editPhotoBase64').value='';
  const preview=document.getElementById('editPhotoPreview');
  const photoSrc=safePhotoSrc(item.photo);
  if(photoSrc){ preview.src=photoSrc; preview.style.display='block'; }
  else{ preview.style.display='none'; }
  document.getElementById('editBackdrop').classList.add('open');
}
function closeEdit(){ document.getElementById('editBackdrop').classList.remove('open'); }
function submitEdit(){
  const code=document.getElementById('editItemCode').value;
  const item=inventory.find(i=>i.code===code); if(!item) return;
  const name=document.getElementById('editName').value.trim();
  if(!name){ alert('請填寫材料名稱'); return; }
  const newCode=document.getElementById('editCode').value.trim();
  if(newCode && newCode!==item.code){
    // Check for duplicates
    if(inventory.some(i=>i.code===newCode)){ alert('此編碼已被使用，請換一個'); return; }
    // Update code in DB
    deleteItemDB(item.code);
    item.code=newCode;
  }
  item.productName=name;
  item.colorCode=document.getElementById('editColor').value.trim();
  item.thickness=document.getElementById('editThickness').value.trim();
  item.vendor=document.getElementById('editVendor').value.trim();
  item.brand=document.getElementById('editBrand').value.trim();
  item.currency=document.getElementById('editCurrency').value;
  item.price=document.getElementById('editPrice').value;
  item.note=document.getElementById('editNote').value.trim();
  if(!item.supplier) item.supplier={};
  item.supplier.name=item.vendor;
  item.supplier.contact=document.getElementById('editContact').value.trim();
  item.supplier.phone=document.getElementById('editPhone').value.trim();
  item.supplier.origin=document.getElementById('editOrigin').value.trim();
  item.supplier.lead=parseInt(document.getElementById('editLead').value)||0;
  const newPhoto=document.getElementById('editPhotoBase64').value;
  if(newPhoto) item.photo=newPhoto;
  saveItemDB(item);
  closeEdit();
  renderAll();
  if(currentDetailCode===code) openDetail(code);
  alert('✅ 已儲存變更');
}

// ─── Delete ───────────────────────────────────────────────────────────
function promptDelete(code){
  if(!isAdmin()){ alert('只有管理員可以刪除材料'); return; }
  pendingDeleteCode=code;
  document.getElementById('deleteItemCode').textContent=code;
  document.getElementById('deleteBackdrop').classList.add('open');
}
async function confirmDelete(){
  if(!pendingDeleteCode) return;
  const code=pendingDeleteCode;
  saveSafetyBackup('刪除材料前');
  const cloudOk=await deleteMaterialsCloud([code]);
  if(!cloudOk) return;
  inventory=inventory.filter(i=>i.code!==code);
  deleteItemDB(code);
  document.getElementById('deleteBackdrop').classList.remove('open');
  pendingDeleteCode=null;
  renderAll();
}

// ─── Print label ──────────────────────────────────────────────────────
async function printLabel(){
  const item=inventory.find(i=>i.code===currentDetailCode); if(!item) return;
  const w=window.open('','_blank','width=500,height=560');
  if(!w){ alert('請允許彈出視窗後再列印'); return; }
  try{ await ensureQRCode(); }catch(e){ alert('QR 功能載入失敗，請檢查網路後再試'); w.close(); return; }
  const qrCanvas=document.querySelector('#detailQR canvas');
  const qrImg=qrCanvas?qrCanvas.toDataURL():createQRData(item.code);
  w.document.write(`<html><head><title>${item.code}</title><style>
    body{font-family:'Microsoft JhengHei',Arial,sans-serif;margin:16px;font-size:13px;}
    .card{border:2px solid #000;border-radius:8px;padding:14px;max-width:320px;}
    h2{text-align:center;font-size:14px;border-bottom:1.5px solid #000;padding-bottom:6px;margin:0 0 8px;}
    .code{text-align:center;font-family:monospace;font-size:17px;font-weight:900;margin:6px 0;}
    p{margin:3px 0;font-weight:600;}
    img{display:block;width:110px;height:110px;margin:8px auto 0;}
    @media print{body{margin:5mm;}}
  </style></head><body><div class="card">
    <h2>北海企業樣品室</h2>
    <div class="code">${esc(item.code)}</div>
    <p>材料：${esc(item.productName||'')}</p>
    ${getOutsoleSize(item)?`<p>尺寸：${esc(getOutsoleSize(item))}</p>`:''}
    <p>規格：${esc(item.specName||'')} ${item.colorCode?'/ '+item.colorCode:''}</p>
    ${item.thickness?`<p>厚度：${esc(item.thickness)}</p>`:''}
    <p>廠商：${esc(item.vendor||'')}</p>
    <p>儲位：${esc(item.locationCode||'尚未入庫')}</p>
    ${qrImg?`<img src="${qrImg}" alt="QR">`:''}
  </div><script>setTimeout(()=>{window.print();window.close();},600);<\/script>`);
  w.document.close();
}

// ─── Batch QR print ───────────────────────────────────────────────────
function createQRData(text){
  if(!window.QRCode) return '';
  const h=document.createElement('div'); h.style.cssText='position:fixed;left:-9999px;top:-9999px;';
  document.body.appendChild(h);
  new QRCode(h,{text,width:130,height:130,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.H});
  const c=h.querySelector('canvas'), img=h.querySelector('img');
  const data=c?c.toDataURL():(img?img.src:'');
  h.remove(); return data;
}
function togglePrintSel(checked){
  document.querySelectorAll('.print-chk').forEach(c=>c.checked=checked);
  const sa=document.getElementById('ovSelectAll'); if(sa) sa.checked=checked;
}
async function printSelectedQR(){
  const codes=Array.from(document.querySelectorAll('.print-chk:checked')).map(c=>c.value);
  if(!codes.length){ alert('請先勾選要列印 QR Code 的材料'); return; }
  const items=codes.map(c=>inventory.find(i=>i.code===c)).filter(Boolean);
  const w=window.open('','_blank','width=900,height=700');
  if(!w){ alert('請允許彈出視窗後再列印'); return; }
  try{ await ensureQRCode(); }catch(e){ alert('QR 功能載入失敗，請檢查網路後再試'); w.close(); return; }
  const labels=items.map(item=>{
    const qr=createQRData(item.code);
    return `<div class="label">
      <h2>北海企業樣品室</h2>
      <div class="code">${esc(item.code)}</div>
      <p>${esc(item.productName||'')}</p>
      ${getOutsoleSize(item)?`<p>${esc(getOutsoleSize(item))}</p>`:''}
      <p>${esc(item.catName?.split(' ')[0]||'')} — ${esc(item.specName||'')} ${item.colorCode?'/ '+item.colorCode:''}</p>
      ${item.thickness?`<p>厚度：${esc(item.thickness)}</p>`:''}
      <p>廠商：${esc(item.vendor||'')}</p>
      <p>儲位：${esc(item.locationCode||'尚未入庫')}</p>
      ${qr?`<img src="${qr}" alt="QR">`:''}
    </div>`;
  }).join('');
  w.document.write(`<html><head><title>批量 QR 列印</title><style>
    body{font-family:'Microsoft JhengHei',Arial,sans-serif;margin:12px;}
    .sheet{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;}
    .label{border:2px solid #000;border-radius:8px;padding:10px;page-break-inside:avoid;}
    h2{text-align:center;font-size:13px;border-bottom:1.5px solid #000;padding-bottom:5px;margin:0 0 6px;}
    .code{text-align:center;font-family:monospace;font-size:16px;font-weight:900;margin:4px 0;}
    p{margin:2px 0;font-size:12px;font-weight:600;}
    img{display:block;width:100px;height:100px;margin:8px auto 0;}
    @media print{body{margin:6mm;}.sheet{gap:6px;}}
  </style></head><body><div class="sheet">${labels}</div>
  <script>setTimeout(()=>{window.print();},500);<\/script>`);
  w.document.close();
}

// ─── Batch cart ───────────────────────────────────────────────────────
function addBatchCartCode(code, qty=1, mergeExisting=false){
  const normalized=String(code||'').trim().toUpperCase();
  if(!normalized) return {ok:false,msg:'空白編碼'};
  const item=inventory.find(i=>(i.code||'').toUpperCase()===normalized);
  if(!item) return {ok:false,msg:'找不到材料編碼：'+normalized};
  if(getStatus(item)==='pending') return {ok:false,msg:'此材料尚未入庫：'+item.code};
  const max=parseFloat(item.qty)||0;
  if(max<=0) return {ok:false,msg:'此材料已耗盡：'+item.code};
  let amount=parseFloat(qty)||1;
  if(amount<=0) amount=1;
  if(amount>max) amount=max;
  const existing=batchCart.find(c=>c.code===item.code);
  if(existing){
    if(mergeExisting){
      existing.outQty=Math.min(max,(parseFloat(existing.outQty)||0)+amount);
      return {ok:true,merged:true,code:item.code};
    }
    return {ok:false,msg:'已在取料車中：'+item.code};
  }
  batchCart.push({code:item.code, outQty:amount});
  return {ok:true,merged:false,code:item.code};
}
function addToBatch(){
  // permission checked by Supabase RLS
  const input=document.getElementById('batchInput');
  const code=(input.value||'').trim().toUpperCase();
  if(!code) return;
  const result=addBatchCartCode(code,1,false);
  if(!result.ok){ alert(result.msg); input.value=''; return; }
  input.value=''; renderBatch();
  updateBatchCount();
}
function addBulkToBatch(){
  const input=document.getElementById('batchBulkInput');
  const rows=parseBulkCodeInput(input?.value,1);
  if(!rows.length){ alert('請先貼上要取料的材料編碼'); return; }
  let added=0, merged=0;
  const failed=[];
  rows.forEach(row=>{
    const result=addBatchCartCode(row.code,row.qty,true);
    if(result.ok){
      if(result.merged) merged++;
      else added++;
    } else failed.push(result.msg);
  });
  if(input) input.value='';
  renderBatch();
  updateBatchCount();
  const msg='批量加入完成：新增 '+added+' 筆，合併 '+merged+' 筆'+(failed.length?'\n\n未加入：\n'+failed.slice(0,10).join('\n')+(failed.length>10?'\n...':''):'');
  alert(msg);
}
function removeFromBatch(code){ batchCart=batchCart.filter(c=>c.code!==code); renderBatch(); updateBatchCount(); }
function updateBatchQty(code, val){
  const item=inventory.find(i=>i.code===code);
  const max=item?parseFloat(item.qty):0;
  let n=parseFloat(val);
  if(isNaN(n)||n<0) n=0;
  if(n>max){ alert(`數量不可超過庫存 (${max})`); n=max; }
  const c=batchCart.find(x=>x.code===code);
  if(c) c.outQty=n;
  renderBatch(); updateBatchCount();
}
function clearBatch(){ if(batchCart.length&&confirm('確定清空取料車？')){ batchCart=[]; renderBatch(); updateBatchCount(); } }
function updateBatchCount(){
  const el=document.getElementById('nc-batch');
  if(el) el.textContent=batchCart.length||'0';
}
function renderBatch(){
  const tbody=document.getElementById('batchBody');
  if(!tbody) return;
  if(!batchCart.length){
    tbody.innerHTML='<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--text3);">取料車為空 — 掃描或輸入材料編碼加入</td></tr>';
    return;
  }
  tbody.innerHTML=batchCart.map(c=>{
    const item=inventory.find(i=>i.code===c.code)||{};
    return `<tr>
      <td>${thumbCell(item)}</td>
      <td><div class="mat-name">${esc(item.productName||'')}</div></td>
      <td class="code-cell">${esc(c.code)}</td>
      <td>${catChip(getItemCat(item))}</td>
      <td style="color:var(--teal);font-weight:600;">${esc(item.locationCode||'—')}</td>
      <td>${esc(item.qty||'0')} <span style="color:var(--text3);font-size:11px;">${esc(item.unit||'')}</span></td>
      <td><input type="number" value="${c.outQty}" min="0.1" max="${item.qty||0}" step="0.1" onchange="updateBatchQty('${esc(c.code)}',this.value)" style="width:80px;color:var(--danger);font-weight:700;text-align:center;"></td>
      <td><button class="btn btn-sm btn-danger" onclick="removeFromBatch('${esc(c.code)}')">移除</button></td>
    </tr>`;
  }).join('');
}
function submitBatch(){
  // permission checked by Supabase RLS
  if(!batchCart.length){ alert('取料車為空'); return; }
  const reason=(document.getElementById('batchReason').value||'').trim();
  if(!reason){ alert('請輸入出庫原因'); return; }
  if(batchCart.some(c=>parseFloat(c.outQty)<=0)){ alert('有數量為 0 的項目'); return; }
  if(!confirm(`確定將 ${batchCart.length} 筆材料出庫嗎？`)) return;
  const t=now();
  batchCart.forEach(c=>{
    const item=inventory.find(i=>i.code===c.code); if(!item) return;
    const newQty=(parseFloat(item.qty)-parseFloat(c.outQty)).toFixed(1);
    item.qty=newQty;
    if(parseFloat(item.qty)<=0) item.status='已耗盡 Empty';
    item.logs.push({time:t,action:'批量出庫 Batch Out',amount:'-'+parseFloat(c.outQty).toFixed(1),reason,balance:newQty,type:'batch-out'});
    saveItemDB(item);
  });
  flushCloudSaves();
  alert('✅ 批量出庫成功！');
  batchCart=[]; document.getElementById('batchReason').value='';
  renderBatch(); renderAll(); switchPage('overview');
}

// ─── Render ───────────────────────────────────────────────────────────
function renderStats(){
  const stTotal=document.getElementById('st-total');
  const stPending=document.getElementById('st-pending');
  const stIn=document.getElementById('st-in');
  const stEmpty=document.getElementById('st-empty');
  if(stTotal) stTotal.textContent=inventory.length;
  if(stPending) stPending.textContent=inventory.filter(i=>getStatus(i)==='pending').length;
  if(stIn) stIn.textContent=inventory.filter(i=>getStatus(i)==='in').length;
  if(stEmpty) stEmpty.textContent=inventory.filter(i=>getStatus(i)==='empty').length;
  const el=document.getElementById('nc-overview');
  if(el) el.textContent=inventory.length;
}

// ─── Analytics ──────────────────────────────────────────────────────
function renderAnalytics(){
  const total=inventory.length;
  const inStock=inventory.filter(i=>getStatus(i)==='in').length;
  const pending=inventory.filter(i=>getStatus(i)==='pending').length;
  const empty=inventory.filter(i=>getStatus(i)==='empty').length;

  // Summary cards
  document.getElementById('analyticsSummary').innerHTML=`
    <div class="stat-card"><div class="stat-label">全部材料</div><div class="stat-val">${total}</div></div>
    <div class="stat-card s-pending"><div class="stat-label">待收料</div><div class="stat-val">${pending}</div></div>
    <div class="stat-card s-in"><div class="stat-label">在庫</div><div class="stat-val">${inStock}</div></div>
    <div class="stat-card s-empty"><div class="stat-label">已耗盡</div><div class="stat-val">${empty}</div></div>
  `;

  // Brand breakdown
  const brands={};
  inventory.forEach(i=>{
    const b=i.brand||'未分類';
    if(!brands[b]) brands[b]={count:0,inStock:0,empty:0};
    brands[b].count++;
    if(getStatus(i)==='in') brands[b].inStock++;
    if(getStatus(i)==='empty') brands[b].empty++;
  });
  const brandArr=Object.entries(brands).sort((a,b)=>b[1].count-a[1].count);
  document.getElementById('brandBody').innerHTML=brandArr.map(([name,d])=>{
    const pct=total?(d.count/total*100).toFixed(1):0;
    return `<tr>
      <td style="font-weight:600;">${esc(name)}</td>
      <td style="font-weight:700;">${d.count}</td>
      <td style="color:var(--success);">${d.inStock}</td>
      <td style="color:${d.empty>0?'var(--danger)':'var(--text3)'};">${d.empty}</td>
      <td><div style="display:flex;align-items:center;gap:6px;"><div style="width:60px;height:6px;background:var(--border);border-radius:3px;overflow:hidden;"><div style="width:${pct}%;height:100%;background:var(--accent);border-radius:3px;"></div></div><span style="font-size:11px;color:var(--text3);">${pct}%</span></div></td>
    </tr>`;
  }).join('');

  // Category breakdown
  const cats={};
  inventory.forEach(i=>{
    const cat=getItemCat(i);
    const label=CAT_LABELS[cat]||cat;
    if(!cats[label]) cats[label]={count:0,inStock:0,empty:0};
    cats[label].count++;
    if(getStatus(i)==='in') cats[label].inStock++;
    if(getStatus(i)==='empty') cats[label].empty++;
  });
  const catArr=Object.entries(cats).sort((a,b)=>b[1].count-a[1].count);
  document.getElementById('catBody').innerHTML=catArr.map(([name,d])=>{
    const pct=total?(d.count/total*100).toFixed(1):0;
    return `<tr>
      <td style="font-weight:600;">${esc(name)}</td>
      <td style="font-weight:700;">${d.count}</td>
      <td style="color:var(--success);">${d.inStock}</td>
      <td style="color:${d.empty>0?'var(--danger)':'var(--text3)'};">${d.empty}</td>
      <td><div style="display:flex;align-items:center;gap:6px;"><div style="width:60px;height:6px;background:var(--border);border-radius:3px;overflow:hidden;"><div style="width:${pct}%;height:100%;background:var(--purple);border-radius:3px;"></div></div><span style="font-size:11px;color:var(--text3);">${pct}%</span></div></td>
    </tr>`;
  }).join('');

  // Vendor breakdown
  const vendors={};
  inventory.forEach(i=>{
    const v=i.vendor||i.supplier?.name||'未知';
    if(!vendors[v]) vendors[v]={count:0,totalQty:0};
    vendors[v].count++;
    vendors[v].totalQty+=parseFloat(i.qty)||0;
  });
  const vendorArr=Object.entries(vendors).sort((a,b)=>b[1].count-a[1].count);
  document.getElementById('vendorBody').innerHTML=vendorArr.map(([name,d])=>{
    const pct=total?(d.count/total*100).toFixed(1):0;
    return `<tr>
      <td style="font-weight:600;">${esc(name)}</td>
      <td style="font-weight:700;">${d.count}</td>
      <td>${d.totalQty.toFixed(1)}</td>
      <td><div style="display:flex;align-items:center;gap:6px;"><div style="width:60px;height:6px;background:var(--border);border-radius:3px;overflow:hidden;"><div style="width:${pct}%;height:100%;background:var(--teal);border-radius:3px;"></div></div><span style="font-size:11px;color:var(--text3);">${pct}%</span></div></td>
    </tr>`;
  }).join('');

  // Warehouse breakdown
  const whs={};
  inventory.forEach(i=>{
    if(!i.locationCode) return;
    const parts=i.locationCode.split('-');
    const wh=parts[0]||'未分配';
    const whLabel=wh==='TEMP'?'TEMP 暫放區':'主倉 '+wh;
    if(!whs[whLabel]) whs[whLabel]={count:0,grids:new Set()};
    whs[whLabel].count++;
    if(parts[1]) whs[whLabel].grids.add(parts[1]);
  });
  const whArr=Object.entries(whs).sort((a,b)=>b[1].count-a[1].count);
  document.getElementById('warehouseBody').innerHTML=whArr.map(([name,d])=>{
    const pct=total?(d.count/total*100).toFixed(1):0;
    return `<tr>
      <td style="font-weight:600;color:var(--teal);">${esc(name)}</td>
      <td style="font-weight:700;">${d.count}</td>
      <td>${d.grids.size}</td>
      <td><div style="display:flex;align-items:center;gap:6px;"><div style="width:60px;height:6px;background:var(--border);border-radius:3px;overflow:hidden;"><div style="width:${pct}%;height:100%;background:var(--warning);border-radius:3px;"></div></div><span style="font-size:11px;color:var(--text3);">${pct}%</span></div></td>
    </tr>`;
  }).join('');

  // Low stock / empty items
  const lowItems=inventory.filter(i=>getStatus(i)==='empty'||parseFloat(i.qty)<=2);
  document.getElementById('lowStockBody').innerHTML=lowItems.length===0
    ?'<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--text3);">目前沒有需注意的材料 ✅</td></tr>'
    :lowItems.map(i=>`<tr>
      <td class="code-cell">${esc(i.code)}</td>
      <td style="font-weight:500;">${esc(i.productName||'')}</td>
      <td>${esc(i.brand||'')}</td>
      <td style="font-weight:700;color:${getStatus(i)==='empty'?'var(--danger)':'var(--warning)'};">${esc(i.qty)} ${esc(i.unit||'')}</td>
      <td style="color:var(--teal);">${esc(i.locationCode||'—')}</td>
      <td>${statusChip(i)}</td>
    </tr>`).join('');
}

function getActivePage(){
  const active=document.querySelector('.page.active');
  if(!active?.id) return 'overview';
  return active.id.replace(/^page/,'').replace(/^./,c=>c.toLowerCase());
}
function renderPage(page){
  if(page==='overview') renderOverview();
  else if(page==='search') renderSearch();
  else if(page==='manage') renderManageTable();
  else if(page==='analytics') renderAnalytics();
  else if(page==='batch') renderBatch();
}
function renderAll(){
  renderStats();
  updateBatchCount();
  renderPage(getActivePage());
}

function renderOverview(){
  const q=(document.getElementById('ovSearchInput')?.value||'').toLowerCase();
  const cat=document.getElementById('ovCatFilter')?.value||'';
  const type=document.getElementById('ovTypeFilter')?.value||'';
  const status=document.getElementById('ovStatusFilter')?.value||'';
  const items=filterItems(q,cat,type,status);
  const tbody=document.getElementById('ovBody');
  if(!tbody) return;
  document.getElementById('ovResultCount').textContent=`(${items.length} 筆)`;
  if(!items.length){ tbody.innerHTML='<tr><td colspan="11" style="text-align:center;padding:32px;color:var(--text3);">沒有符合條件的材料</td></tr>'; return; }
  tbody.innerHTML=items.map(item=>`<tr>
    <td><input type="checkbox" class="print-chk" value="${esc(item.code)}"></td>
    <td>${thumbCell(item)}</td>
    <td><div class="mat-name">${esc(item.productName||'')}</div><div class="mat-sub">${esc(item.brand||'')} ${item.colorCode?'· '+item.colorCode:''}</div></td>
    <td class="code-cell">${esc(item.code)}</td>
    <td style="font-weight:500;">${esc(item.colorCode||'—')}</td>
    <td>${catChip(getItemCat(item))}</td>
    <td style="font-size:12px;color:var(--text2);">${esc(item.specName||'')} ${item.thickness?'/ '+item.thickness:''}</td>
    <td style="color:var(--teal);font-weight:600;font-size:12px;">${esc(item.locationCode||'—')}</td>
    <td style="font-weight:700;">${esc(item.qty||'0')} <span style="color:var(--text3);font-size:11px;">${esc(item.unit||'')}</span></td>
    <td style="font-size:12px;color:var(--text2);">${esc(item.supplier?.name||item.vendor||'—')}</td>
    <td>${statusChip(item)}</td>
    <td><button class="btn btn-sm" onclick="openEdit('${esc(item.code)}')">編輯</button></td>
  </tr>`).join('');
}

function renderManageTable(){
  const q=(document.getElementById('mgSearch')?.value||'').toLowerCase();
  const cat=document.getElementById('mgCat')?.value||'';
  const type=document.getElementById('mgType')?.value||'';
  const items=filterItems(q,cat,type,'');
  const tbody=document.getElementById('mgBody');
  if(!tbody) return;
  if(!items.length){ tbody.innerHTML='<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--text3);">沒有符合條件的材料</td></tr>'; return; }
  tbody.innerHTML=items.map(item=>{
    const s=getStatus(item);
    const canIn=s==='pending';
    const canOut=s==='in'&&parseFloat(item.qty)>0;
    return `<tr>
      <td><input type="checkbox" class="mg-chk" value="${esc(item.code)}"></td>
      <td>${thumbCell(item)}</td>
      <td><div class="mat-name">${esc(item.productName||'')}</div><div class="mat-sub">${esc(item.brand||'')} ${item.colorCode?'· '+item.colorCode:''}</div></td>
      <td class="code-cell">${esc(item.code)}</td>
      <td>${catChip(getItemCat(item))}</td>
      <td style="font-size:12px;color:var(--text2);">${esc(item.thickness||'—')}</td>
      <td style="font-weight:700;">${esc(item.qty||'0')} <span style="color:var(--text3);font-size:11px;">${esc(item.unit||'')}</span></td>
      <td style="color:var(--teal);font-weight:600;font-size:12px;">${esc(item.locationCode||'—')}</td>
      <td>${statusChip(item)}</td>
      <td class="action-cell">
        ${canIn?`<button class="btn btn-sm btn-success" onclick="openInbound('${esc(item.code)}')">入庫</button>`:''}
        ${canOut?`<button class="btn btn-sm btn-danger" onclick="openOutbound('${esc(item.code)}','out')">出庫</button>`:''}
        <button class="btn btn-sm btn-purple" onclick="openOutbound('${esc(item.code)}','return')" title="還料">還料</button>
        <button class="btn btn-sm" onclick="openDetail('${esc(item.code)}')">詳情</button>
        ${isAdmin()?`<button class="btn btn-sm btn-danger btn-ghost" onclick="promptDelete('${esc(item.code)}')" title="刪除" style="padding:0 6px;">✕</button>`:''}
      </td>
    </tr>`;
  }).join('');
}

function renderSearch(){
  const q=(document.getElementById('srInput')?.value||'').toLowerCase();
  const cat=document.getElementById('srCat')?.value||'';
  const type=document.getElementById('srType')?.value||'';
  const status=document.getElementById('srStatus')?.value||'';
  const items=filterItems(q,cat,type,status);
  const el=document.getElementById('srCount');
  if(el) el.textContent=items.length;
  const tbody=document.getElementById('srBody');
  if(!tbody) return;
  if(!items.length){ tbody.innerHTML='<tr><td colspan="12" style="text-align:center;padding:32px;color:var(--text3);">沒有符合條件的材料</td></tr>'; return; }
  tbody.innerHTML=items.map(item=>`<tr>
    <td>${thumbCell(item)}</td>
    <td><div class="mat-name">${esc(item.productName||'')}</div></td>
    <td class="code-cell">${esc(item.code)}</td>
    <td style="font-weight:500;">${esc(item.colorCode||'—')}</td>
    <td>${catChip(getItemCat(item))}</td>
    <td style="font-size:12px;">${esc(item.specName||'')} ${item.thickness?'/ '+item.thickness:''}</td>
    <td style="font-weight:700;">${esc(item.qty||'0')} <span style="color:var(--text3);font-size:11px;">${esc(item.unit||'')}</span></td>
    <td style="color:var(--teal);font-weight:600;font-size:12px;">${esc(item.locationCode||'—')}</td>
    <td style="font-size:12px;">${esc(item.supplier?.name||item.vendor||'—')}</td>
    <td style="font-size:12px;">${esc(item.brand||'—')}</td>
    <td>${statusChip(item)}</td>
    <td><button class="btn btn-sm" onclick="openEdit('${esc(item.code)}')">編輯</button></td>
  </tr>`).join('');
}

function printSearchResult(){
  const tbody=document.getElementById('srBody');
  if(!tbody) return;
  const w=window.open('','_blank','width=1200,height=800');
  w.document.write(`<html><head><title>材料總表列印</title><style>
    body{font-family:'Microsoft JhengHei',Arial,sans-serif;font-size:12px;margin:10mm;}
    h1{font-size:16px;margin-bottom:8px;} .sub{color:#666;font-size:11px;margin-bottom:12px;}
    table{width:100%;border-collapse:collapse;} th,td{border:1px solid #ccc;padding:5px 7px;text-align:left;font-size:11px;}
    th{background:#f5f5f5;font-weight:700;} tr:nth-child(even) td{background:#fafafa;}
    @media print{body{margin:8mm;}}
  </style></head><body>
  <h1>北海企業樣品室 — 材料總表</h1>
  <div class="sub">列印時間：${now()} ／ 共 ${document.getElementById('srCount')?.textContent||0} 筆</div>
  <table><thead><tr><th>材料名稱</th><th>編碼</th><th>大類</th><th>規格</th><th>庫存</th><th>儲位</th><th>廠商</th><th>品牌</th><th>狀態</th></tr></thead>
  <tbody>${tbody.innerHTML.replace(/<img[^>]*>/g,'').replace(/<div class="mat-name">/g,'').replace(/<\/div>/g,'').replace(/class="[^"]*"/g,'')}</tbody>
  </table><script>setTimeout(()=>{window.print();},500);<\/script>`);
  w.document.close();
}

// ─── Type filter helpers ──────────────────────────────────────────────
function updateTypeFilter(catId, typeId, cb){
  const cat=document.getElementById(catId)?.value;
  const sel=document.getElementById(typeId);
  if(!sel) return;
  sel.innerHTML='<option value="">所有種類</option>';
  if(cat&&cat!=='O'&&MAT[cat]){
    Object.entries(MAT[cat].types).forEach(([k,v])=>{ sel.innerHTML+=`<option value="${v}">${k} — ${v}</option>`; });
  }
  if(cb) cb();
}
function updateOvTypes(){ updateTypeFilter('ovCatFilter','ovTypeFilter',renderOverview); }
function updateSrTypes(){ updateTypeFilter('srCat','srType',renderSearch); }
function updateMgTypes(){ updateTypeFilter('mgCat','mgType',renderManageTable); }

// ─── Navigation ───────────────────────────────────────────────────────
function switchPage(page){
  ['overview','search','manage','analytics','batch'].forEach(p=>{
    const pg=document.getElementById('page'+p.charAt(0).toUpperCase()+p.slice(1));
    const nav=document.getElementById('nav'+p.charAt(0).toUpperCase()+p.slice(1));
    if(pg){ pg.style.display='none'; pg.classList.remove('active'); }
    if(nav) nav.classList.remove('active');
  });
  const target=document.getElementById('page'+page.charAt(0).toUpperCase()+page.slice(1));
  const navEl=document.getElementById('nav'+page.charAt(0).toUpperCase()+page.slice(1));
  if(target){ target.style.display='block'; target.classList.add('active'); }
  if(navEl) navEl.classList.add('active');
  renderPage(page);
  if(page==='batch') document.getElementById('batchInput')?.focus();
}

// ─── Storage / Backup ─────────────────────────────────────────────────
function setStorageStatus(msg){ const el=document.getElementById('storageStatus'); if(el) el.textContent=msg; }
function createBackup(){
  return JSON.stringify({app:'Peihai WMS',version:2,exportedAt:new Date().toISOString(),inventory,brands:customBrands},null,2);
}
function saveSafetyBackup(reason){
  try{
    localStorage.setItem('peihai_last_safety_backup',createBackup());
    localStorage.setItem('peihai_last_safety_backup_reason',(reason||'自動備份')+' · '+now());
  }catch(e){}
}
async function chooseFolder(){
  if(!window.showDirectoryPicker){ alert('此瀏覽器不支援資料夾選取，請用下載 / 匯入備份'); return; }
  try{ dataDirectoryHandle=await window.showDirectoryPicker({mode:'readwrite'}); setStorageStatus('已選擇資料夾'); }
  catch(e){ setStorageStatus('尚未選擇資料夾'); }
}
async function saveToFolder(){
  try{
    if(!dataDirectoryHandle) await chooseFolder();
    if(!dataDirectoryHandle) return;
    const fh=await dataDirectoryHandle.getFileHandle(BACKUP_FILE,{create:true});
    const wr=await fh.createWritable();
    await wr.write(createBackup()); await wr.close();
    setStorageStatus('已儲存 '+now());
    alert('已儲存到資料夾');
  }catch(e){ alert('儲存失敗，請重試'); }
}
async function loadFromFolder(){
  try{
    if(!dataDirectoryHandle) await chooseFolder();
    if(!dataDirectoryHandle) return;
    const fh=await dataDirectoryHandle.getFileHandle(BACKUP_FILE);
    const file=await fh.getFile();
    await importText(await file.text(), file.name);
  }catch(e){ alert('找不到備份檔，請確認資料夾或改用匯入備份'); }
}
function downloadBackup(){
  const blob=new Blob([createBackup()],{type:'application/json'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='peihai-backup-'+new Date().toISOString().slice(0,10).replace(/-/g,'')+'.json';
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  setStorageStatus('已下載備份');
}
async function importBackup(event){
  const file=event.target.files?.[0]; event.target.value='';
  if(!file) return;
  await importText(await file.text(), file.name);
}
async function importText(text, name){
  if(cloudEnabled&&!isAdmin()){ alert('只有管理員可以匯入'); return; }
  try{
    const p=JSON.parse(text);
    const next=Array.isArray(p)?p:p.inventory;
    if(!Array.isArray(next)){ alert('不是正確的備份檔格式'); return; }
    if(!confirm('匯入後會取代目前所有材料資料，確定繼續嗎？')) return;
    saveSafetyBackup('匯入前');
    inventory=next.map(i=>({...i,logs:Array.isArray(i.logs)?i.logs:[],timestamp:i.timestamp||Date.now(),qty:i.qty===undefined?'0.0':String(i.qty)}));
    inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
    if(p.brands) { customBrands=p.brands; localStorage.setItem('peihai_brands',JSON.stringify(customBrands)); renderBrands(); }
    localStorage.setItem('peihai_ex','1');
    replaceInDB(inventory,()=>{ renderAll(); setStorageStatus('已載入：'+name); alert('資料載入完成'); },true);
  }catch(e){ alert('讀取備份檔失敗，請確認格式'); }
}
async function clearAllData(){
  if(!isAdmin()){ alert('只有管理員可以清除所有資料'); return; }
  const confirm1=confirm('⚠️ 確定要清除所有庫存資料嗎？此操作無法復原。');
  if(!confirm1) return;
  const code2=prompt('請輸入 "CLEAR" 確認清除：');
  if(code2!=='CLEAR'){ alert('已取消'); return; }
  try{
    saveSafetyBackup('清除全部前');
    if(cloudEnabled&&supabaseClient&&currentUser){
      const {error}=await supabaseClient.from('materials').delete().neq('code','');
      if(error){ setCloud('雲端清除失敗','cloud-err'); alert('雲端清除失敗：'+error.message); return; }
    }
    if(db) db.transaction(['inventory'],'readwrite').objectStore('inventory').clear();
    localStorage.setItem('peihai_ex','1');
    inventory=[]; renderAll();
  }catch(e){ console.error(e); alert('清除失敗，請重新整理後再試'); }
}

// ─── Example data ─────────────────────────────────────────────────────
function clearSrFilters(){
  ['srInput','srCat','srType','srStatus'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  renderSearch();
}
function loadExamples(){
  if(!isAdmin()){ alert('只有管理員可以載入範例'); return; }
  if(!confirm('這會清空目前庫存並載入範例資料，確定嗎？')) return;
  db.transaction(['inventory'],'readwrite').objectStore('inventory').clear().onsuccess=()=>{
    inventory=[]; injectMock(true);
  };
}
function injectMock(showAlert){
  const t=now();
  const rows=[
    ['O','','','OUTSOLE','BLK','185.00','TWD','A','5','','120','PRS','宏達橡膠','Trail Pro','In','Outdoor Trekker 2026','','42','7.5','8.5','26.5'],
    ['S','V','1','面皮紋 Smooth','RED','95.00','TWD','B','3','','180','YD','新合成皮革','通用開發 General','In','Red City Sneaker',''],
    ['L','S','2','反毛 Suede','TAN','4.80','USD','A','8','','260','SF','義大利 MASTrotto','Onitsuka Tiger','In','suede 1.2-1.4',''],
    ['L','C','1','面皮 Full Grain','BLU','5.20','USD','A','12','','180','SF','聯騏皮業','Alpinestars','In','Blue Leather Boot',''],
    ['L','C','8','後處理 Treated','BLK','5.70','USD','A','16','','160','SF','大成皮業','通用開發 General','In','Black Pebbled Upper',''],
    ['H','E','12','12 mm','GLD','1.40','TWD','C','2','','3000','PCS','裕興五金','Onitsuka Tiger','In','Gold Eyelet Pack',''],
    ['H','E','12','12 mm','SIL','1.20','TWD','','','','0','PCS','裕興五金','通用開發 General','Pending','Silver Eyelet Incoming',''],
    ['T','M','3','三明治 Sandwich','BLU','55.00','TWD','B','7','','200','YD','優格布料','Trail Pro','In','三明治網布 Blue',''],
    ['A','S','30','泡棉 3mm','—','12.00','TWD','C','1','','500','PCS','輔料王','通用開發 General','In','泡棉 3.0mm',''],
    ['W','L','08','8 mm','TEAL','6.50','TWD','B','4','','500','PRS','冠美織帶','FitFlop','In','Teal Lace Pack',''],
  ];
  const catMap={L:'皮類 Leather',S:'人造皮 Synthetic',T:'布類 Textile',W:'條狀類 Webbing',A:'副料 Auxiliary',O:'OUTSOLE 大底',H:'五金 Hardware'};
  let count=0;
  rows.forEach(data=>{
    const [cat,type,specCode,specName,colorCode,price,currency,warehouse,rack,,qty,unit,vendor,brand,statusFlag,productName,photo,sizeEU,sizeUK,sizeUS,sizeJP]=data;
    count++;
    const prefix=cat===type&&!type?'O':cat+type+specCode;
    const finalCode=cat==='O'?productName:`${prefix}-${String(count).padStart(4,'0')}-${colorCode}`;
    const locCode=warehouse?`${warehouse}-${rack}`:'';
    const locName=warehouse?`主倉 ${warehouse} — 第 ${rack} 格`:'';
    const currentQty=parseFloat(qty);
    const logs=[{time:t,action:'建檔 Master Created',amount:'+0',reason:'範例資料',balance:'0',type:'create'}];
    if(statusFlag==='In') logs.push({time:t,action:'收料入庫 Receive',amount:'+'+currentQty.toFixed(1),reason:'採購到貨',balance:currentQty.toFixed(1),type:'in'});
    const item={
      code:finalCode, categoryCode:cat, catName:catMap[cat]||cat,
      typeName:cat==='O'?'大底 Outsole':(MAT[cat]?.types[type]||type),
      specName, colorCode, currency, price,
      sizeEU:sizeEU||'', sizeUK:sizeUK||'', sizeUS:sizeUS||'', sizeJP:sizeJP||'',
      locationCode:locCode, locationName:locName,
      qty:currentQty.toFixed(1), unit,
      photo:photo||'', productName, vendor, brand, thickness:'',
      supplier:{name:vendor,contact:'',phone:'',origin:'台灣',lead:7,moq:''},
      date:new Date().toISOString().split('T')[0],
      status:statusFlag==='In'?'在庫 In':'待到貨 Pending',
      timestamp:Date.now()-count*60000, logs
    };
    inventory.push(item);
    saveItemDB(item);
  });
  inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
  localStorage.setItem('peihai_ex','1');
  renderAll();
  if(showAlert) alert('範例資料已載入');
}

// ─── Excel import preload (121 items from material_inventory_260528.xlsx) ──
function loadExcelImport(){
  if(!isAdmin()){ alert('只有管理員可以匯入'); return; }
  if(!confirm('這會載入 121 筆材料資料（來自 Excel 匯入），確定嗎？')) return;
  const excelItems = [{"code": "IMP-0001-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "1.2", "colorCode": "昆謠土", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-1", "locationName": "主倉 A — 第 1 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 1.2 / 昆謠土", "vendor": "言成", "brand": "GEOX", "thickness": "1.2", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976849400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 1.2 / 昆謠土", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0002-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "1.2", "colorCode": "金黃", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-2", "locationName": "主倉 A — 第 2 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 1.2 / 金黃", "vendor": "言成", "brand": "GEOX", "thickness": "1.2", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976848400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 1.2 / 金黃", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0003-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "1.2", "colorCode": "銀鐵灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-3", "locationName": "主倉 A — 第 3 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 1.2 / 銀鐵灰", "vendor": "言成", "brand": "GEOX", "thickness": "1.2", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976847400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 1.2 / 銀鐵灰", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0004-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "0.6", "colorCode": "金黃", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-4", "locationName": "主倉 A — 第 4 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 0.6 / 金黃", "vendor": "言成", "brand": "GEOX", "thickness": "0.6", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976846400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 0.6 / 金黃", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0005-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "0.6", "colorCode": "銀鐵灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-5", "locationName": "主倉 A — 第 5 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 0.6 / 銀鐵灰", "vendor": "言成", "brand": "GEOX", "thickness": "0.6", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976845400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 0.6 / 銀鐵灰", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0006-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "超籤", "specName": "0.6", "colorCode": "昆謠土", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-6", "locationName": "主倉 A — 第 6 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超籤 / FW27 超件 0.6 / 昆謠土", "vendor": "言成", "brand": "GEOX", "thickness": "0.6", "supplier": {"name": "言成", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976844400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 超件 0.6 / 昆謠土", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0007-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "金黃", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-24", "locationName": "主倉 D — 第 24 格", "qty": "17.5", "unit": "SF", "photo": "", "productName": "反毛皮 / FW27 反毛皮 1.2-1.4 / 金黃", "vendor": "峰昌", "brand": "GEOX", "thickness": "1.2-1.4", "supplier": {"name": "峰昌", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976843400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+17.5", "reason": "Excel 匯入 / FW27 反毛皮 1.2-1.4 / 金黃", "balance": "17.5", "type": "in"}]}, {"code": "IMP-0008-GEOX", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "銀鐵灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-25", "locationName": "主倉 D — 第 25 格", "qty": "30.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW27 反毛皮 1.2-1.4 / 銀鐵灰", "vendor": "峰昌", "brand": "GEOX", "thickness": "1.2-1.4", "supplier": {"name": "峰昌", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976842400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+30.0", "reason": "Excel 匯入 / FW27 反毛皮 1.2-1.4 / 銀鐵灰", "balance": "30.0", "type": "in"}]}, {"code": "IMP-0009-NAVA", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "MR811", "specName": "1.24", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "T-1", "locationName": "主倉 T — 第 1 格", "qty": "59.4", "unit": "PRS", "photo": "", "productName": "MR811 / SS27 SMS MR811 / 本白 / 1.24", "vendor": "通慧蓮", "brand": "NAVAT?", "thickness": "1.24", "supplier": {"name": "通慧蓮", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976841400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+59.4", "reason": "Excel 匯入 / SS27 SMS MR811 / 本白 / 1.24", "balance": "59.4", "type": "in"}]}, {"code": "IMP-0010-NAVA", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "MR811", "specName": "1.24", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "T-18", "locationName": "主倉 T — 第 18 格", "qty": "10.6", "unit": "PRS", "photo": "", "productName": "MR811 / SS27 SMS MR811 / 本白 數(45?)", "vendor": "通慧蓮", "brand": "NAVAT?", "thickness": "1.24", "supplier": {"name": "通慧蓮", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976840400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+10.6", "reason": "Excel 匯入 / SS27 SMS MR811 / 本白 數(45?)", "balance": "10.6", "type": "in"}]}, {"code": "IMP-0011-GEOX", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "GEOXPO PU", "specName": "PU", "colorCode": "#04 深咖啡", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-1", "locationName": "主倉 K — 第 1 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "GEOXPO PU / FW27 GEOXPO PU #04", "vendor": "易尚", "brand": "GEOX", "thickness": "PU", "supplier": {"name": "易尚", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976839400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 GEOXPO PU #04", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0012-GEOX", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "GEOXPO PU", "specName": "PU", "colorCode": "#08 橘棕", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-2", "locationName": "主倉 K — 第 2 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "GEOXPO PU / FW27 GEOXPO PU YS-G342 #08", "vendor": "易尚", "brand": "GEOX", "thickness": "PU", "supplier": {"name": "易尚", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976838400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / FW27 GEOXPO PU YS-G342 #08", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0013-GEOX", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "超纖布面", "specName": "", "colorCode": "#08 黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-3", "locationName": "主倉 K — 第 3 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超纖布面 / HJ435636-08 超纖布面", "vendor": "宏錦", "brand": "GEOX", "thickness": "", "supplier": {"name": "宏錦", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976837400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HJ435636-08 超纖布面", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0014-GEOX", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "超纖布面", "specName": "", "colorCode": "#001 花高", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-4", "locationName": "主倉 K — 第 4 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "超纖布面 / HJ435335-001 超纖布面", "vendor": "宏錦", "brand": "GEOX", "thickness": "", "supplier": {"name": "宏錦", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976836400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HJ435335-001 超纖布面", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0015-OAKL", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "濕式PU", "specName": "0.7", "colorCode": "白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-22", "locationName": "主倉 K — 第 22 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "濕式PU / OKYCP2 無氟防虹吸 濕式PU", "vendor": "錦利", "brand": "OAKLEY", "thickness": "0.7", "supplier": {"name": "錦利", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976835400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / OKYCP2 無氟防虹吸 濕式PU", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0016-OAKL", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "濕式PU", "specName": "0.7", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-33", "locationName": "主倉 K — 第 33 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "濕式PU / OKYCP2 無氟防虹吸 濕式PU", "vendor": "錦利", "brand": "OAKLEY", "thickness": "0.7", "supplier": {"name": "錦利", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976834400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / OKYCP2 無氟防虹吸 濕式PU", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0017-OAKL", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "PU", "specName": "1.35", "colorCode": "亮黃 11-0601TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-34", "locationName": "主倉 K — 第 34 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "PU / OKYCP2 仿超件PU 防虹吸 +6060", "vendor": "錦利", "brand": "OAKLEY", "thickness": "1.35", "supplier": {"name": "錦利", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976833400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / OKYCP2 仿超件PU 防虹吸 +6060", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0018-OAKL", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "PU", "specName": "1.2", "colorCode": "亮白 11-0601TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-35", "locationName": "主倉 K — 第 35 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "PU / OKYCP2 SS-GH9460紋 無氟防虹吸", "vendor": "錦利", "brand": "OAKLEY", "thickness": "1.2", "supplier": {"name": "錦利", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976832400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / OKYCP2 SS-GH9460紋 無氟防虹吸", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0019-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 布", "specName": "", "colorCode": "黑 庫存", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-16", "locationName": "主倉 G — 第 16 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "CP2 布 / CP2 啞光萊克布貼佳績 無氟防虹吸", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976831400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / CP2 啞光萊克布貼佳績 無氟防虹吸", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0020-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 布", "specName": "", "colorCode": "花固石灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-17", "locationName": "主倉 G — 第 17 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 布 / CP2 錦綸啞光萊克布貼佳績", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976830400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / CP2 錦綸啞光萊克布貼佳績", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0021-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 布", "specName": "", "colorCode": "花固石灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-18", "locationName": "主倉 G — 第 18 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 布 / CP2 錦綸啞光萊克布貼 SBR", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976829400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / CP2 錦綸啞光萊克布貼 SBR", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0022-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 布", "specName": "", "colorCode": "花固石灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-19", "locationName": "主倉 G — 第 19 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 布 / CP2 錦綸啞光萊克布", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976828400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / CP2 錦綸啞光萊克布", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0023-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 cloth lycra", "specName": "260g", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-1", "locationName": "主倉 G — 第 1 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 cloth lycra / cloth lycra 啞光萊克布 260g 佳績布", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "260g", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976827400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / cloth lycra 啞光萊克布 260g 佳績布", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0024-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 cloth lycra", "specName": "260g / 2mm SBR", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-2", "locationName": "主倉 G — 第 2 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 cloth lycra / cloth lycra 啞光萊克布 260g 2mm SBR", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "260g / 2mm SBR", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976826400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / cloth lycra 啞光萊克布 260g 2mm SBR", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0025-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 cloth lycra", "specName": "260g / 1mm SBR", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-3", "locationName": "主倉 G — 第 3 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 cloth lycra / cloth lycra 啞光萊克布 260g 1mm SBR", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "260g / 1mm SBR", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976825400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / cloth lycra 啞光萊克布 260g 1mm SBR", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0026-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 Mesh Air", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-4", "locationName": "主倉 G — 第 4 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 Mesh Air / Mesh Air XR-DS8036 無氟防虹吸", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976824400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / Mesh Air XR-DS8036 無氟防虹吸", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0027-OAKL", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "CP2 XR-430", "specName": "", "colorCode": "水晶米", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-13", "locationName": "主倉 G — 第 13 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 XR-430 / XR-430 +6060", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976823400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / XR-430 +6060", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0028-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 萊卡", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-14", "locationName": "主倉 G — 第 14 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "CP2 萊卡 / 萊卡 防虹吸", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976822400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / 萊卡 防虹吸", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0029-OAKL", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CP2 布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-15", "locationName": "主倉 G — 第 15 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "CP2 布 / 啞光萊克布貼佳績 無氟防虹吸", "vendor": "興瑞", "brand": "OAKLEY", "thickness": "", "supplier": {"name": "興瑞", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976821400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / 啞光萊克布貼佳績 無氟防虹吸", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0030-AXM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "本白 11-0602TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-36", "locationName": "主倉 K — 第 36 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / BAK123 承載布 本白", "vendor": "歐高", "brand": "AXM", "thickness": "", "supplier": {"name": "歐高", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976820400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / BAK123 承載布 本白", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0031-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "TS115-1 +6060", "specName": "", "colorCode": "義水灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-20", "locationName": "主倉 G — 第 20 格", "qty": "4.0", "unit": "YD", "photo": "", "productName": "TS115-1 +6060 / SS27 SMS TS115-1 +6060", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976819400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / SS27 SMS TS115-1 +6060", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0032-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "TS115-1 +6060", "specName": "", "colorCode": "三海青灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-21", "locationName": "主倉 G — 第 21 格", "qty": "4.0", "unit": "YD", "photo": "", "productName": "TS115-1 +6060 / FW26 JD SMS TS115-1 +6060", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976818400, "note": "季別: FW26 JD SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / FW26 JD SMS TS115-1 +6060", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0033-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "TS115-1 +6060", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-22", "locationName": "主倉 G — 第 22 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "TS115-1 +6060 / SS27 SMS + JD SMS TS115-1 +6060", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976817400, "note": "季別: SS27 SMS + JD SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / SS27 SMS + JD SMS TS115-1 +6060", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0034-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "XT028 斜菱形", "specName": "", "colorCode": "義水灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-23", "locationName": "主倉 G — 第 23 格", "qty": "4.0", "unit": "YD", "photo": "", "productName": "XT028 斜菱形 / XT028斜菱形 生活布+佳績+6060", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976816400, "note": "季別: SS27 SMS + JD SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / XT028斜菱形 生活布+佳績+6060", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0035-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "XT028 斜菱形", "specName": "", "colorCode": "牛調", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-24", "locationName": "主倉 G — 第 24 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "XT028 斜菱形 / FW24 SMS XT028斜菱形 生活布+佳績+6060", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976815400, "note": "季別: FW24 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / FW24 SMS XT028斜菱形 生活布+佳績+6060", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0036-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "XT058 + MT115FR", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-25", "locationName": "主倉 G — 第 25 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "XT058 + MT115FR / XT058 + 佳績 +6060 + MT115FR", "vendor": "信泰", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976814400, "note": "季別: SS27 SMS + JD SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / XT058 + 佳績 +6060 + MT115FR", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0037-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-1", "locationName": "主倉 D — 第 1 格", "qty": "20.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW26 SMS SS27 SMS 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976813400, "note": "季別: FW26 SMS / SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+20.0", "reason": "Excel 匯入 / FW26 SMS SS27 SMS 反毛皮", "balance": "20.0", "type": "in"}]}, {"code": "IMP-0038-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "多麼咖啡", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-2", "locationName": "主倉 D — 第 2 格", "qty": "28.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW26 SMS AW26 Q4 SMS 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976812400, "note": "季別: FW26 SMS / AW26 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+28.0", "reason": "Excel 匯入 / FW26 SMS AW26 Q4 SMS 反毛皮", "balance": "28.0", "type": "in"}]}, {"code": "IMP-0039-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "火炬灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-3", "locationName": "主倉 D — 第 3 格", "qty": "30.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS26 SMS 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976811400, "note": "季別: SS26 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+30.0", "reason": "Excel 匯入 / SS26 SMS 反毛皮", "balance": "30.0", "type": "in"}]}, {"code": "IMP-0040-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "亮土", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-4", "locationName": "主倉 D — 第 4 格", "qty": "40.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27-JD-SMS 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976810400, "note": "季別: SS27-JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+40.0", "reason": "Excel 匯入 / SS27-JD-SMS 反毛皮", "balance": "40.0", "type": "in"}]}, {"code": "IMP-0041-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "米白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-5", "locationName": "主倉 D — 第 5 格", "qty": "62.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS26 SMS SS27 JP SMS 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976809400, "note": "季別: SS26 SMS / SS27 JP SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+62.0", "reason": "Excel 匯入 / SS26 SMS SS27 JP SMS 反毛皮", "balance": "62.0", "type": "in"}]}, {"code": "IMP-0042-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "駝色", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-6", "locationName": "主倉 D — 第 6 格", "qty": "35.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27-JD-SMS HOFF 反毛皮", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976808400, "note": "季別: SS27-JD-SMS / HOFF ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+35.0", "reason": "Excel 匯入 / SS27-JD-SMS HOFF 反毛皮", "balance": "35.0", "type": "in"}]}, {"code": "IMP-0043-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "熱紙反毛", "specName": "1.2-1.4", "colorCode": "信風灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-7", "locationName": "主倉 D — 第 7 格", "qty": "40.0", "unit": "SF", "photo": "", "productName": "熱紙反毛 / FW26 SMS 熱紙反毛", "vendor": "侁昕", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "侁昕", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976807400, "note": "季別: FW26 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+40.0", "reason": "Excel 匯入 / FW26 SMS 熱紙反毛", "balance": "40.0", "type": "in"}]}, {"code": "IMP-0044-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "茶褐", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-8", "locationName": "主倉 D — 第 8 格", "qty": "60.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW26 SMS AW26 Q4 SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.2-1.4", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976806400, "note": "季別: FW26 SMS / AW26 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+60.0", "reason": "Excel 匯入 / FW26 SMS AW26 Q4 SMS 反毛皮", "balance": "60.0", "type": "in"}]}, {"code": "IMP-0045-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "高絲絨反毛皮", "specName": "1.24", "colorCode": "石木色", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-9", "locationName": "主倉 D — 第 9 格", "qty": "8.0", "unit": "SF", "photo": "", "productName": "高絲絨反毛皮 / SS27 JD-SMS 高絲絨反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976805400, "note": "季別: SS27 JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+8.0", "reason": "Excel 匯入 / SS27 JD-SMS 高絲絨反毛皮", "balance": "8.0", "type": "in"}]}, {"code": "IMP-0046-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "石木岩", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-10", "locationName": "主倉 D — 第 10 格", "qty": "10.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27 SMS SS27 PROTO 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976804400, "note": "季別: SS27 SMS / SS27 PROTO ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+10.0", "reason": "Excel 匯入 / SS27 SMS SS27 PROTO 反毛皮", "balance": "10.0", "type": "in"}]}, {"code": "IMP-0047-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "橄欖布色", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-11", "locationName": "主倉 D — 第 11 格", "qty": "6.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27 JD SMS HOFF 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976803400, "note": "季別: SS27 JD SMS / HOFF ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / SS27 JD SMS HOFF 反毛皮", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0048-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-12", "locationName": "主倉 D — 第 12 格", "qty": "56.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW15 Q4 SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976802400, "note": "季別: FW15 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+56.0", "reason": "Excel 匯入 / FW15 Q4 SMS 反毛皮", "balance": "56.0", "type": "in"}]}, {"code": "IMP-0049-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "義水灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-13", "locationName": "主倉 D — 第 13 格", "qty": "19.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS26 AW26 SS27 PROTO 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976801400, "note": "季別: SS26 SMS / AW26 Q4 SMS / SS27 PROTO ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+19.0", "reason": "Excel 匯入 / SS26 AW26 SS27 PROTO 反毛皮", "balance": "19.0", "type": "in"}]}, {"code": "IMP-0050-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "高絲絨反毛皮", "specName": "1.24", "colorCode": "棕色 CRM", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-14", "locationName": "主倉 D — 第 14 格", "qty": "65.0", "unit": "SF", "photo": "", "productName": "高絲絨反毛皮 / SS27 SMS 高絲絨反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976800400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+65.0", "reason": "Excel 匯入 / SS27 SMS 高絲絨反毛皮", "balance": "65.0", "type": "in"}]}, {"code": "IMP-0051-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-5", "locationName": "主倉 K — 第 5 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330 T9紋", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976799400, "note": "季別: FW26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2330 T9紋", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0052-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-6", "locationName": "主倉 K — 第 6 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330 T9紋", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976798400, "note": "季別: FW26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2330 T9紋", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0053-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "PU", "specName": "PU 1.4", "colorCode": "灰光黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-7", "locationName": "主倉 K — 第 7 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "PU / RX-SM", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976797400, "note": "季別: ", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / RX-SM", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0054-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "PU", "specName": "PU 1.2", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-8", "locationName": "主倉 K — 第 8 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "PU / HW2321-1", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.2", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976796400, "note": "季別: ", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HW2321-1", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0055-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU", "colorCode": "三海青灰 19-0201TPG", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-9", "locationName": "主倉 K — 第 9 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330 T9紋 HW2330-28", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976795400, "note": "季別: ", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HW2330 T9紋 HW2330-28", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0056-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "高面平紋 PU", "specName": "PU 1.4", "colorCode": "深灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-10", "locationName": "主倉 K — 第 10 格", "qty": "1.0", "unit": "YD", "photo": "", "productName": "高面平紋 PU / HW2422-1", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976794400, "note": "季別: SS26 SMS / AW26 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+1.0", "reason": "Excel 匯入 / HW2422-1", "balance": "1.0", "type": "in"}]}, {"code": "IMP-0057-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "冰霜灰 14TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-11", "locationName": "主倉 K — 第 11 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2370 T9紋 HW2370-58", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976793400, "note": "季別: FW25 PROTO", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2370 T9紋 HW2370-58", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0058-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "咖啡力", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-15", "locationName": "主倉 D — 第 15 格", "qty": "80.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27 JD SMS HOFF 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976792400, "note": "季別: SS27 JD SMS / HOFF ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+80.0", "reason": "Excel 匯入 / SS27 JD SMS HOFF 反毛皮", "balance": "80.0", "type": "in"}]}, {"code": "IMP-0059-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "飛棕 18-1304TPG", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-16", "locationName": "主倉 D — 第 16 格", "qty": "4.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW26 Q4 SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976791400, "note": "季別: FW26 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / FW26 Q4 SMS 反毛皮", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0060-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "冰灰色", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-17", "locationName": "主倉 D — 第 17 格", "qty": "0.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27-JD-SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "已耗盡 Empty", "timestamp": 1779976790400, "note": "季別: SS27-JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+0.0", "reason": "Excel 匯入 / SS27-JD-SMS 反毛皮", "balance": "0.0", "type": "in"}]}, {"code": "IMP-0061-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "牛39", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-18", "locationName": "主倉 D — 第 18 格", "qty": "2.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27-JD-SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976789400, "note": "季別: SS27-JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / SS27-JD-SMS 反毛皮", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0062-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "牛調", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-19", "locationName": "主倉 D — 第 19 格", "qty": "0.0", "unit": "SF", "photo": "", "productName": "反毛皮 / SS27 SMS 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "已耗盡 Empty", "timestamp": 1779976788400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+0.0", "reason": "Excel 匯入 / SS27 SMS 反毛皮", "balance": "0.0", "type": "in"}]}, {"code": "IMP-0063-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.24", "colorCode": "三海青灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-20", "locationName": "主倉 D — 第 20 格", "qty": "60.0", "unit": "SF", "photo": "", "productName": "反毛皮 / FW26 SS27 AW26 Q4 反毛皮", "vendor": "", "brand": "CRUYFF", "thickness": "1.24", "supplier": {"name": "", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976787400, "note": "季別: FW26 SMS / SS27 SMS / AW26 Q4 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+60.0", "reason": "Excel 匯入 / FW26 SS27 AW26 Q4 反毛皮", "balance": "60.0", "type": "in"}]}, {"code": "IMP-0064-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "101白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-12", "locationName": "主倉 K — 第 12 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330 T9紋 HW2330-34", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976786400, "note": "季別: SS26 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / HW2330 T9紋 HW2330-34", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0065-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "無折痕霧面 PU", "specName": "PU 0.8", "colorCode": "牛調", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-13", "locationName": "主倉 K — 第 13 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "無折痕霧面 PU / HW1719-117K HW1719-117", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 0.8", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976785400, "note": "季別: SS26 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW1719-117K HW1719-117", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0066-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "義水灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-14", "locationName": "主倉 K — 第 14 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330-36 HW2330-T9", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976784400, "note": "季別: SS26 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2330-36 HW2330-T9", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0067-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.2", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-15", "locationName": "主倉 K — 第 15 格", "qty": "3.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2445-1 銘石紋", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.2", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976783400, "note": "季別: SS26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / HW2445-1 銘石紋", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0068-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "破皮紋 PU", "specName": "PU 0.8", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-16", "locationName": "主倉 K — 第 16 格", "qty": "7.0", "unit": "YD", "photo": "", "productName": "破皮紋 PU / HW1803-22", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 0.8", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976782400, "note": "季別: SS26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+7.0", "reason": "Excel 匯入 / HW1803-22", "balance": "7.0", "type": "in"}]}, {"code": "IMP-0069-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "三海青灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-17", "locationName": "主倉 K — 第 17 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2330-28", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976781400, "note": "季別: SS27 JD", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2330-28", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0070-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "無折痕霧面 PU", "specName": "PU 0.8", "colorCode": "101白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-18", "locationName": "主倉 K — 第 18 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "無折痕霧面 PU / HW1719-177", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 0.8", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976780400, "note": "季別: SS26 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW1719-177", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0071-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-19", "locationName": "主倉 K — 第 19 格", "qty": "6.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2452-1", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976779400, "note": "季別: SS27-JD", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW2452-1", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0072-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "經紋超件 PU", "specName": "1.2", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-20", "locationName": "主倉 K — 第 20 格", "qty": "9.0", "unit": "YD", "photo": "", "productName": "經紋超件 PU / HW2444 T9 HW2444-1", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.2", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976778400, "note": "季別: SS26 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+9.0", "reason": "Excel 匯入 / HW2444 T9 HW2444-1", "balance": "9.0", "type": "in"}]}, {"code": "IMP-0073-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "三海灰 18-5203TPG", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-21", "locationName": "主倉 K — 第 21 格", "qty": "5.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2370 U41紋 HW2370-28", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976777400, "note": "季別: SS27-JD-SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+5.0", "reason": "Excel 匯入 / HW2370 U41紋 HW2370-28", "balance": "5.0", "type": "in"}]}, {"code": "IMP-0074-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "碎皮紋 PU", "specName": "PU 0.8", "colorCode": "奶油", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-22", "locationName": "主倉 K — 第 22 格", "qty": "6.0", "unit": "YD", "photo": "", "productName": "碎皮紋 PU / HW1803 HW1803-24", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 0.8", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976776400, "note": "季別: SS27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW1803 HW1803-24", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0075-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "無折痕霧面 PU", "specName": "PU 0.8", "colorCode": "三海青灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-23", "locationName": "主倉 K — 第 23 格", "qty": "4.0", "unit": "YD", "photo": "", "productName": "無折痕霧面 PU / HW1719-199", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 0.8", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976775400, "note": "季別: SS27 JD SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / HW1719-199", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0076-CRUY", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布標絨", "specName": "1.2", "colorCode": "義水灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-24", "locationName": "主倉 K — 第 24 格", "qty": "9.0", "unit": "YD", "photo": "", "productName": "布標絨 / PROTO HW1810", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.2", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976774400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+9.0", "reason": "Excel 匯入 / PROTO HW1810", "balance": "9.0", "type": "in"}]}, {"code": "IMP-0077-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "白色", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "K-25", "locationName": "主倉 K — 第 25 格", "qty": "6.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW3377 CM1紋 HW3377-JF", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976773400, "note": "季別: SS27 JD SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW3377 CM1紋 HW3377-JF", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0078-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "優面軟件 PU", "specName": "PU 1.2", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "A-7", "locationName": "主倉 A — 第 7 格", "qty": "14.0", "unit": "YD", "photo": "", "productName": "優面軟件 PU / HW1717-1", "vendor": "華博", "brand": "CRUYFF", "thickness": "PU 1.2", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976772400, "note": "季別: FW25 PROTO ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+14.0", "reason": "Excel 匯入 / HW1717-1", "balance": "14.0", "type": "in"}]}, {"code": "IMP-0079-CRUY", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-26", "locationName": "主倉 G — 第 26 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / ZM-BF2303/401 + 3M K329 + 6W", "vendor": "仲銘", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "仲銘", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976771400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / ZM-BF2303/401 + 3M K329 + 6W", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0080-CRUY", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "深灰 18-5203TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-27", "locationName": "主倉 G — 第 27 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / ZM-BF2303/401 + ZX K329 + 6W", "vendor": "仲銘", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "仲銘", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976770400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / ZM-BF2303/401 + ZX K329 + 6W", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0081-CRUY", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "G-28", "locationName": "主倉 G — 第 28 格", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / ZM-BF2303/401 + ZX K329 + 6W", "vendor": "仲銘", "brand": "CRUYFF", "thickness": "", "supplier": {"name": "仲銘", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976769400, "note": "季別:  ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / ZM-BF2303/401 + ZX K329 + 6W", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0082-CRUY", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "GW細紋 + 油1D", "specName": "6W", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "I-2", "locationName": "主倉 I — 第 2 格", "qty": "0.0", "unit": "YD", "photo": "", "productName": "GW細紋 + 油1D / SS27 SMS GW細紋 + 油1D", "vendor": "集各", "brand": "CRUYFF", "thickness": "6W", "supplier": {"name": "集各", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "已耗盡 Empty", "timestamp": 1779976768400, "note": "季別: SS27 SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+0.0", "reason": "Excel 匯入 / SS27 SMS GW細紋 + 油1D", "balance": "0.0", "type": "in"}]}, {"code": "IMP-0083-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "奶油", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-1", "locationName": "主倉 F — 第 1 格", "qty": "43.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-248", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976767400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+43.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-248", "balance": "43.0", "type": "in"}]}, {"code": "IMP-0084-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "淺冰灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-2", "locationName": "主倉 F — 第 2 格", "qty": "40.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-1003", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976766400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+40.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-1003", "balance": "40.0", "type": "in"}]}, {"code": "IMP-0085-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "101白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-3", "locationName": "主倉 F — 第 3 格", "qty": "6.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-182", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976765400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-182", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0086-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "奶油", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-4", "locationName": "主倉 F — 第 4 格", "qty": "8.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-248", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976764400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+8.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-248", "balance": "8.0", "type": "in"}]}, {"code": "IMP-0087-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-5", "locationName": "主倉 F — 第 5 格", "qty": "6.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-399", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976763400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-399", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0088-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "本白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-6", "locationName": "主倉 F — 第 6 格", "qty": "7.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-399 沖孔", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976762400, "note": "季別: SS27-JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+7.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-399 沖孔", "balance": "7.0", "type": "in"}]}, {"code": "IMP-0089-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "101白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-7", "locationName": "主倉 F — 第 7 格", "qty": "10.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20打型", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976761400, "note": "季別: SS27-JD-SMS ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+10.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20打型", "balance": "10.0", "type": "in"}]}, {"code": "IMP-0090-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-8", "locationName": "主倉 F — 第 8 格", "qty": "15.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-194", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976760400, "note": "季別: FW25 PROTO", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+15.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-194", "balance": "15.0", "type": "in"}]}, {"code": "IMP-0091-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "三隱青灰 19-0201TPG", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-9", "locationName": "主倉 F — 第 9 格", "qty": "7.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-176", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976759400, "note": "季別: FW26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+7.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-176", "balance": "7.0", "type": "in"}]}, {"code": "IMP-0092-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "研磨咖啡", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-10", "locationName": "主倉 F — 第 10 格", "qty": "12.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-176", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976758400, "note": "季別: SS27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+12.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-176", "balance": "12.0", "type": "in"}]}, {"code": "IMP-0093-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "三隱青灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-11", "locationName": "主倉 F — 第 11 格", "qty": "0.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-176", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "已耗盡 Empty", "timestamp": 1779976757400, "note": "季別: SS27-JD ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+0.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-176", "balance": "0.0", "type": "in"}]}, {"code": "IMP-0094-CRUY", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "椰皮", "specName": "1.3-1.5", "colorCode": "海藻綠 16-6008TPG", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "F-12", "locationName": "主倉 F — 第 12 格", "qty": "8.0", "unit": "SF", "photo": "", "productName": "椰皮 / HW20 MAKOU B100 LUX紋 HW20-1682", "vendor": "華博", "brand": "CRUYFF", "thickness": "1.3-1.5", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976756400, "note": "季別: SS27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+8.0", "reason": "Excel 匯入 / HW20 MAKOU B100 LUX紋 HW20-1682", "balance": "8.0", "type": "in"}]}, {"code": "IMP-0095-FRAT", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "JUSTY MARINE", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-21", "locationName": "主倉 D — 第 21 格", "qty": "40.0", "unit": "SF", "photo": "", "productName": "反毛皮 / JUSTY MARINE 高手感", "vendor": "傅達", "brand": "FRATELLI ROSSETTI", "thickness": "1.2-1.4", "supplier": {"name": "傅達", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976755400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+40.0", "reason": "Excel 匯入 / JUSTY MARINE 高手感", "balance": "40.0", "type": "in"}]}, {"code": "IMP-0096-FRAT", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "WHITE", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-22", "locationName": "主倉 D — 第 22 格", "qty": "34.0", "unit": "SF", "photo": "", "productName": "反毛皮 / WHITE 高手感", "vendor": "傅達", "brand": "FRATELLI ROSSETTI", "thickness": "1.2-1.4", "supplier": {"name": "傅達", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976754400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+34.0", "reason": "Excel 匯入 / WHITE 高手感", "balance": "34.0", "type": "in"}]}, {"code": "IMP-0097-FRAT", "categoryCode": "L", "catName": "皮類 Leather", "typeName": "反毛皮", "specName": "1.2-1.4", "colorCode": "棕", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "D-23", "locationName": "主倉 D — 第 23 格", "qty": "40.0", "unit": "SF", "photo": "", "productName": "反毛皮 / 棕 高手感", "vendor": "傅達", "brand": "FRATELLI ROSSETTI", "thickness": "1.2-1.4", "supplier": {"name": "傅達", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976753400, "note": "季別: SS27 SMS", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+40.0", "reason": "Excel 匯入 / 棕 高手感", "balance": "40.0", "type": "in"}]}, {"code": "IMP-0098-未辨識", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "地球灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / JC1749 + PU", "vendor": "容全", "brand": "未辨識", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976752400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC1749 + PU", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0099-未辨識", "categoryCode": "T", "catName": "布類 Textile", "typeName": "毛巾布", "specName": "", "colorCode": "純卡賣 16-1103TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "4.0", "unit": "YD", "photo": "", "productName": "毛巾布 / JC1940", "vendor": "容全", "brand": "未辨識", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976751400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / JC1940", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0100-未辨識", "categoryCode": "T", "catName": "布類 Textile", "typeName": "毛巾布", "specName": "", "colorCode": "103白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "毛巾布 / JC1940", "vendor": "容全", "brand": "未辨識", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976750400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC1940", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0101-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "毛巾布", "specName": "", "colorCode": "地普咖 19-0910TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "1.0", "unit": "YD", "photo": "", "productName": "毛巾布 / JC2161 貼38G", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976749400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+1.0", "reason": "Excel 匯入 / JC2161 貼38G", "balance": "1.0", "type": "in"}]}, {"code": "IMP-0102-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "地普咖 19-0910TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / JC2270 貼佳績", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976748400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC2270 貼佳績", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0103-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "1.0", "unit": "YD", "photo": "", "productName": "布 / JC1749 貼佳績", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976747400, "note": "季別: SS27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+1.0", "reason": "Excel 匯入 / JC1749 貼佳績", "balance": "1.0", "type": "in"}]}, {"code": "IMP-0104-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "四葉綠 18-0420TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "4.0", "unit": "YD", "photo": "", "productName": "布 / JC2273 格子布+PU+佳績", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976746400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / JC2273 格子布+PU+佳績", "balance": "4.0", "type": "in"}]}, {"code": "IMP-0105-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "地球灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / JC1749 貼佳績", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976745400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC1749 貼佳績", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0106-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "白", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / JC0121-2", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976744400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC0121-2", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0107-EAM", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "布 / JC2302", "vendor": "容全", "brand": "EAM", "thickness": "", "supplier": {"name": "容全", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976743400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / JC2302", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0108-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "1/2斜紋布", "specName": "", "colorCode": "原胚", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "1/2斜紋布 / DQ-XC9040 SS26013-1", "vendor": "大立發", "brand": "EA7", "thickness": "", "supplier": {"name": "大立發", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976742400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / DQ-XC9040 SS26013-1", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0109-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "防水膜", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "防水膜 / 啞光空氣襪套防水膜", "vendor": "COSMOSTAR", "brand": "EA7", "thickness": "", "supplier": {"name": "COSMOSTAR", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976741400, "note": "季別: SS27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / 啞光空氣襪套防水膜", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0110-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "黑 小荔枝紋", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "3.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2385 T9 HW2385-1", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976740400, "note": "季別: FW27 RESEARCH ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / HW2385 T9 HW2385-1", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0111-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2374-1 B100紋", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976739400, "note": "季別: FW26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HW2374-1 B100紋", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0112-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "地被咖 19-0910TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2374 B100紋 HW2374-80", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976738400, "note": "季別: FW27 RESEARCH", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW2374 B100紋 HW2374-80", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0113-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "純卡賣 16-1103TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2385 T9 HW2385-48 小荔枝紋", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976737400, "note": "季別: FW27 RESEARCH ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW2385 T9 HW2385-48 小荔枝紋", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0114-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "純卡賣 16-1103TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "6.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2370 T9紋 HW2370-110", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976736400, "note": "季別: FW27 RESEARCH", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+6.0", "reason": "Excel 匯入 / HW2370 T9紋 HW2370-110", "balance": "6.0", "type": "in"}]}, {"code": "IMP-0115-EA7", "categoryCode": "S", "catName": "人造皮 Synthetic", "typeName": "仿超件 PU", "specName": "PU 1.4", "colorCode": "銀", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "仿超件 PU / HW2370 T9紋 HW2370-64", "vendor": "華博", "brand": "EA7", "thickness": "PU 1.4", "supplier": {"name": "華博", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976735400, "note": "季別: FW26", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / HW2370 T9紋 HW2370-64", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0116-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "3.0", "unit": "YD", "photo": "", "productName": "布 / SKY-030", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976734400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / SKY-030", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0117-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "3.0", "unit": "YD", "photo": "", "productName": "布 / 彈力提花格布", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976733400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+3.0", "reason": "Excel 匯入 / 彈力提花格布", "balance": "3.0", "type": "in"}]}, {"code": "IMP-0118-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "地球灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / CH019", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976732400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / CH019", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0119-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "CH019布", "specName": "", "colorCode": "地球灰", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "1.0", "unit": "YD", "photo": "", "productName": "CH019布 / CH019 貼小BK", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976731400, "note": "季別: FW27 ⚠️需確認", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+1.0", "reason": "Excel 匯入 / CH019 貼小BK", "balance": "1.0", "type": "in"}]}, {"code": "IMP-0120-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "純卡賣 16-1103TPX", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "2.0", "unit": "YD", "photo": "", "productName": "布 / JC1942", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976730400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+2.0", "reason": "Excel 匯入 / JC1942", "balance": "2.0", "type": "in"}]}, {"code": "IMP-0121-EA7", "categoryCode": "T", "catName": "布類 Textile", "typeName": "布", "specName": "", "colorCode": "黑", "currency": "TWD", "price": "0", "sizeEU": "", "sizeUK": "", "sizeUS": "", "sizeJP": "", "locationCode": "TEMP-1", "locationName": "TEMP — 暫放區", "qty": "4.0", "unit": "YD", "photo": "", "productName": "布 / SY-129", "vendor": "信泰", "brand": "EA7", "thickness": "", "supplier": {"name": "信泰", "contact": "", "phone": "", "origin": "台灣", "lead": 0, "moq": ""}, "date": "2026-05-28", "status": "在庫 In", "timestamp": 1779976729400, "note": "季別: FW27", "logs": [{"time": "2026/05/28 00:00", "action": "匯入建檔 Import", "amount": "+4.0", "reason": "Excel 匯入 / SY-129", "balance": "4.0", "type": "in"}]}];
  // Merge: keep existing, add new ones that don't conflict
  const existingCodes = new Set(inventory.map(i=>i.code));
  const newItems = excelItems.filter(i=>!existingCodes.has(i.code));
  inventory = [...newItems, ...inventory];
  inventory.sort((a,b)=>(b.timestamp||0)-(a.timestamp||0));
  replaceInDB(inventory, ()=>{ renderAll(); alert('✅ 匯入完成！共載入 '+newItems.length+' 筆材料。'); }, true);
}

// ─── Init ──────────────────────────────────────────────────────────────

window.onload=function(){
  initBrands();
  buildOutsoleSels();
  initDB();
  initSupabase();
  // Set today's date default
  const ibDate=document.getElementById('ibDate');
  if(ibDate) ibDate.valueAsDate=new Date();
};
