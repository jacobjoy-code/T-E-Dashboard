/**
 * T&E Command Center – Apps Script (v2)
 * ======================================
 * HOW TO SET UP:
 *   1. Paste this as Code.gs and the other file as Index.html in Extensions > Apps Script.
 *   2. Reload the sheet. Use the "T&E Dashboard" menu > "Set up access sheet" to create or refresh the access tab.
 *   3. Fill in the Email column on the Access sheet. Put your own email with Role = Admin.
 *   4. Run "Rebuild data now" from the menu and approve permissions.
 *   5. Deploy > New deployment > Web app. Execute as: Me.  Who has access: Anyone within your organisation.
 *   6. Share the web-app link. Each person sees only what the Access sheet allows.
 *
 * ACCESS RULES (from the "Access" sheet):
 *   Role = Admin        → sees every leader (full dashboard)
 *   Role = Leader       → sees only the leader(s) listed in column C
 *   No matching email   → "access denied" page
 *   The person who runs the script is always treated as Admin, even without a row.
 */

// ===================== SETTINGS =====================
const CONFIG = {
  sheets: {
    spend:     'T&E Q3 QTD',
    ta:        'TA',
    budget:    'Budget',
    hierarchy: 'Employee Hierachy',
    access:    'Access',
    checks:    'Dashboard Checks',
    emails:    'Message Emails',
    leaderIds: 'Leader IDs',
    messageLog:'Message Log'
  },
  budgetRegionMap: { ASN:'ASEAN', ASEAN:'ASEAN', India:'India', IND:'India' },
  aliases: {
    'Quy Nguyen Thi Kim':'Quy Nguyen', 'Nguyen Quy':'Quy Nguyen',
    'Sridhar Hariharasubramania':'Sridhar Hariharasubramanian',
    'Prathap G R':'Prathap GR', 'Gavin Barfield':'Gavin David Barfield'
  },
  notALeader: ['Employee Attrition'],
  inrTaPerDayThreshold: 2500,
  individualOnly: ['Arun Kumar Parameswaran','Mankiran Chowhan','Paul Carvouni','Arundhati Bhattacharya'],   // used if the Budget tab has no "Budget covers" column
  autoRefreshMinutes: 30,    // the dashboard rebuilds itself when its data is older than this (set 0 to turn off)
  editWatchSheets: ['T&E Q3 QTD','TA','Budget','Employee Hierachy']   // editing these marks the data as out of date
};
const CODE_VERSION = '2026.10.27-f';   // must match the version shown by the page; a mismatch shows a warning banner

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('T&E Dashboard')
    .addItem('Set up access sheet', 'setupAccessSheet')
    .addItem('Create missing tabs (run once)', 'createMissingTabs')
    .addItem('Import Workday (India & ASEAN)', 'importWorkdayFromTab')
    .addItem('Sync leaders from Workday', 'syncLeadersFromWorkday')
    .addItem('Rebuild data now', 'rebuildData')
    .addItem('Show web-app link', 'showLink')
    .addSeparator()
    .addItem('Turn on instant refresh when data is edited', 'installEditTrigger')
    .addItem('Rebuild every morning (7 AM)', 'installDailyTrigger')
    .addToUi();
}
function showLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'Share this link with leaders:\n\n' + url + '\n\nEach person must be listed on the Access sheet to open it.' : 'Deploy the script first:\nDeploy > New deployment > Web app');
}
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'rebuildData').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('rebuildData').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActive().toast('Dashboard data will rebuild every day at about 7 AM.','T&E Dashboard');
}

// ===================== ACCESS SHEET =====================
function setupAccessSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.sheets.access);
  const full = loadData_();
  const leaders = full ? full.meta.recipients : [];
  if (!sh) {
    sh = ss.insertSheet(CONFIG.sheets.access);
  } else {
    // preserve existing data rows
  }
  // ensure header
  const existing = sh.getDataRange().getValues();
  const hdr = ['Email','Role','Leader (if Role = Leader)','Notes'];
  if (!existing.length || String(existing[0][0]).toLowerCase() !== 'email') {
    sh.insertRowBefore(1);
    sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  }
  sh.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#032D60').setFontColor('#FFFFFF');
  // add the script owner if not already there
  const owner = Session.getEffectiveUser().getEmail().toLowerCase();
  const rows = sh.getDataRange().getValues();
  if (!rows.slice(1).some(r => String(r[0]).toLowerCase().trim() === owner)) {
    sh.appendRow([owner, 'Admin', '', 'Script owner (added automatically)']);
  }
  // set dropdowns
  const last = Math.max(sh.getLastRow(), 2);
  sh.getRange(2,2,last-1,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Admin','Leader'],true).setAllowInvalid(false).build());
  if (leaders.length) {
    sh.getRange(2,3,last-1,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(leaders,true).setAllowInvalid(true).build());
  }
  sh.setColumnWidth(1,300); sh.setColumnWidth(2,120); sh.setColumnWidth(3,280); sh.setColumnWidth(4,320); sh.setColumnWidth(5,300);
  sh.setFrozenRows(1);
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    'ACCESS SHEET READY\n\n' +
    'For each person who should see the dashboard, add a row:\n\n' +
    '• Email: their work Google account\n' +
    '• Role = Admin: sees all leaders (for you and finance)\n' +
    '• Role = Leader, then pick their name in column C: sees only their team\n\n' +
    'One person can have multiple rows to see multiple leaders.\n' +
    (full ? leaders.length + ' leaders are available.\n\nRole = Region lets someone see an entire region. Put India or ASEAN in column C, and any budget owners to hide in column D.' : 'Run "Rebuild data now" first to populate the leader dropdown.'));
}

function readAccess_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.access);
  const out = {};
  if (!sh) return out;
  const rows = sh.getDataRange().getValues();
  let owner = ''; try { owner = Session.getEffectiveUser().getEmail().toLowerCase(); } catch(e) {}
  rows.slice(1).forEach(r => {
    const em = String(r[0] || '').toLowerCase().trim();
    const role = String(r[1] || '').trim().toLowerCase();
    const scope = String(r[2] || '').trim();
    const hideBud = String(r[3] || '').trim();
    if (!em) return;
    if (role === 'admin') { out[em] = [{type:'admin'}]; return; }
    if (role === 'leader' && scope) { (out[em] = out[em] || []).push({type:'leader',name:scope}); }
    if (role === 'region' && scope) { (out[em] = out[em] || []).push({type:'region',region:scope,hide:hideBud?hideBud.split(',').map(s=>s.trim()).filter(Boolean):[]}); }
  });
  if (owner && !out[owner]) out[owner] = [{type:'admin'}];
  return out;
}

// ===================== WEB APP =====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('T&E Command Center')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** The data a user may see, given their access. Used by the dashboard and by messages, so both always agree. */
/** The full data with the server-only fields (Workday chains) removed. */
function stripServer_(full){
  const omit=(o,k)=>{const r={};Object.keys(o).forEach(x=>{if(x!==k)r[x]=o[x];});return r;};
  return {meta:full.meta,dict:full.dict,tx:omit(full.tx,'ch'),ta:omit(full.ta,'ch'),budget:full.budget};
}
function viewFor_(full, scopes){
  if (scopes.some(s => s.type === 'admin')) return stripServer_(full);
  const leaderScopes = scopes.filter(s => s.type === 'leader').map(s => s.name);
  const regionScopes = scopes.filter(s => s.type === 'region');
  return regionScopes.length ? regionSubsetPayload_(full, regionScopes, leaderScopes) : subsetPayload_(full, leaderScopes);
}

function getPayload() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!email) return JSON.stringify({ error: 'We could not confirm your identity. Please open this link while signed in to your work Google account.' });
  const access = readAccess_();
  const scopes = access[email];
  if (!scopes || !scopes.length) return JSON.stringify({ error: 'Your account (' + email + ') is not on the Access sheet for this dashboard. Ask the dashboard owner to add you.' });
  const fresh = ensureFresh_();
  const full = loadData_();
  if (!full) return JSON.stringify({ error: 'The dashboard data could not be built' + (fresh.error ? ': ' + String(fresh.error).replace(/\.+$/, '') : '') + '. The owner can check the sheet names and columns, then run T&E Dashboard > Rebuild data now.' });
  const isAdmin = scopes.some(s => s.type === 'admin');
  let p; try { p = viewFor_(full, scopes); } catch (e) { return JSON.stringify({ error: String(e && e.message ? e.message : e) }); }
  p.meta.user = email; p.meta.codeVersion = CODE_VERSION;
  if (fresh.error) p.meta.refreshError = fresh.error; else if (PropertiesService.getScriptProperties().getProperty('LAST_ERROR')) p.meta.refreshError = PropertiesService.getScriptProperties().getProperty('LAST_ERROR');
  p.meta.role = isAdmin ? 'Admin' : scopes.some(s => s.type === 'region') ? 'Region' : 'Leader';
  { const note = PropertiesService.getScriptProperties().getProperty('LEADER_SYNC_NOTE'); if (isAdmin && note) p.meta.syncWarning = note; }
  const gz = Utilities.gzip(Utilities.newBlob(JSON.stringify(p), 'application/json'));
  return JSON.stringify({ gz: Utilities.base64Encode(gz.getBytes()) });
}

// ===================== DATA STORAGE =====================
function saveData_(obj) {
  const props = PropertiesService.getScriptProperties();
  const blob = Utilities.gzip(Utilities.newBlob(JSON.stringify(obj),'application/json','d.json'),'d.json.gz');
  const b64 = Utilities.base64Encode(blob.getBytes());
  const id = props.getProperty('DATA_FILE_ID'); let file = null;
  if (id) { try { file = DriveApp.getFileById(id); } catch(e) {} }
  if (file) file.setContent(b64);
  else { file = DriveApp.createFile('T&E dashboard data (do not share).txt', b64, 'text/plain'); props.setProperty('DATA_FILE_ID', file.getId()); }
}
function loadData_() {
  const id = PropertiesService.getScriptProperties().getProperty('DATA_FILE_ID');
  if (!id) return null;
  try { const b64 = DriveApp.getFileById(id).getBlob().getDataAsString();
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64),'application/x-gzip')).getDataAsString());
  } catch(e) { return null; }
}

// ===================== HELPERS =====================
const DAY=864e5;
function ymd_(s){const p=s.split('-').map(Number);return Date.UTC(p[0],p[1]-1,p[2]);}
function dayNum_(v,q0){if(v===''||v===null||v===undefined)return null;let d;if(Object.prototype.toString.call(v)==='[object Date]'){if(isNaN(v))return null;d=Date.UTC(v.getFullYear(),v.getMonth(),v.getDate());}else{const s=String(v).trim();const m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(!m)return null;d=Date.UTC(+m[1],+m[2]-1,+m[3]);}return Math.round((d-q0)/DAY);}
function num_(v){if(typeof v==='number')return v;if(v===''||v===null||v===undefined)return NaN;return+String(v).replace(/,/g,'').trim()||NaN;}
function str_(v){return v===null||v===undefined?'':String(v).trim();}
function blank_(v){return v===null||v===undefined||String(v).trim()===''||String(v).trim().toLowerCase()==='nan';}
function cleanName_(n){if(blank_(n))return null;return String(n).replace(/\s*\(\d+\)\s*$/,'').replace(/\s*[（(][^）)]*[）)]\s*/g,' ').replace(/\s+/g,' ').trim()||null;}
function ascii_(s){return s.normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/[^\x00-\x7F]/g,'');}
function k1_(s){return ascii_(s).toLowerCase().replace(/[^a-z]/g,'');}
function toks_(s){return ascii_(s).toLowerCase().split(/\s+/).map(t=>t.replace(/[^a-z]/g,'')).filter(Boolean);}
function title_(s){return s.replace(/\s+/g,' ').toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu,(m,a,b)=>a+b.toUpperCase());}
function isUpper_(s){return s===s.toUpperCase()&&s!==s.toLowerCase();}
function median_(a){if(!a.length)return NaN;const s=a.slice().sort((x,y)=>x-y),m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;}
function header_(row){const h={};row.forEach((v,i)=>{const k=str_(v).toLowerCase();if(k&&!(k in h))h[k]=i;});return h;}
function col_(h,name,sheet){const i=h[name.toLowerCase()];if(i===undefined)throw new Error('Column "'+name+'" not found in '+sheet);return i;}
function sheetValues_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh)throw new Error('Sheet "'+name+'" not found.');return sh.getDataRange().getValues();}

// ===================== RULES =====================
const CAT={'Airfare':'Airfare','Travel/Booking Fees':'Booking & fees','Bank & FX Fees':'Booking & fees','Gratuities/Tips':'Booking & fees','Hotel & Lodging':'Lodging','Room Tax':'Lodging','Hotel Deposit':'Lodging','Guest Hotel & Lodging':'Lodging','Laundry':'Lodging','Taxi':'Ground transport','Rental Car':'Ground transport','Rental Car - Fuel':'Ground transport','Tolls':'Ground transport','Parking':'Ground transport','Personal Car Mileage':'Ground transport','Car Service/Limousine':'Ground transport','Subway/Train/Trolley - Local Public Transit':'Ground transport','Subway/Train/Trolley - Cross State/Country':'Ground transport','Individual Meals':'Meals','Business Meals - Staff':'Meals','Business Meals - Clients/Non-Employees':'Meals','Meal -- Onsite Overtime or Rebillable to Client':'Meals','Entertainment - Staff':'Entertainment','Entertainment - Clients/Non-Employees':'Entertainment'};
function catOf_(x){if(CAT[x])return CAT[x];const l=x.toLowerCase();if(/air|flight/.test(l))return'Airfare';if(/hotel|lodg|room/.test(l))return'Lodging';if(/meal|food/.test(l))return'Meals';if(/entertain/.test(l))return'Entertainment';if(/taxi|car|train|transit|toll|park|mileage|fuel/.test(l))return'Ground transport';return'Booking & fees';}
const PURP={'4) Customer Facing Travel':'Customer-facing travel','1) Dreamforce':'Dreamforce','2) Company Kickoff (CKO)':'Company kickoff','3) Internal Business Travel':'Internal business travel','6) General Business Expenses/Other':'General business','A) New Hire Orientation/Bootcamp':'New hire bootcamp','Undefined':'Purpose not set','Works Council':'Other events','U) Event in a Box':'Other events','7) Education Workshop Subcontractor Expenses':'Other events'};
function purpOf_(p){if(!p)return'Purpose not set';if(PURP[p])return PURP[p];const l=p.toLowerCase();if(l.indexOf('customer')>=0)return'Customer-facing travel';if(l.indexOf('dreamforce')>=0)return'Dreamforce';if(l.indexOf('kickoff')>=0)return'Company kickoff';if(l.indexOf('internal')>=0)return'Internal business travel';if(l.indexOf('new hire')>=0||l.indexOf('bootcamp')>=0)return'New hire bootcamp';return'Other events';}
const VR=[[/uber/i,'Uber'],[/grab|gp network asia/i,'Grab'],[/american express global|global business ?travel|gbt india|amex gbt|advanced reservation/i,'Amex GBT (agency)'],[/air india/i,'Air India'],[/singapore air/i,'Singapore Airlines'],[/emirates/i,'Emirates'],[/hilton|doubletree|hampton inn|conrad/i,'Hilton'],[/marriott|sheraton|st\.? regis|le meridien|westin|courtyard|ritz|w hotel|jw /i,'Marriott'],[/hyatt/i,'Hyatt'],[/^itc /i,'ITC Hotels'],[/indigo|interglobe/i,'IndiGo'],[/avis/i,'Avis'],[/makemytrip/i,'MakeMyTrip'],[/lufthansa/i,'Lufthansa'],[/^ua bsp|^united/i,'United Airlines'],[/eva air/i,'EVA Air'],[/thai airways/i,'Thai Airways'],[/scoot/i,'Scoot'],[/bluebird/i,'Blue Bird taxi'],[/japan airlines/i,'Japan Airlines'],[/taj |vivanta/i,'Taj Hotels'],[/akasa/i,'Akasa Air'],[/vistara/i,'Vistara'],[/cathay/i,'Cathay Pacific'],[/malaysia airlines/i,'Malaysia Airlines'],[/philippine airlines/i,'Philippine Airlines'],[/garuda/i,'Garuda Indonesia'],[/airasia|air asia/i,'AirAsia'],[/drivecarclub/i,'DriveCarClub'],[/british airways/i,'British Airways'],[/virgin atlantic/i,'Virgin Atlantic'],[/accor|novotel|pullman|sofitel|ibis|mercure/i,'Accor'],[/ihg|intercontinental|crowne plaza|holiday inn/i,'IHG'],[/shangri/i,'Shangri-La'],[/oberoi|trident/i,'Oberoi'],[/^ola|ani technologies/i,'Ola']];
function vendor_(v){const s=str_(v);if(!s||/^(\*|unspecified|none|null|n\/?a|nan|-)$/i.test(s))return'Not specified';for(const[re,n]of VR)if(re.test(s))return n;return isUpper_(s)?title_(s):s;}
const CITY={'Bangalore':'Bengaluru','Delhi':'New Delhi','Calcutta':'Kolkata','Chennai (Ex Madras)':'Chennai','Singapur':'Singapore','Changi':'Singapore','Ho Chi Minh':'Ho Chi Minh City','Makati City':'Makati','Davao City':'Davao','Trivandrum':'Thiruvananthapuram','Thiruvananthapuram (Ex Trivandrum':'Thiruvananthapuram','Cochin':'Kochi','Frankfurt/Main':'Frankfurt','Muenchen':'Munich','Dehra Dun':'Dehradun','Baroda':'Vadodara','*':''};
const ASEAN_COUNTRIES=['Singapore','Malaysia','Indonesia','Thailand','Philippines','Viet Nam','Vietnam','Lao PDR','Cambodia','Myanmar','Brunei'];
const CC={'IN':'India','PH':'Philippines','TH':'Thailand','MY':'Malaysia','ID':'Indonesia','SG':'Singapore','VN':'Viet Nam','US':'USA','SE':'Sweden','PG':'Indonesia','LK':'Sri Lanka','JP':'Japan','KR':'South Korea','AU':'Australia','AE':'United Arab Emirates','GB':'United Kingdom'};
const INCITY=['hyderabad','bangalore','bengaluru','delhi','mumbai','bombay','chennai','madras','kolkata','calcutta','pune','goa','ahmedabad','cochin','kochi','jaipur','lucknow','chandigarh','nagpur','indore','coimbatore','bhubaneswar','guwahati','trivandrum','thiruvananthapuram','mangalore','vijayawada','vishakhapatnam','visakhapatnam','raipur','bhopal','udaipur','amritsar','varanasi','surat','patna','ranchi','srinagar','dehradun','bagdogra','calicut','madurai','tirupati','baroda','vadodara','jodhpur','rajkot','ludhiana'];
function country_(a){const s=str_(a);if(!s||s==='*')return null;const m=s.match(/-([A-Z]{2})-/);if(m)return CC[m[1]]||m[1];if(s.indexOf(',')>=0){const c=s.split(',').pop().trim();if(c==='USA'||c==='SA'||/^[A-Z]{2}$/.test(c))return'USA';return({Maledives:'Maldives',Vietnam:'Viet Nam'})[c]||c;}const lo=s.toLowerCase();return INCITY.some(k=>lo.indexOf(k)>=0)?'India':null;}
function city_(a){const s=str_(a);if(!s||s==='*')return null;const c=s.split(/ - |,|-[A-Z]{2}-/)[0].trim();return({'Bombay (Mumbai)':'Mumbai','Madras (Chennai)':'Chennai','Calcutta (Kolkata)':'Kolkata','Saigon (Ho Chi Minh City)':'Ho Chi Minh City','Bangalore':'Bengaluru','Seattle/Tacoma':'Seattle'})[c]||c;}
const FLAGS=['air_cap','hotel_cap','premium','late','freeze','split'];
const ASEAN_CUR=['SGD','IDR','THB','PHP','MYR','VND','KHR','MMK','LAK','BND'];

// ===================== NAME MERGING =====================
function buildCanon_(weights,idByName,leaderish){
  const cnt={};Object.keys(weights).forEach(n=>{const c=cleanName_(n);if(c)cnt[c]=(cnt[c]||0)+weights[n];});
  const names=Object.keys(cnt).sort((a,b)=>cnt[b]-cnt[a]||(a<b?-1:1));
  const groups=[],byK={},byFL={},log=[];
  const conflict=(a,b)=>{const ia=idByName[a],ib=idByName[b];if(!ia||!ib)return false;for(const x in ia)if(ib[x])return false;return true;};
  names.forEach(n=>{const tk=toks_(n),k=k1_(n),fl=tk.length>=2?tk[0]+'|'+tk[tk.length-1]:null;
    const cands=[].concat(byK[k]||[],fl?(byFL[fl]||[]):[]).sort((a,b)=>a-b);let placed=-1;
    for(const gi of cands){const g=groups[gi],head=g[0];let why=null;
      if(k1_(head)===k)why='spacing or capitals';else{const th=toks_(head);const sa=new Set(tk),sb=new Set(th);if(tk.length>=2&&th.length>=2&&tk[0]===th[0]&&tk[tk.length-1]===th[th.length-1]&&([...sa].every(x=>sb.has(x))||[...sb].every(x=>sa.has(x))))why='middle name';}
      if(why&&!g.some(m=>conflict(n,m))){g.push(n);log.push([n,head,why]);placed=gi;break;}}
    if(placed<0){groups.push([n]);placed=groups.length-1;}
    (byK[k]=byK[k]||[]).indexOf(placed)<0&&byK[k].push(placed);if(fl)(byFL[fl]=byFL[fl]||[]).indexOf(placed)<0&&byFL[fl].push(placed);});
  const map={};groups.forEach(g=>{const score=n=>[leaderish[n]?1:0,(n!==n.toUpperCase()&&n!==n.toLowerCase())?1:0,cnt[n],n.length];
    const canon=g.slice().sort((a,b)=>{const x=score(a),y=score(b);for(let i=0;i<4;i++)if(x[i]!==y[i])return y[i]-x[i];return 0;})[0];g.forEach(n=>map[n]=canon);});
  Object.keys(CONFIG.aliases).forEach(a=>{const ca=cleanName_(a),cb=cleanName_(CONFIG.aliases[a]),tgt=map[cb]||cb,src=map[ca]||ca;
    Object.keys(map).forEach(n=>{if(map[n]===src)map[n]=tgt;});map[ca]=tgt;log.push([a,tgt,'alias']);});
  const canon=v=>{const c=cleanName_(v);if(!c)return null;const r=map[c]||c;return(isUpper_(r)&&r.length>3)?title_(r):r;};
  return{canon,log:log.filter(l=>cleanName_(l[0])!==(map[cleanName_(l[0])]||l[1]))};
}

// ===================== REBUILD =====================
function rebuildData(){
  const lock=LockService.getScriptLock();lock.waitLock(120000);
  try{const t0=Date.now();const full=rebuildCore_();
    const msg='Rebuilt: '+full.meta.lines.toLocaleString()+' lines, '+full.meta.taRows.toLocaleString()+' TAs, budget $'+Math.round(full.meta.audit.budgetTotal).toLocaleString()+' in '+Math.round((Date.now()-t0)/1000)+'s.';
    try{SpreadsheetApp.getActive().toast(msg,'T&E Dashboard',8);}catch(e){}return msg;}
  finally{lock.releaseLock();}
}
/** Builds the data, stores it, and records when it was built (start time, so edits made DURING a build are never lost). */
function rebuildCore_(){
  const props=PropertiesService.getScriptProperties();const started=Date.now();
  let sync; try{ sync=syncLeaders_(); }catch(e){ sync={ok:false,error:'Leader sync failed: '+(e&&e.message?e.message:e)}; }
  props.setProperty('LEADER_SYNC_NOTE', sync.ok?'':sync.error);
  const full=processAll_(); full.meta.audit.leaderSync=sync; saveData_(full);
  try{writeChecks_(full);}catch(e){console.error('Dashboard Checks not written: '+e);}
  try{writeVarianceTab_(full);}catch(e){console.error('Variance Audit not written: '+e);}
  props.setProperties({BUILT_START:String(started),BUILT_AT:String(Date.now()),LAST_ERROR:''});
  return full;
}
/** Called every time someone opens the dashboard. Rebuilds first if the data is old or a data sheet was edited. */
function ensureFresh_(){
  const props=PropertiesService.getScriptProperties();
  const stale=()=>{
    if(!props.getProperty('DATA_FILE_ID'))return true;
    const built=Number(props.getProperty('BUILT_START')||0);if(!built)return true;
    if(Number(props.getProperty('DIRTY_AT')||0)>built)return true;                       // a data sheet was edited after the last build
    return CONFIG.autoRefreshMinutes>0&&(Date.now()-Number(props.getProperty('BUILT_AT')||0))>CONFIG.autoRefreshMinutes*60000;
  };
  if(!stale())return {refreshed:false};
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(90000))return {refreshed:false,note:'another refresh was already running'};   // serve what exists
  try{if(!stale())return {refreshed:false};rebuildCore_();return {refreshed:true};}
  catch(e){console.error('Auto-refresh failed: '+e);props.setProperty('LAST_ERROR',String(e&&e.message?e.message:e));return {refreshed:false,error:String(e&&e.message?e.message:e)};}
  finally{lock.releaseLock();}
}
/** Installable edit trigger: only notes THAT something changed. The rebuild itself happens when the dashboard is next opened. */
function onSourceEdit(e){
  try{const name=e&&e.range&&e.range.getSheet?e.range.getSheet().getName():'';
    if(CONFIG.editWatchSheets.indexOf(name)>=0)PropertiesService.getScriptProperties().setProperty('DIRTY_AT',String(Date.now()));}catch(err){console.error(err);}
}
function installEditTrigger(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='onSourceEdit').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSourceEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getActive().toast('Done. Edits to the data sheets now refresh the dashboard the next time someone opens it.','T&E Dashboard',8);
}

function processAll_(){
  const ST=readSettings_(), P=readPolicy_();
  const Q0=ymd_(ST.qStart),QD=Math.round((ymd_(ST.qEnd)-Q0)/DAY)+1,QL=QD-1;
  const sv=sheetValues_(CONFIG.sheets.spend);const h=header_(sv[0]);const S=CONFIG.sheets.spend;
  const C={};['Day of Tran Dt','OU','Leader','Leader+1','Leader+2','Employee ID','Employee Name','Report Purpose','Expense Type','Approval Status','Vendor','Flight Class','Arriving Airport','City','Departing Airport','Departure Date','Transaction Currency1','Expense Amount (USD)','Expense Amt Transaction Rate','Total # of Attendees'].forEach(n=>C[n]=col_(h,n,S));
  const raw=sv.slice(1).filter(r=>!blank_(r[C['Employee ID']])&&!isNaN(num_(r[C['Expense Amount (USD)']])));
  const bsh=SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.budget);
  if(!bsh)throw new Error('Sheet "'+CONFIG.sheets.budget+'" not found.');
  const brng=bsh.getDataRange(),bvals=brng.getValues(),bdisp=brng.getDisplayValues();
  const INDIV=individualSet_(bvals);
  const budgetRaw=[];
  bvals.slice(1).forEach((r,i)=>{if(blank_(r[1]))return;const d=bdisp[i+1]||[];
    budgetRaw.push({ou:CONFIG.budgetRegionMap[str_(r[0])]||str_(r[0]),n:str_(r[1]),amt:parseBudget_(r[2],d[2]),region:str_(r[0]),rawVal:r[2],disp:str_(d[2])});});
  const HD=hierarchyData_(); const hv=HD.values;
  const taRaw=readTA_(Q0);
  // names
  const idByName={};const addId=(n,id)=>{const c=cleanName_(n);if(!c||blank_(id))return;(idByName[c]=idByName[c]||{})[String(Math.round(num_(id)))]=1;};
  raw.forEach(r=>addId(r[C['Employee Name']],r[C['Employee ID']]));
  if(hv.length){const hh=header_(hv[0]);const pi=hh['preferred name'],ii=hh['employee id'];if(pi!==undefined&&ii!==undefined)hv.slice(1).forEach(r=>addId(r[pi],r[ii]));}
  taRaw.forEach(r=>addId(r.employee,r.empNo));
  const SPECIAL={'Leader not in chain':1};CONFIG.notALeader.forEach(n=>SPECIAL[n]=1);
  const w={};['Leader','Leader+1','Leader+2','Employee Name'].forEach(c=>raw.forEach(r=>{const v=str_(r[C[c]]);if(v&&!SPECIAL[v])w[v]=(w[v]||0)+1;}));
  budgetRaw.forEach(b=>w[b.n]=(w[b.n]||0)+1);taRaw.forEach(r=>{const v=str_(r.employee);if(v)w[v]=(w[v]||0)+1;});
  const leaderish={};budgetRaw.concat(raw.map(r=>({n:str_(r[C['Leader']])}))).forEach(x=>{const c=cleanName_(x.n);if(c)leaderish[c]=1;});
  const CN=buildCanon_(w,idByName,leaderish);const cname=CN.canon;
  const empName={};raw.forEach(r=>{const id=String(Math.round(num_(r[C['Employee ID']])));if(!(id in empName))empName[id]=cname(r[C['Employee Name']])||('Employee '+id);});
  const nameCount={};Object.keys(empName).forEach(id=>nameCount[empName[id]]=(nameCount[empName[id]]||0)+1);
  const namesakes=Object.keys(nameCount).filter(n=>nameCount[n]>1).sort();
  Object.keys(empName).forEach(id=>{if(nameCount[empName[id]]>1)empName[id]=empName[id]+' (ID '+id+')';});
  const lab=(v,ph)=>{if(blank_(v))return ph;const s=str_(v);if(s==='Leader not in chain')return'Outside reporting line';return cname(s);};
  const budgets=budgetRaw.map(b=>({ou:b.ou,n:cname(b.n),amt:b.amt}));
  const budKey={};budgets.forEach(b=>budKey[b.ou+'|'+b.n]=1);
  const T=raw.map(r=>{const ldRaw=str_(r[C['Leader']]);
    const o={id:String(Math.round(num_(r[C['Employee ID']]))),ou:str_(r[C['OU']]),ld:SPECIAL[ldRaw]?ldRaw:(cname(ldRaw)||'Not in employee master'),l1:lab(r[C['Leader+1']],PH_L1),l2:lab(r[C['Leader+2']],PH_L2),d:dayNum_(r[C['Day of Tran Dt']],Q0),a:num_(r[C['Expense Amount (USD)']]),x:str_(r[C['Expense Type']])||'Other',st:str_(r[C['Approval Status']])||'Unknown',vRaw:r[C['Vendor']],fc:str_(r[C['Flight Class']]),dep:r[C['Departing Airport']],arr:r[C['Arriving Airport']],cur:str_(r[C['Transaction Currency1']]),loc:num_(r[C['Expense Amt Transaction Rate']]),att:Math.round(num_(r[C['Total # of Attendees']])||0),purpRaw:str_(r[C['Report Purpose']]),depDate:r[C['Departure Date']]};
    o.e=empName[o.id];o.ow=(o.ld&&budKey[o.ou+'|'+o.ld])?o.ld:'No budget line';
    o.cat=catOf_(o.x);o.p=purpOf_(o.purpRaw);o.cu=o.purpRaw==='4) Customer Facing Travel'||o.x.indexOf('Clients')>=0;o.v=vendor_(o.vRaw);const cRaw=str_(r[C['City']]);o.c=cRaw in CITY?CITY[cRaw]:cRaw;
    if(o.x==='Airfare'){const ca=country_(o.dep),cb=country_(o.arr);o.sc=(!ca||!cb)?'Route not captured':ca===cb?'Domestic':(o.ou==='ASEAN'&&ASEAN_COUNTRIES.indexOf(ca)>=0&&ASEAN_COUNTRIES.indexOf(cb)>=0)?'Within ASEAN':'International';const ci=city_(o.dep),cj=city_(o.arr);o.rt=ci&&cj?ci+' – '+cj:'';const dd=dayNum_(o.depDate,Q0);o.adv=dd===null||o.d===null?null:dd-o.d;}else{o.sc='';o.rt='';o.adv=null;}return o;}).filter(o=>o.d!==null);
  const daily={};T.forEach(o=>daily[o.d]=(daily[o.d]||0)+o.a);
  const dd=Object.keys(daily).map(Number).sort((a,b)=>a-b);const med=median_(dd.map(d=>daily[d]));
  const LC=Math.max.apply(null,dd.filter(d=>daily[d]>=0.25*med));const LAST=Math.max.apply(null,dd);
  const inrR=T.filter(o=>o.cur==='INR'&&o.a&&!isNaN(o.loc)).map(o=>o.loc/o.a).filter(isFinite);const INR=median_(inrR)||84;
  // policy flags
  T.forEach(o=>{let f=0;const a=o.a;if(a>0){if(o.x==='Airfare'){if(o.ou==='India'&&o.sc==='Domestic'){const inr=(o.cur==='INR'&&!isNaN(o.loc))?o.loc:a*INR;if(inr>P.inrAir)f|=1;}else if(o.ou==='ASEAN'&&(o.sc==='Domestic'||o.sc==='Within ASEAN')){if(a>P.aseanAir)f|=1;}else if(a>P.intlAir)f|=1;if(o.fc==='Business'||o.fc==='First')f|=4;if(o.adv!==null&&o.adv>=0&&o.adv<=3)f|=8;}if(o.x==='Hotel & Lodging'){if(o.ou==='India'&&o.cur==='INR'){if((!isNaN(o.loc)?o.loc:a*INR)>P.inrHotel)f|=2;}else if(o.ou==='ASEAN'&&ASEAN_CUR.indexOf(o.cur)>=0){if(a>P.aseanHotel)f|=2;}else if(a>P.intlHotel)f|=2;}if(o.p==='Internal business travel'&&(o.cat==='Airfare'||o.cat==='Lodging'))f|=16;}o.f=f;});
  const sp={};T.forEach((o,i)=>{if((o.cat==='Meals'||o.cat==='Lodging'||o.cat==='Entertainment')&&o.a>0&&o.v!=='Not specified'){const k=o.id+'|'+o.d+'|'+o.v+'|'+o.x;(sp[k]=sp[k]||[]).push(i);}});
  Object.keys(sp).forEach(k=>{if(sp[k].length>=2)sp[k].forEach(i=>T[i].f|=32);});
  // TA model
  const ouByEmp={};{const c={};T.forEach(o=>{c[o.id]=c[o.id]||{};c[o.id][o.ou]=(c[o.id][o.ou]||0)+1;});Object.keys(c).forEach(k=>ouByEmp[k]=Object.keys(c[k]).sort((a,b)=>c[k][b]-c[k][a])[0]);}
  const TA=taRaw.map(r=>{const o={id:String(r.empNo),employee:r.employee,mgmt:r.mgmt,reason:r.reason||'Not stated',intl:/international/i.test(r.domIntl),route:str_(r.route).slice(0,48),cday:r.cday,sday:r.sday,eday:r.eday,air:r.air,hotel:r.hotel,rail:r.rail,other:r.other};
    o.days=o.eday-o.sday+1;o.total=o.air+o.hotel+o.rail+o.other;o.ou=ouByEmp[o.id]||null;
    o.fix=!o.intl&&o.ou==='India'&&o.total/o.days>CONFIG.inrTaPerDayThreshold;o.origTotal=o.total;
    if(o.fix){['air','hotel','rail','other','total'].forEach(k=>o[k]=o[k]/INR);}o.nonair=o.hotel+o.rail+o.other;
    let inq=0,aft=0;for(let d=o.sday;d<=o.eday;d++){if(d>=0&&d<=QL)inq++;if(d>LC&&d<=QL)aft++;}o.q3share=inq/o.days;o.opensh=aft/o.days;return o;}).filter(o=>o.days>0);
  const airBy={};T.forEach((o,i)=>{if(o.x==='Airfare'&&o.a>0)(airBy[o.id]=airBy[o.id]||[]).push(i);});
  const used={};TA.map((o,j)=>j).sort((a,b)=>TA[a].cday-TA[b].cday||a-b).forEach(j=>{const o=TA[j];o.mt=false;if(!(o.air>0)||!airBy[o.id])return;for(const i of airBy[o.id]){const d=T[i].d;if(!used[i]&&d>=o.cday-3&&d<=o.eday){used[i]=1;o.mt=true;break;}}});
  TA.forEach(o=>{const prior=o.cday<0;o.prior_air=prior?o.air:0;o.q3_air=prior?0:o.air;o.open_air=(!prior&&!o.mt&&o.sday>LC)?o.air:0;o.q3_nonair=o.nonair*o.q3share;o.open_nonair=o.nonair*o.opensh;});
  const byEmp={};T.forEach((o,i)=>(byEmp[o.id]=byEmp[o.id]||[]).push(i));
  TA.forEach(r=>{(byEmp[r.id]||[]).forEach(i=>{const o=T[i];if(o.x==='Airfare'?(o.d>=r.cday-3&&o.d<=r.eday):(o.d>=r.sday-1&&o.d<=r.eday+1))o.lk=1;});});
  const obs=TA.filter(o=>o.sday>=14&&o.sday<=LC);const wsum=obs.reduce((s,o)=>s+o.nonair,0);
  const COV=[];for(let k=0;k<40;k++){const v=wsum>0?obs.reduce((s,o)=>s+(o.sday-o.cday>=k?o.nonair:0),0)/wsum:1;COV.push(Math.round(Math.max(0.2,v)*1e4)/1e4);}
  const cov=k=>Math.max(0.2,wsum>0?obs.reduce((s,o)=>s+(o.sday-o.cday>=k?o.nonair:0),0)/wsum:1);
  TA.forEach(o=>{let s=0;for(let d=Math.max(o.sday,LC+1);d<=Math.min(o.eday,QL);d++)s+=o.nonair/o.days*(1/cov(d-LC)-1);o.exp_unraised=s;});
  // Workday chains: used only on the server, to decide who may see which rows. They are never sent to a browser.
  const LCTX=(()=>{try{const x=buildLeaderCtx_();return x.ok?x:null;}catch(e){return null;}})();
  const CHAINS=[], chKey={};
  const chOf=id=>{ if(!LCTX||LCTX.chain[id]===undefined) return -1;
    const names=LCTX.chain[id].slice().reverse().map(a=>cname(LCTX.idNm[a]||a)).filter(Boolean); const k=names.join('\u0001');
    if(!(k in chKey)){chKey[k]=CHAINS.length;CHAINS.push(names);} return chKey[k]; };
  T.forEach(o=>{o.ch=chOf(o.id);});
  const ANC={};
  if(LCTX) Object.keys(LCTX.chain).forEach(id=>{ const nm=cname(LCTX.idNm[id]||''); if(!nm) return; const set=(ANC[nm]=ANC[nm]||{});
    LCTX.chain[id].forEach(a=>{const an=cname(LCTX.idNm[a]||a); if(an&&an!==nm) set[an]=1;}); });
  Object.keys(ANC).forEach(k=>ANC[k]=Object.keys(ANC[k]));
  const pos={};{const c={};T.forEach(o=>{const k=[o.ou,o.ow,o.ld,o.l1,o.l2].join('\u0001');c[o.id]=c[o.id]||{};c[o.id][k]=(c[o.id][k]||0)+1;});Object.keys(c).forEach(id=>pos[id]=Object.keys(c[id]).sort((a,b)=>c[id][b]-c[id][a])[0].split('\u0001'));}
  const leaders={};T.forEach(o=>{if(!SPECIAL[o.ld])leaders[o.ld]=1;});
  const leaderOu={};{const c={};T.forEach(o=>{c[o.ld]=c[o.ld]||{};c[o.ld][o.ou]=(c[o.ld][o.ou]||0)+1;});Object.keys(c).forEach(l=>leaderOu[l]=Object.keys(c[l]).sort((a,b)=>c[l][b]-c[l][a])[0]);}
  TA.forEach(o=>{o.e=empName[o.id]||cname(o.employee)||('Employee '+o.id);let p=pos[o.id];
    if(!p){ const ou=o.ou||(LCTX&&LCTX.ouOfId[o.id])||null;
      if(LCTX&&ou&&LCTX.chain[o.id]!==undefined){ const a=assignLeaders_(LCTX,o.id,ou), L=a.L.map((x,i)=>x?cname(x):''), own=L[0]&&L[0]!=='No budget line'&&budKey[ou+'|'+L[0]]?L[0]:'No budget line';
        p=[ou,own,L[0]||'No budget line',L[1]||PH_L1,L[2]||PH_L2]; }
      else p=[ou,'No budget line','No budget line',PH_L1,PH_L2]; }
    if(!o.ou)o.ou=p[0];o.ow=p[1];o.ld=p[2];o.l1=p[3];o.l2=p[4];if(!o.ou)o.ou='India';o.ch=chOf(o.id);});
  const recipients=Object.keys(leaders).concat(budgets.map(b=>b.n)).filter((v,i,a)=>a.indexOf(v)===i&&T.some(o=>o.ld===v||o.l1===v||o.l2===v)).sort();
  const fixes=TA.filter(o=>o.fix).map(o=>[o.e,o.route,o.origTotal,Math.round(o.total)]);
  const audit={merges:CN.log,namesakes,inrFixes:fixes,inrRate:INR,columnShift:taRaw.shifted,lines:T.length,total:T.reduce((s,o)=>s+o.a,0),taRows:TA.length,taTotal:TA.reduce((s,o)=>s+o.total,0),budgetTotal:budgets.reduce((s,b)=>s+b.amt,0),budgetRows:budgetRaw.map((b,i)=>[b.region||b.ou,b.n,str_(b.rawVal),b.disp,b.amt,budgets[i].n]),skippedRows:sv.length-1-T.length,unknownTypes:[...new Set(T.filter(o=>!CAT[o.x]).map(o=>o.x))],unknownPurposes:[...new Set(T.filter(o=>o.purpRaw&&!PURP[o.purpRaw]).map(o=>o.purpRaw))]};
  return encode_(T,TA,budgets,{Q0,QD,LC,LAST,COV,recipients,audit,scope:'All leaders',scopeKind:'all',settings:ST,policy:P,chains:CHAINS,anc:ANC,wdPeople:LCTX?LCTX.people:0});
}
function parseBudget_(v,disp){
  const d=str_(disp).toUpperCase().replace(/[$,\s]/g,'');const md=d.match(/^([\d.]+)([KM])$/);
  if(md)return parseFloat(md[1])*(md[2]==='K'?1e3:1e6);            // the cell SHOWS 93K or 1.2M: trust what is shown
  if(typeof v==='number')return v;
  const s=str_(v).toUpperCase().replace(/[$,\s]/g,'');const m=s.match(/^([\d.]+)([KM]?)$/);
  if(!m)return 0;return parseFloat(m[1])*(m[2]==='K'?1e3:m[2]==='M'?1e6:1);
}
function readTA_(Q0){let v;try{v=sheetValues_(CONFIG.sheets.ta);}catch(e){return Object.assign([],{shifted:false});}
  const hdr=v[0].map(str_);const hIdx=hdr.indexOf('Travel Approval');let shifted=false;const out=[];
  v.slice(1).forEach(r=>{let k=-1;for(let i=0;i<r.length;i++)if(/^TA-\d+$/.test(str_(r[i]))){k=i;break;}if(k<0)return;if(hIdx>=0&&k!==hIdx)shifted=true;
    const o={empNo:Math.round(num_(r[0])),employee:str_(r[1]),mgmt:r.slice(2,k).map(str_).filter(Boolean),taId:str_(r[k]),cday:dayNum_(r[k+1],Q0),status:str_(r[k+2]),reason:str_(r[k+3]),route:str_(r[k+4]),domIntl:str_(r[k+5]),sday:dayNum_(r[k+7],Q0),eday:dayNum_(r[k+8],Q0),air:num_(r[k+9])||0,hotel:num_(r[k+10])||0,rail:num_(r[k+11])||0,other:num_(r[k+12])||0};
    if(o.status&&!/approved/i.test(o.status))return;if(o.sday===null||o.eday===null||o.cday===null||isNaN(o.empNo))return;out.push(o);});
  out.shifted=shifted;return out;}
function encode_(T,TA,budgets,X){
  const D={},I={};const dict=(k,vals)=>{const s=[...new Set(vals.map(v=>v===null||v===undefined?'':String(v)))].sort();D[k]=s;I[k]={};s.forEach((x,i)=>I[k][x]=i);};
  ['ou','ow','ld','l1','l2','e'].forEach(k=>dict(k,T.map(o=>o[k]).concat(TA.map(o=>o[k]))));
  const own={x:'x',p:'p',s:'st',v:'v',c:'c',sc:'sc',rt:'rt',fc:'fc'};Object.keys(own).forEach(k=>dict(k,T.map(o=>o[own[k]])));dict('tr',TA.map(o=>o.reason));
  const tx={};['ou','ow','ld','l1','l2','e'].forEach(k=>tx[k]=T.map(o=>I[k][String(o[k])]));Object.keys(own).forEach(k=>tx[k]=T.map(o=>I[k][String(o[own[k]]===null||o[own[k]]===undefined?'':o[own[k]])]));
  tx.d=T.map(o=>o.d);tx.a=T.map(o=>Math.round(o.a*100)/100);tx.att=T.map(o=>o.att||0);tx.adv=T.map(o=>o.adv);tx.f=T.map(o=>o.f);tx.cu=T.map(o=>o.cu?1:0);tx.lk=T.map(o=>o.lk?1:0);
  const ta={};['e','ou','ow','ld','l1','l2'].forEach(k=>ta[k]=TA.map(o=>I[k][String(o[k])]));ta.tr=TA.map(o=>I.tr[String(o.reason)]);ta.intl=TA.map(o=>o.intl?1:0);
  ['cday','sday','eday','days'].forEach(k=>ta[k]=TA.map(o=>o[k]));const r2=v=>Math.round(v*100)/100;
  ta.air=TA.map(o=>r2(o.air));ta.nonair=TA.map(o=>r2(o.nonair));ta.total=TA.map(o=>r2(o.total));ta.prior_air=TA.map(o=>r2(o.prior_air));ta.q3_air=TA.map(o=>r2(o.q3_air));ta.q3_nonair=TA.map(o=>r2(o.q3_nonair));ta.open_air=TA.map(o=>r2(o.open_air));ta.open_nonair=TA.map(o=>r2(o.open_nonair));ta.exp_unraised=TA.map(o=>r2(o.exp_unraised));ta.q3share=TA.map(o=>Math.round(o.q3share*1e4)/1e4);ta.mt=TA.map(o=>o.mt?1:0);ta.fix=TA.map(o=>o.fix?1:0);ta.route=TA.map(o=>o.route);
  const x2cat={};T.forEach(o=>x2cat[o.x]=o.cat);const fixes=TA.filter(o=>o.fix);
  const S=X.settings, P=X.policy;
  const meta={qStart:S.qStart,qDays:X.QD,lastDay:X.LAST,lastComplete:X.LC,flags:FLAGS,limits:P,policyLines:P.lines,cats:['Airfare','Lodging','Ground transport','Meals','Entertainment','Booking & fees'],events:S.events,cov:X.COV,scope:X.scope,scopeKind:X.scopeKind,x2cat,inrFixes:fixes.length,inrOrig:fixes.reduce((s,o)=>s+o.origTotal,0),inrUsd:fixes.reduce((s,o)=>s+o.total,0),lines:T.length,taRows:TA.length,generated:new Date().toISOString(),quarterLabel:S.quarterLabel,region:S.region,recipients:X.recipients,audit:X.audit};
  tx.ch=T.map(o=>o.ch);ta.ch=TA.map(o=>o.ch);meta.workdayPeople=X.wdPeople||0;
  return{meta,dict:D,tx,ta,budget:budgets.map(b=>({ou:b.ou,n:b.n,b:r2(b.amt),full:b.amt,share:1})),chains:X.chains||[],anc:X.anc||{}};
}

function regionSubsetPayload_(full, regionScopes, leaderScopes) {
  const D=full.dict, X=full.tx, A=full.ta;
  const wantOU={};
  regionScopes.forEach(s=>{const oi=D.ou.indexOf(s.region);if(oi>=0)wantOU[oi]=s;});
  const allHide={};
  regionScopes.forEach(s=>(s.hide||[]).forEach(n=>allHide[n]=1));
  const wantLeader={};
  leaderScopes.forEach(n=>wantLeader[n]=1);
  const keepT=[],keepA=[];
  for(let i=0;i<X.a.length;i++){if(wantOU[X.ou[i]]||wantLeader[D.ld[X.ld[i]]]||wantLeader[D.ow[X.ow[i]]])keepT.push(i);}
  for(let j=0;j<A.e.length;j++){if(wantOU[A.ou[j]]||wantLeader[D.ld[A.ld[j]]])keepA.push(j);}
  const nd={},remap={};
  ['ou','ow','ld','l1','l2','e'].forEach(k=>{const used=[...new Set(keepT.map(i=>X[k][i]).concat(keepA.map(j=>A[k][j])))].map(i=>D[k][i]).sort();nd[k]=used;remap[k]={};used.forEach((v,n)=>remap[k][v]=n);});
  ['x','p','s','v','c','sc','rt','fc'].forEach(k=>{const used=[...new Set(keepT.map(i=>X[k][i]))].map(i=>D[k][i]).sort();nd[k]=used;remap[k]={};used.forEach((v,n)=>remap[k][v]=n);});
  {const used=[...new Set(keepA.map(j=>A.tr[j]))].map(i=>D.tr[i]).sort();nd.tr=used;remap.tr={};used.forEach((v,n)=>remap.tr[v]=n);}
  const tx={};['ou','ow','ld','l1','l2','e','x','p','s','v','c','sc','rt','fc'].forEach(k=>tx[k]=keepT.map(i=>remap[k][D[k][X[k][i]]]));
  ['d','a','att','adv','f','cu','lk'].forEach(k=>tx[k]=keepT.map(i=>X[k][i]));
  const ta={};['ou','ow','ld','l1','l2','e'].forEach(k=>ta[k]=keepA.map(j=>remap[k][D[k][A[k][j]]]));
  ta.tr=keepA.map(j=>remap.tr[D.tr[A.tr[j]]]);
  Object.keys(A).forEach(k=>{if(!(k in ta)&&k!=='ch')ta[k]=keepA.map(j=>A[k][j]);});
  const budget=[];
  full.budget.forEach(b=>{const oui=D.ou.indexOf(b.ou);if(!wantOU[oui]&&!wantLeader[b.n])return;if(allHide[b.n])return;budget.push({ou:b.ou,n:b.n,b:b.b,full:b.full,share:b.share});});
  const meta=JSON.parse(JSON.stringify(full.meta));
  meta.scope=regionScopes.map(s=>s.region).join(' + ')+(leaderScopes.length?' + '+leaderScopes.join(', '):'');
  meta.scopeKind='region';meta.lines=keepT.length;meta.taRows=keepA.length;
  meta.hiddenBudgets=Object.keys(allHide);
  delete meta.audit;delete meta.recipients;
  return{meta,dict:nd,tx,ta,budget};
}

/** Cuts the full data to one leader's view.
 *  WHO SEES WHAT follows the Workday reporting chain, not the Leader columns: a leader sees every line for themselves and for everyone below them.
 *  Names ABOVE the leader in Workday (their managers) are replaced by "Above your level", and so are their budgets. */
function subsetPayload_(full,scopes){
  if(!(full.meta.workdayPeople>0)||!full.chains||!full.anc||!full.tx.ch) throw new Error('Workday reporting lines are not loaded yet, so team views are switched off. Ask the dashboard owner to load the Workday data and rebuild.');
  const D=full.dict,X=full.tx,A=full.ta,MASK='Above your level', want=new Set(scopes);
  const unknown=scopes.filter(s=>!full.anc[s]);
  if(unknown.length) throw new Error('Your name ('+unknown.join(', ')+') could not be matched to Workday, so your team cannot be shown. Ask the dashboard owner to add you on the Leader IDs tab.');
  const shared=(full.meta.audit&&full.meta.audit.namesakes)||[], amb=scopes.filter(s=>shared.indexOf(s)>=0);
  if(amb.length) throw new Error('The name '+amb.join(', ')+' is shared by two people in Workday, so a team view cannot be built safely. Ask the dashboard owner to add your Employee ID on the Leader IDs tab.');
  const above=new Set(); scopes.forEach(s=>full.anc[s].forEach(n=>{if(!want.has(n))above.add(n);}));
  const memo=new Map();
  const inOrg=(ci,emp)=>{ if(want.has(emp)) return true; if(ci<0) return false; let v=memo.get(ci);
    if(v===undefined){v=false;const ch=full.chains[ci]||[];for(const n of ch){if(want.has(n)){v=true;break;}}memo.set(ci,v);} return v; };
  const keepT=[],keepA=[];
  for(let i=0;i<X.a.length;i++) if(inOrg(X.ch[i],D.e[X.e[i]])) keepT.push(i);
  for(let j=0;j<A.e.length;j++) if(inOrg(A.ch[j],D.e[A.e[j]])) keepA.push(j);
  const mk=n=>above.has(n)?MASK:n, mo=n=>(n==='No budget line'||!above.has(n))?n:'No budget line';
  const ldR=keepT.map(i=>mk(D.ld[X.ld[i]])), l1R=keepT.map(i=>mk(D.l1[X.l1[i]])), l2R=keepT.map(i=>mk(D.l2[X.l2[i]])), owR=keepT.map(i=>mo(D.ow[X.ow[i]]));
  const tLd=keepA.map(j=>mk(D.ld[A.ld[j]])), tL1=keepA.map(j=>mk(D.l1[A.l1[j]])), tL2=keepA.map(j=>mk(D.l2[A.l2[j]])), tOw=keepA.map(j=>mo(D.ow[A.ow[j]]));
  const eArr=keepT.map(i=>D.e[X.e[i]]), ouArr=keepT.map(i=>D.ou[X.ou[i]]);
  const dictOf=arr=>{const u=[...new Set(arr)].sort();const ix={};u.forEach((v,i)=>ix[v]=i);return{u,ix};};
  const dLd=dictOf(ldR.concat(tLd)), dL1=dictOf(l1R.concat(tL1)), dL2=dictOf(l2R.concat(tL2)), dOw=dictOf(owR.concat(tOw));
  const dE=dictOf(eArr.concat(keepA.map(j=>D.e[A.e[j]]))), dOu=dictOf(ouArr.concat(keepA.map(j=>D.ou[A.ou[j]])));
  const nd={ld:dLd.u,l1:dL1.u,l2:dL2.u,ow:dOw.u,e:dE.u,ou:dOu.u}, tx={};
  tx.ld=ldR.map(v=>dLd.ix[v]); tx.l1=l1R.map(v=>dL1.ix[v]); tx.l2=l2R.map(v=>dL2.ix[v]); tx.ow=owR.map(v=>dOw.ix[v]);
  tx.e=eArr.map(v=>dE.ix[v]); tx.ou=ouArr.map(v=>dOu.ix[v]);
  ['x','p','s','v','c','sc','rt','fc'].forEach(k=>{const arr=keepT.map(i=>D[k][X[k][i]]);const d=dictOf(arr);nd[k]=d.u;tx[k]=arr.map(v=>d.ix[v]);});
  ['d','a','att','adv','f','cu','lk'].forEach(k=>tx[k]=keepT.map(i=>X[k][i]));
  const ta={};
  ta.ld=tLd.map(v=>dLd.ix[v]); ta.l1=tL1.map(v=>dL1.ix[v]); ta.l2=tL2.map(v=>dL2.ix[v]); ta.ow=tOw.map(v=>dOw.ix[v]);
  ta.e=keepA.map(j=>dE.ix[D.e[A.e[j]]]); ta.ou=keepA.map(j=>dOu.ix[D.ou[A.ou[j]]]);
  {const arr=keepA.map(j=>D.tr[A.tr[j]]);const d=dictOf(arr);nd.tr=d.u;ta.tr=arr.map(v=>d.ix[v]);}
  Object.keys(A).forEach(k=>{if(!(k in ta)&&k!=='ch')ta[k]=keepA.map(j=>A[k][j]);});
  const fullOwn={},subOwn={};
  for(let i=0;i<X.a.length;i++){const k=D.ou[X.ou[i]]+'|'+D.ow[X.ow[i]];fullOwn[k]=(fullOwn[k]||0)+X.a[i];}
  keepT.forEach((i,n)=>{const k=ouArr[n]+'|'+owR[n];subOwn[k]=(subOwn[k]||0)+X.a[i];});
  const budget=[];
  full.budget.forEach(b=>{if(above.has(b.n)&&!want.has(b.n))return;const k=b.ou+'|'+b.n;if(!(k in subOwn))return;const fo=fullOwn[k]||0,so=subOwn[k]||0;const sh=fo<=0?1:Math.max(0,Math.min(1,so/fo));budget.push({ou:b.ou,n:b.n,b:Math.round(b.full*sh*100)/100,full:b.full,share:Math.round(sh*1e4)/1e4});});
  const meta=JSON.parse(JSON.stringify(full.meta));
  meta.scope=scopes.join(', ');meta.scopeKind='leader';meta.lines=keepT.length;meta.taRows=keepA.length;
  const fx=keepA.filter(j=>A.fix[j]);meta.inrFixes=fx.length;
  meta.inrOrig=full.meta.audit.inrFixes.filter(f=>fx.some(j=>D.e[A.e[j]]===f[0])).reduce((s,f)=>s+f[2],0);
  meta.inrUsd=fx.reduce((s,j)=>s+A.total[j],0);
  delete meta.audit;delete meta.recipients;
  return{meta,dict:nd,tx,ta,budget};
}

const HIER_HEADERS = ["Employee ID", "Worker", "Preferred Name", "Preferred Name - First Name", "Preferred Name - Last Name", "Currently Active", "Primary Work Email", "Additional Work Email", "Supervisory Organization", "Cost Center", "Company", "Region Hierarchy", "Company - ID", "Product", "Market Segment", "Industry", "Division", "Primary Coverage Country", "Location", "Location Address - Province (Canada)", "Location Address - State (United States)", "Location Address - Country", "Work Space from Worker Primary Job", "Work Address Line 1", "Work Address Line 2", "Work Address - Subdivision", "Work Address - Postal Code", "Phone - Primary Work", "Is Manager?", "Number of Direct Reports", "Job Profile", "Job Title", "Business Title", "Job Family", "Job Family Group", "Position", "Position ID", "Worker Type", "Employee Type", "Time Type", "Work Location Status", "Manager - Level 01", "Manager Title", "Manager's Location", "Manager's Country", "Manager's Division", "Manager's Work Email", "Manager Tableau Email", "1st Line Manager", "2nd Line Manager", "3rd Line Manager", "4th Line Manager", "Supervisory Organization.1", "Manager ID", "Management Level", "Management_Chain_Level_01", "Management_Chain_Level_02", "Management_Chain_Level_03", "Management_Chain_Level_04", "Management_Chain_Level_05", "Management_Chain_Level_06", "Management_Chain_Level_07", "Management_Chain_Level_08", "Management_Chain_Level_09", "Management_Chain_Level_10", "Management_Chain_Level_11", "Management_Chain_Level_12"];
// ===================== CREATE THE TABS THIS DASHBOARD NEEDS (run once) =====================
// Creates only the tabs that are missing. Existing tabs are never changed.
function createMissingTabs(){
  const ss=SpreadsheetApp.getActive(); const created=[];
  const specs=[
    {name:'Dashboard Settings', rows:[['Setting','Value'],['Quarter label','Q3 FY27'],['Quarter start','2026-08-01'],['Quarter end','2026-10-31'],['Region','South Asia'],['Event purposes','Dreamforce, Company kickoff']],
     note:'Check these values before the first rebuild. Dates are yyyy-mm-dd. Event purposes are one-off report purposes left out of the run-rate.'},
    {name:CONFIG.sheets.hierarchy, rows:[HIER_HEADERS],
     note:'Paste the Workday export below these headers (same column layout). The leader sync needs at least 100 employees here.'},
    {name:CONFIG.sheets.leaderIds, rows:[['Name','Employee ID']],
     note:'Only for budget owners the sync cannot match by name. One row per person.'},
    {name:CONFIG.sheets.emails, rows:[['Name','Email']],
     note:'Work emails for people with no email in Workday. Used for Slack messages.'},
    {name:CONFIG.sheets.messageLog, rows:[['Time','Sent by','Recipient','Message type','Result','Detail']],
     note:'Written by the messaging feature.'},
    {name:'Variance Audit', rows:[], note:'Written by every rebuild.'},
    {name:'Leader Sync', rows:[], note:'Written by every leader sync.'}
  ];
  specs.forEach(s=>{
    if(ss.getSheetByName(s.name)) return;
    const sh=ss.insertSheet(s.name); created.push(s.name);
    if(s.rows.length){ const w=Math.max(...s.rows.map(r=>r.length)); sh.getRange(1,1,s.rows.length,w).setValues(s.rows.map(r=>{const x=r.slice();while(x.length<w)x.push('');return x;})); sh.getRange(1,1,1,w).setFontWeight('bold'); sh.setFrozenRows(1); }
    sh.getRange(1,1).setNote(s.note);
  });
  const msg=created.length?('Created: '+created.join(', ')+'. Check Dashboard Settings, paste the Workday export into '+CONFIG.sheets.hierarchy+', then run Sync leaders from Workday and Rebuild data now.'):'All dashboard tabs already exist.';
  try{ss.toast(msg,'T&E Dashboard',12);}catch(e){}
  return msg;
}

// ===================== WORKDAY: finding the data, assigning leaders, sync, import =====================
// Rule (agreed with the dashboard owner):
//   Leader   = the person's own budget if they hold one in the region where the line is booked; otherwise the NEAREST leader above
//              them whose budget covers their team in that region. Budgets marked "Individual only" on the Budget tab are skipped.
//              If there is no such leader, the line is "No budget line".
//   Leader+1 = the next level BELOW the Leader on the way to the employee (the manager just under the Leader).
//   Leader+2 = the level below that. Both are blank for the Leader's own spend and for people who report straight to the Leader.
// Who can see which rows does NOT depend on these three columns. It follows the Workday reporting chain (see subsetPayload_).
const WORKDAY_REGIONS = ['INDIA','ASEAN'];
const PH_L1 = 'Own spend or direct reports', PH_L2 = 'No second level';

function ensureRows_(sh, needRows){ const have=sh.getMaxRows(); if(needRows>have) sh.insertRowsAfter(have, needRows-have); }

function individualSet_(bvals){
  const h=header_(bvals[0]||[]), ci=h['budget covers'], set=new Set();
  if(ci!==undefined){ bvals.slice(1).forEach(r=>{ if(/individual\s*only/i.test(str_(r[ci]))) set.add(cleanName_(r[1])); }); }
  if(ci===undefined||!set.size) CONFIG.individualOnly.forEach(n=>set.add(cleanName_(n)));
  return set;
}

/** Finds the Workday data. Uses the "Employee Hierachy" tab, or "Workday Import" if that is where the data is. The header row may sit below title lines. */
function hierarchyData_(){
  const ss=SpreadsheetApp.getActive(); let best=null;
  for(const n of [CONFIG.sheets.hierarchy,'Employee Hierarchy','Workday Import']){
    const sh=ss.getSheetByName(n); if(!sh) continue;
    const v=sh.getDataRange().getValues(), hr=v.findIndex(r=>str_(r[0])==='Employee ID'); if(hr<0) continue;
    const rows=v.slice(hr), people=rows.slice(1).filter(r=>str_(r[0])).length, cand={name:n,values:rows,people};
    if(people>=100) return cand;
    if(!best||people>best.people) best=cand;
  }
  return best||{name:null,values:[],people:0};
}

/** Everything needed to place a person in the budget structure: Workday chains, budget owners, who is individual-only. */
function buildLeaderCtx_(){
  const ss=SpreadsheetApp.getActive(), bud=ss.getSheetByName(CONFIG.sheets.budget);
  if(!bud) return {ok:false,error:'The "'+CONFIG.sheets.budget+'" tab is missing.'};
  const hd=hierarchyData_();
  if(hd.people<100) return {ok:false,error:'The Workday data has '+hd.people+' employees'+(hd.name?' (tab "'+hd.name+'")':'')+'. Paste the Workday export into a tab named "'+CONFIG.sheets.hierarchy+'" (at least 100 employees are needed).'};
  const hv=hd.values, hh=header_(hv[0]), eidCol=hh['employee id'];
  const chainCols=[]; for(let i=1;i<=12;i++){const c=hh['management_chain_level_'+String(i).padStart(2,'0')]; if(c!==undefined)chainCols.push(c);}
  if(!chainCols.length) return {ok:false,error:'The Workday tab has no Management_Chain_Level columns.'};
  const idOf=s=>{const m=String(s||'').match(/\((\d+)\)\s*$/);return m?m[1]:null;};
  const idClean=v=>{const s=str_(v);return s?String(Math.round(Number(s.replace(/,/g,'')))||s):'';};
  const nameIds={}, idNm={}, ouOfId={}, chain={};
  const addName=(n,id)=>{const c=cleanName_(n);if(!c||!id)return;const k=k1_(c);(nameIds[k]=nameIds[k]||{})[id]=1;};
  const regCol=hh['region hierarchy'];
  hv.slice(1).forEach(r=>{const id=idClean(r[eidCol]);if(!id)return;const nm=cleanName_(r[hh['preferred name']]);if(nm)idNm[id]=nm;addName(r[hh['preferred name']],id);addName(r[hh['worker']],id);
    if(regCol!==undefined){const reg=str_(r[regCol]).replace(/\\n|\n/g,' ').replace(/\s+/g,' ').trim().toUpperCase();if(reg.endsWith('INDIA'))ouOfId[id]='India';else if(reg.endsWith('ASEAN'))ouOfId[id]='ASEAN';}
    const ids=[];chainCols.forEach(c=>{const v=str_(r[c]);const e=idOf(v);if(e){addName(v,e);if(!idNm[e])idNm[e]=cleanName_(v);if(ids[ids.length-1]!==e)ids.push(e);}});chain[id]=ids.reverse();});
  const ov=ss.getSheetByName(CONFIG.sheets.leaderIds);
  if(ov) ov.getDataRange().getValues().slice(1).forEach(r=>{const n=str_(r[0]),id=idClean(r[1]);if(n&&id)nameIds[k1_(cleanName_(n)||'')]={[id]:1};});
  const idFor=name=>{const c=cleanName_(name)||'',keys=[k1_(c)];
    Object.keys(CONFIG.aliases).forEach(a=>{if(cleanName_(CONFIG.aliases[a])===c)keys.push(k1_(a));}); if(CONFIG.aliases[c])keys.push(k1_(CONFIG.aliases[c]));
    const cand={};keys.forEach(k=>Object.keys(nameIds[k]||{}).forEach(x=>cand[x]=1));const ids=Object.keys(cand);return ids.length===1?ids[0]:null;};
  const ouOf=rg=>CONFIG.budgetRegionMap[str_(rg)]||str_(rg);
  const bvals=bud.getDataRange().getValues(), indiv=individualSet_(bvals);
  const hasBudget={}, indivAt={}, isOwner={}, ownerName={}, unresolved=[];
  bvals.slice(1).forEach(r=>{const n=cleanName_(r[1]);if(!n)return;const id=idFor(n);if(!id){if(unresolved.indexOf(n)<0)unresolved.push(n);return;}
    const k=ouOf(r[0])+'|'+id; hasBudget[k]=1; isOwner[id]=1; ownerName[id]=ownerName[id]||n; if(indiv.has(n)) indivAt[k]=1;});
  return {ok:true,chain,idNm,ouOfId,hasBudget,indivAt,isOwner,ownerName,unresolved,indiv:[...indiv],idFor,hierName:hd.name,people:hd.people};
}

/** Places one person's spend in the budget structure. Returns {kind, L:[Leader, Leader+1, Leader+2]}. */
function assignLeaders_(c,id,ou){
  if(c.chain[id]===undefined) return {kind:'not in Workday',L:['No budget line','','']};
  const path=[id].concat(c.chain[id]); let li=-1;                       // path[0] = the employee, then their managers, nearest first
  if(c.hasBudget[ou+'|'+id]) li=0;
  else for(let k=1;k<path.length;k++){const a=path[k];if(c.hasBudget[ou+'|'+a]&&!c.indivAt[ou+'|'+a]){li=k;break;}}
  if(li<0) return {kind:'no team budget above',L:['No budget line','','']};
  const below=[]; for(let k=li-1;k>=1&&below.length<2;k--) below.push(c.ownerName[path[k]]||c.idNm[path[k]]||path[k]);
  return {kind:li===0?'own budget':'team budget',L:[c.ownerName[path[li]],below[0]||'',below[1]||'']};
}

function syncLeaders_(){
  const ss=SpreadsheetApp.getActive(), spend=ss.getSheetByName(CONFIG.sheets.spend);
  if(!spend) return {ok:false,error:'The "'+CONFIG.sheets.spend+'" tab is missing.'};
  const c=buildLeaderCtx_(); if(!c.ok) return c;
  const sv=spend.getDataRange().getValues(), h=header_(sv[0]);
  for(const n of ['Employee ID','OU','Leader','Leader+1','Leader+2']) if(h[n.toLowerCase()]===undefined) return {ok:false,error:'Column "'+n+'" is missing from the spend tab.'};
  const iL=h['leader'], i1=h['leader+1'], i2=h['leader+2'], iE=h['employee id'], iO=h['ou'], iN=h['employee name'];
  if(i1!==iL+1||i2!==iL+2) return {ok:false,error:'Leader, Leader+1 and Leader+2 must be next to each other in the spend tab.'};
  const idClean=v=>{const s=str_(v);return s?String(Math.round(Number(s.replace(/,/g,'')))||s):'';};
  const stats={'own budget':0,'team budget':0,'no team budget above':0,'not in Workday':0,'blank employee ID':0}, out=[];
  for(let i=1;i<sv.length;i++){
    const r=sv[i], id=idClean(r[iE]), ou=str_(r[iO]);
    if(!id){stats['blank employee ID']++;out.push([r[iL],r[i1],r[i2]]);continue;}
    const a=assignLeaders_(c,id,ou); stats[a.kind]++; out.push(a.L);
  }
  let changed=0; for(let i=0;i<out.length;i++){const r=sv[i+1];if(str_(r[iL])!==str_(out[i][0])||str_(r[i1])!==str_(out[i][1])||str_(r[i2])!==str_(out[i][2]))changed++;}
  if(changed){
    const bk=CONFIG.sheets.spend+' (before leader sync)';
    if(!ss.getSheetByName(bk)) spend.copyTo(ss).setName(bk);
    spend.getRange(2,iL+1,out.length,3).setValues(out);
  }
  let rs=ss.getSheetByName('Leader Sync'); if(!rs)rs=ss.insertSheet('Leader Sync'); rs.clear();
  const sum=[['Leader sync from Workday',''],['Synced at',new Date()],['Workday data used',c.hierName+' ('+c.people+' employees)'],['Rows checked',out.length],['Rows changed in this run',changed],
    ['Charged to their own budget',stats['own budget']],['Charged to the nearest team budget above',stats['team budget']],
    ['No team budget above (shows No budget line)',stats['no team budget above']],['Employee not in the Workday data (shows No budget line)',stats['not in Workday']],
    ['Rows with no Employee ID (left unchanged)',stats['blank employee ID']],['Budgets that cover only the person\'s own spend',c.indiv.join(', ')],
    ['Budget owners not found in Workday',c.unresolved.join(', ')||'None'],['Backup of the spend tab before the first change',"'"+CONFIG.sheets.spend+" (before leader sync)'"]];
  rs.getRange(1,1,sum.length,2).setValues(sum);
  const seen={}, table=[['Employee ID','Employee','OU','How the line was placed','Leader','Leader+1','Leader+2']];
  for(let i=1;i<sv.length;i++){const id=idClean(sv[i][iE]),ou=str_(sv[i][iO]),k=id+'|'+ou;if(!id||seen[k])continue;seen[k]=1;
    table.push([id,iN!==undefined?str_(sv[i][iN]):'',ou,assignLeaders_(c,id,ou).kind,out[i-1][0],out[i-1][1],out[i-1][2]]);}
  ensureRows_(rs, sum.length+2+table.length);
  rs.getRange(sum.length+2,1,table.length,7).setValues(table);
  PropertiesService.getScriptProperties().setProperties({LEADER_SYNC_AT:String(Date.now()),LEADER_SYNC_NOTE:''});
  return {ok:true,rows:out.length,changed,stats,unresolved:c.unresolved,individual:c.indiv};
}
function syncLeadersFromWorkday(){
  const r=syncLeaders_();
  if(!r.ok){PropertiesService.getScriptProperties().setProperty('LEADER_SYNC_NOTE',r.error);try{SpreadsheetApp.getUi().alert(r.error);}catch(e){}return r.error;}
  const msg='Leaders synced: '+r.rows+' rows checked, '+r.changed+' changed. No budget line: '+(r.stats['no team budget above']+r.stats['not in Workday'])+' rows. Details are on the Leader Sync tab.';
  try{SpreadsheetApp.getActive().toast(msg,'Leader sync',15);}catch(e){}
  return msg;
}

/** Optional: load a full Workday export (any region) from a tab called "Workday Import" into "Employee Hierachy".
 *  Keeps India and ASEAN employees and anyone who has spend. Earlier rows for people missing from the new export (for example people who have left, or are on leave) are kept. */
function importWorkdayFromTab(){
  const ss=SpreadsheetApp.getActive(), src=ss.getSheetByName('Workday Import');
  if(!src) throw new Error('Import the Workday file into a tab named "Workday Import" first.');
  const v=src.getDataRange().getValues(), hr=v.findIndex(r=>str_(r[0])==='Employee ID');
  if(hr<0) throw new Error('Could not find the "Employee ID" header in the Workday Import tab.');
  const hdrSrc=v[hr].map(str_), ri=hdrSrc.indexOf('Region Hierarchy');
  if(ri<0) throw new Error('The Workday Import tab has no "Region Hierarchy" column.');
  let dst=ss.getSheetByName(CONFIG.sheets.hierarchy);
  if(!dst){dst=ss.insertSheet(CONFIG.sheets.hierarchy);dst.getRange(1,1,1,HIER_HEADERS.length).setValues([HIER_HEADERS]).setFontWeight('bold');dst.setFrozenRows(1);}
  const dstW=Math.max(dst.getLastColumn(),HIER_HEADERS.length), hdrDst=dst.getRange(1,1,1,dstW).getValues()[0].map(str_);
  const map=hdrDst.map(h=>{let k=hdrSrc.indexOf(h); if(k<0&&/\.1$/.test(h)) k=hdrSrc.lastIndexOf(h.replace(/\.1$/,'')); return k;});
  const idClean=x=>{const s=str_(x);return s?String(Math.round(Number(s.replace(/,/g,'')))||s):'';};
  const sp=ss.getSheetByName(CONFIG.sheets.spend), spend={};
  if(sp){const sv=sp.getDataRange().getValues(), sh=header_(sv[0]), ci=sh['employee id']; if(ci!==undefined) sv.slice(1).forEach(r=>{const id=idClean(r[ci]); if(id) spend[id]=1;});}
  const body=v.slice(hr+1), inSrc={}; body.forEach(r=>{const id=idClean(r[0]); if(id) inSrc[id]=1;});
  const pick=body.filter(r=>{const id=idClean(r[0]); if(!id) return false;
    const reg=str_(r[ri]).replace(/\\n|\n/g,' ').replace(/\s+/g,' ').trim().toUpperCase();
    return WORKDAY_REGIONS.some(x=>reg.endsWith(x))||spend[id];});
  if(pick.length<100) throw new Error('Only '+pick.length+' employees matched. Check that the Workday Import tab is the right export.');
  const out=pick.map(r=>map.map(k=>k>=0?r[k]:''));
  // keep earlier rows for people who are not in the new export at all (for example people who have left)
  const idCol=hdrDst.indexOf('Employee ID'); let kept=0;
  if(dst.getLastRow()>1) dst.getRange(2,1,dst.getLastRow()-1,dstW).getValues().forEach(r=>{const id=idClean(r[idCol]); if(id&&!inSrc[id]){out.push(r);kept++;}});
  ensureRows_(dst,out.length+1);
  const maxR=dst.getMaxRows(); if(maxR>1) dst.getRange(2,1,maxR-1,dstW).clearContent();
  dst.getRange(2,1,out.length,dstW).setValues(out);
  const msg='Loaded '+(out.length-kept)+' employees from the new export and kept '+kept+' earlier rows for people missing from it. Next: Sync leaders from Workday, then Rebuild data now.';
  try{ss.toast(msg,'Workday import',15);}catch(e){}
  return msg;
}

// ===================== LIVE SETTINGS, POLICY AND VARIANCE (read from the spreadsheet) =====================
// Every figure the dashboard uses comes from a tab in this spreadsheet:
//   "Dashboard Settings" : quarter label, quarter start and end, region, event purposes
//   "Policy Guildlines"  : the policy caps, read from the sentences in that tab
//   "Budget"             : budgets and the "Budget covers" column
//   "T&E Q3 QTD", "TA", "Employee Hierachy" : the data
function settingsRows_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Dashboard Settings');
  if(!sh) throw new Error('The "Dashboard Settings" tab is missing. Run T&E Dashboard > Create settings tab once, then check the values.');
  const m={};
  sh.getDataRange().getValues().slice(1).forEach(r=>{const k=str_(r[0]).toLowerCase();if(k)m[k]=r[1];});
  return m;
}
function dateText_(v){
  if(Object.prototype.toString.call(v)==='[object Date]') return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const s=str_(v); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('"'+s+'" is not a date. Use yyyy-mm-dd in the Dashboard Settings tab.');
  return s;
}
function readSettings_(){
  const m=settingsRows_(), need=['quarter label','quarter start','quarter end','region'];
  const miss=need.filter(k=>blank_(m[k])); if(miss.length) throw new Error('Dashboard Settings is missing: '+miss.join(', ')+'.');
  const ev=str_(m['event purposes']).split(',').map(s=>s.trim()).filter(Boolean);
  return {quarterLabel:str_(m['quarter label']), qStart:dateText_(m['quarter start']), qEnd:dateText_(m['quarter end']), region:str_(m['region']), events:ev};
}
/** Reads the caps from the sentences in the Policy Guildlines tab. Fails loudly if a number cannot be found. */
function readPolicy_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Policy Guildlines');
  if(!sh) throw new Error('The "Policy Guildlines" tab is missing.');
  const lines=sh.getDataRange().getValues().map(r=>str_(r[0])).filter(Boolean);
  const text=lines.join('\n');
  const after=(src,key,sym)=>{const i=src.toLowerCase().indexOf(key.toLowerCase()); if(i<0) return null; const m=src.slice(i).match(new RegExp(sym+'\\s*([\\d,]+)')); return m?Number(m[1].replace(/,/g,'')):null;};
  const india=lines.find(l=>/^india\b/i.test(l))||'', asean=lines.find(l=>/^asean\b/i.test(l))||'';
  const P={
    inrAir:after(india,'Airfare','₹'), inrHotel:after(india,'Hotel','₹'),
    aseanAir:after(asean,'Airfare','\\$'), aseanHotel:after(asean,'Hotel','\\$'),
    intlAir:after(text,'Airfare Max Limit','\\$'), intlHotel:after(text,'Hotel Per night Max','\\$'),
    freeze:/FREEZE IN EFFECT/i.test(text), lines:lines
  };
  const bad=['inrAir','inrHotel','aseanAir','aseanHotel','intlAir','intlHotel'].filter(k=>!(P[k]>0));
  if(bad.length) throw new Error('Could not read these caps from the Policy Guildlines tab: '+bad.join(', ')+'. Check the wording of those lines.');
  return P;
}
/** Live variance: budget per budget owner (both regions) against actual spend charged to them, as at the last complete day. */
function writeVarianceTab_(full){
  const D=full.dict, X=full.tx, ss=SpreadsheetApp.getActive();
  const act={}; for(let i=0;i<X.a.length;i++){const k=D.ou[X.ou[i]]+'|'+D.ow[X.ow[i]]; act[k]=(act[k]||0)+X.a[i];}
  const byName={}; full.budget.forEach(b=>{const n=b.n; (byName[n]=byName[n]||{budget:0,actual:0,regions:[]}); byName[n].budget+=b.full; byName[n].actual+=act[b.ou+'|'+n]||0; byName[n].regions.push(b.ou);});
  const elapsed=(full.meta.lastComplete+1)/full.meta.qDays;
  const rows=Object.keys(byName).map(n=>{const r=byName[n];const used=r.budget?r.actual/r.budget:null;const gap=used==null?null:used-elapsed;
    const flags=[]; if(used!=null&&used>1)flags.push('OVER BUDGET'); if(gap!=null&&gap>=0.2)flags.push('AHEAD OF PACE'); if(gap!=null&&gap<=-0.25)flags.push('BEHIND PACE'); if(r.actual<0)flags.push('NET CREDIT');
    return [n,r.regions.join(' + '),Math.round(r.budget),Math.round(r.actual),Math.round(r.actual-r.budget),used==null?'':Math.round(used*1000)/10,gap==null?'':Math.round(gap*1000)/10,flags.join(', ')];}).sort((a,b)=>b[5]-a[5]);
  const totBud=rows.reduce((s,r)=>s+r[2],0), totAct=rows.reduce((s,r)=>s+r[3],0);
  const unbudgeted=full.meta.audit.total-totAct;
  let sh=ss.getSheetByName('Variance Audit'); if(!sh) sh=ss.insertSheet('Variance Audit'); sh.clear();
  const out=[['Actual vs budget by leader (live)',''],['Quarter elapsed',Math.round(elapsed*1000)/10+'%'],['As at last complete day',full.meta.lastDay],[],
    ['Leader','Region','Budget','Actual to date','Variance $','Used %','Pace gap (pts vs elapsed)','Flag']].concat(rows,[[],
    ['Total budget / charged to budgets',"",Math.round(totBud),Math.round(totAct),Math.round(totAct-totBud),Math.round(totAct/totBud*1000)/10,'',''],
    ['Spend with No budget line (not charged to any budget)','',"",Math.round(unbudgeted),'','','',''],
    ['Check: total spend in the data',"",'',Math.round(full.meta.audit.total),'','','','']]);
  sh.getRange(1,1,out.length,8).setValues(out.map(r=>{const x=r.slice();while(x.length<8)x.push('');return x;}));
  return {rows:rows.length,totalActual:Math.round(totAct),unbudgeted:Math.round(unbudgeted)};
}
function setupDashboardSettings(){
  const ss=SpreadsheetApp.getActive();
  if(ss.getSheetByName('Dashboard Settings')){ss.getSheetByName('Dashboard Settings').activate();return 'Dashboard Settings already exists. Check the values.';}
  const sh=ss.insertSheet('Dashboard Settings');
  sh.getRange(1,1,6,2).setValues([['Setting','Value'],['Quarter label','Q3 FY27'],['Quarter start','2026-08-01'],['Quarter end','2026-10-31'],['Region','South Asia'],['Event purposes','Dreamforce, Company kickoff']]);
  sh.getRange(1,1,1,2).setFontWeight('bold');
  sh.getRange(8,1).setValue('Check these values before the first rebuild. Quarter dates are yyyy-mm-dd. Event purposes are the report purposes that are one-off events and are left out of the run-rate.');
  ss.setActiveSheet(sh);
  return 'Created Dashboard Settings. Check the values, then Rebuild data now.';
}

function writeChecks_(full){
  const a=full.meta.audit,ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(CONFIG.sheets.checks);if(!sh)sh=ss.insertSheet(CONFIG.sheets.checks);sh.clear();
  const rows=[['Check','Result'],['Rebuilt at',new Date()],['Expense lines loaded',a.lines],['Rows skipped',a.skippedRows],['Total spend (USD)',Math.round(a.total*100)/100],['Total Q3 budget (USD)',a.budgetTotal],['Travel approvals loaded',a.taRows],['Approved travel (USD, after fixes)',Math.round(a.taTotal)],['TA sheet columns shifted',a.columnShift?'Yes, auto-corrected':'No'],['INR rate used',Math.round(a.inrRate*100)/100],['Unknown expense types',a.unknownTypes.join(', ')||'None'],['Unknown report purposes',a.unknownPurposes.join(', ')||'None'],['Namesakes kept separate',a.namesakes.join(', ')||'None'],['',''],['INR-entered approvals converted',''],['Traveller | route','As entered → USD']].concat(a.inrFixes.map(f=>[f[0]+' | '+f[1],f[2]+' → '+f[3]])).concat([['',''],['Leader sync from Workday', a.leaderSync&&a.leaderSync.ok?('Synced. '+a.leaderSync.changed+' of '+a.leaderSync.rows+' rows changed this time'):('NOT SYNCED: '+((a.leaderSync&&a.leaderSync.error)||'unknown'))],['Budgets that cover only the person\'s own spend', a.leaderSync&&a.leaderSync.individual?a.leaderSync.individual.join(', '):'']]).concat([['',''],['Budget lines as read','Stored value | shown in sheet → USD used']]).concat((a.budgetRows||[]).map(r=>[r[0]+' | '+r[1]+(r[5]!==r[1]?' → '+r[5]:''),r[2]+' | '+r[3]+' → '+r[4]])).concat([['',''],['Names merged','Into']]).concat(a.merges.map(m=>[m[0]+' ('+m[2]+')',m[1]]));
  ensureRows_(sh, rows.length);
  sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#032D60').setFontColor('#FFFFFF');sh.setColumnWidth(1,420);sh.setColumnWidth(2,320);
}

// ===================== MESSAGES (Slack) =====================
// Every message is built from the RECIPIENT'S OWN dashboard view, so it can only contain figures that person can already see.
// The server re-checks every recipient against the sender's access. Nothing is sent without an explicit confirm flag.
const MSG = {
  maxRecipients: 50, maxNote: 500, dailyPerSender: 300, pauseMs: 1100,
  templates: [
    {id:'attention',   title:'Full position and what needs attention', desc:'Budget, unsubmitted spend, travel without approval and policy exceptions in one message.'},
    {id:'budget',      title:'Budget position',                        desc:'Spend charged to the budget, budget used and projected landing for the quarter.'},
    {id:'unsubmitted', title:'Unsubmitted and unapproved spend',       desc:'What still needs submitting or approving, with the oldest items called out.'},
    {id:'travel',      title:'Travel without an approval',             desc:'Travellers with spend but no travel approval on file.'},
    {id:'policy',      title:'Policy exceptions',                      desc:'Flagged spend by exception type.'},
    {id:'note',        title:'Note only',                              desc:'Your note and a link to the dashboard. No figures are added.'}
  ]
};
const MSG_SKIP = ['Employee Attrition','Not in employee master','Outside reporting line','No budget line','Not in spend data','Above your level','Top of reporting line','Own spend or direct reports','No second level'];
const FLAG_LABELS = {air_cap:'Airfare over limit', hotel_cap:'Hotel over nightly limit', premium:'Business or first class', late:'Flight booked 3 days or less ahead', freeze:'Internal travel during freeze', split:'Possible split receipt'};
const TRAVEL_TYPES = ['Airfare','Hotel & Lodging','Room Tax','Hotel Deposit','Subway/Train/Trolley - Cross State/Country'];

function msgEsc_(s){ return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B-\u001F]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function usd_(n){ n = Number(n) || 0; return (n < 0 ? '−' : '') + '$' + Math.round(Math.abs(n)).toLocaleString('en-US'); }
function pct_(x){ return (x == null || !isFinite(x)) ? '–' : (x < 0 ? '−' : '') + Math.round(Math.abs(x) * 100) + '%'; }

/** Forecast for any set of rows, same rule as the dashboard's TA-informed landing (no filters). */
function fcView_(v, rows, tas){
  const X = v.tx, A = v.ta, M = v.meta, D = v.dict;
  const LC = M.lastComplete, REMAIN = M.qDays - (LC + 1);
  const evSet = new Set(M.events), evIdx = D.p.map(x => evSet.has(x));
  const AIRX = D.x.indexOf('Airfare');
  let qtd = 0, air = 0, base = 0;
  for (const i of rows) {
    const a = X.a[i]; qtd += a;
    if (X.d[i] > LC || evIdx[X.p[i]]) continue;
    if (X.x[i] === AIRX) air += a; else if (!X.lk[i]) base += a;
  }
  base = base / (LC + 1) * REMAIN; air = air / (LC + 1) * REMAIN;
  let openNon = 0, unr = 0, openAir = 0;
  for (const j of tas) { openNon += A.open_nonair[j]; unr += A.exp_unraised[j]; openAir += A.open_air[j]; }
  const airF = Math.max(openAir, air);
  return { qtd, base, openNon, openAir, unr, airF, land: qtd + base + openNon + unr + airF };
}

/** Status rule, identical to the Budget owner pacing table. */
function statusOf_(b, q, land){ return !b ? 'No budget' : q > b ? 'Over budget' : land > b ? 'Projected over' : land > 0.92 * b ? 'Watch' : 'On track'; }
const STATUS_RANK = {'Over budget':0, 'Projected over':1, 'Watch':2, 'On track':3, 'No budget':4};
function worstStatus_(list){ return list.slice().sort((a, b) => STATUS_RANK[a] - STATUS_RANK[b])[0] || 'No budget'; }

/** Budget owner pacing for one person, one line per region (exactly the rows of the Budget owner pacing table). */
function msgOwnerPacing_(v, name){
  const bl = (v.budget || []).filter(b => b.n === name);
  if (!bl.length) return null;
  const D = v.dict, X = v.tx, A = v.ta;
  const owi = D.ow.indexOf(name);
  const lines = [];
  let budget = 0, qtd = 0, land = 0, open = 0;
  for (const b of bl) {
    const oui = D.ou.indexOf(b.ou);
    const rows = [], tas = [];
    if (owi >= 0) for (let i = 0; i < X.a.length; i++) if (X.ow[i] === owi && X.ou[i] === oui) rows.push(i);
    if (owi >= 0) for (let j = 0; j < A.e.length; j++) if (A.ow[j] === owi && A.ou[j] === oui) tas.push(j);
    const f = fcView_(v, rows, tas);
    lines.push({ou: b.ou, budget: b.b, qtd: f.qtd, land: f.land, status: statusOf_(b.b, f.qtd, f.land)});
    budget += b.b; qtd += f.qtd; land += f.land; open += f.openNon + f.openAir;
  }
  return { budget, qtd, land, open, used: budget ? qtd / budget : null, status: worstStatus_(lines.map(l => l.status)), lines };
}

/** The figures a person sees in their own view (Policy & controls and Travel pages), computed the same way as the dashboard. */
function msgMetrics_(v){
  const X = v.tx, D = v.dict, A = v.ta, M = v.meta, LAST = M.lastDay;
  const NS = D.s.indexOf('Not Submitted'), APP = D.s.indexOf('Approved');
  const travel = new Set(TRAVEL_TYPES.map(x => D.x.indexOf(x)).filter(i => i >= 0));
  const taNames = new Set(A.e.map(j => D.e[j]));
  let qtd = 0, notApp = 0, notSub = 0, aged = 0, flagged = 0, flagLines = 0;
  const flagPeople = new Set(), flagBy = M.flags.map(() => [0, 0]), noTa = new Map();
  for (let i = 0; i < X.a.length; i++) {
    const a = X.a[i]; qtd += a;
    if (X.s[i] !== APP) notApp += a;
    if (X.s[i] === NS) { notSub += a; if (LAST - X.d[i] > 30) aged += a; }
    if (X.f[i] && a > 0) {
      flagged += a; flagLines++; flagPeople.add(X.e[i]);
      M.flags.forEach((f, b) => { if (X.f[i] & (1 << b)) { flagBy[b][0] += a; flagBy[b][1]++; } });
    }
    if (travel.has(X.x[i]) && a > 0) { const nm = D.e[X.e[i]]; if (!taNames.has(nm)) noTa.set(nm, (noTa.get(nm) || 0) + a); }
  }
  const noTaTop = [...noTa.entries()].map(([name, amount]) => ({ name, amount })).sort((a, b) => b.amount - a.amount);
  const noTaSum = noTaTop.reduce((s, x) => s + x.amount, 0);
  const flagTypes = M.flags.map((f, b) => ({ label: FLAG_LABELS[f], amount: flagBy[b][0], lines: flagBy[b][1] })).filter(x => x.lines > 0).sort((a, b) => b.amount - a.amount);
  return { qtd, notApp, notSub, aged, flagged, flagLines, flagPeople: flagPeople.size, flagTypes, noTaSum, noTaPeople: noTaTop.length, noTaTop };
}

function msgAsAt_(v){
  const ms = Date.parse(v.meta.qStart + 'T00:00:00Z') + v.meta.lastDay * DAY;
  return new Date(ms).toLocaleDateString('en-US', {timeZone:'UTC', month:'short', day:'numeric', year:'numeric'});
}

/** Compact money for Slack, as in the standard format: $191.0K, $1.85M, $950. */
function usdK_(n){ n = Number(n) || 0; const a = Math.abs(n), s = n < 0 ? '−' : ''; if (a >= 1e6) return s + '$' + (a / 1e6).toFixed(2) + 'M'; if (a >= 1000) return s + '$' + (a / 1000).toFixed(1) + 'K'; return s + '$' + Math.round(a).toLocaleString('en-US'); }
function gapK_(land, budget){ const d = land - budget; return d > 0 ? usdK_(d) + ' over budget' : usdK_(-d) + ' under budget'; }

/** Builds the message for one recipient from THEIR view, in the standard format:
 *  header line, bullets, divider, dashboard link, tip. Figures come from the recipient's own view. */
function msgBuild_(tpl, name, v, senderEmail, note){
  const m = msgMetrics_(v), p = msgOwnerPacing_(v, name);
  const q = v.meta.quarterLabel || 'Q3 FY27';
  const url = ScriptApp.getService().getUrl();
  const spendLines = (withLanding) => {
    if (!p) return ['• Spend to date: ' + usdK_(m.qtd) + ' (no budget line is held in your name this quarter)'];
    if (p.lines.length === 1) {
      const L = p.lines[0];
      const out = ['• Spend: ' + usdK_(L.qtd) + ' / ' + usdK_(L.budget) + ' budget — ' + pct_(L.budget ? L.qtd / L.budget : 0) + ' used'];
      if (withLanding) out.push('• Projected landing: ' + usdK_(L.land) + ' — ' + gapK_(L.land, L.budget) + ' (' + L.status + ')');
      return out;
    }
    const out = ['• Spend: ' + usdK_(p.qtd) + ' / ' + usdK_(p.budget) + ' budget — ' + pct_(p.used) + ' used'];
    p.lines.forEach(L => out.push('• Spend (' + L.ou + '): ' + usdK_(L.qtd) + ' / ' + usdK_(L.budget) + ' budget — ' + pct_(L.budget ? L.qtd / L.budget : 0) + ' used' + (withLanding ? ', projected ' + usdK_(L.land) + ' (' + gapK_(L.land, L.budget) + ', ' + L.status + ')' : '')));
    return out;
  };
  const unsubLines = () => ['• Not submitted: ' + usdK_(m.notSub) + ' (older than 30 days: ' + usdK_(m.aged) + ')',
                            '• Not yet approved: ' + usdK_(m.notApp) + ' (' + pct_(m.qtd ? m.notApp / m.qtd : 0) + ' of spend)'];
  const travelLines = () => {
    if (!m.noTaSum) return ['• Travel without an approval: none in your view'];
    const top = m.noTaTop.slice(0, 3).map(x => msgEsc_(x.name) + ' (' + usdK_(x.amount) + ')').join(', ');
    return ['• Travel without an approval: ' + usdK_(m.noTaSum) + ' across ' + m.noTaPeople + ' people. Largest: ' + top + '.'];
  };
  const policyLines = () => {
    if (!m.flagged) return ['• Policy exceptions: none in your view'];
    const types = m.flagTypes.slice(0, 3).map(x => msgEsc_(x.label) + ' ' + usdK_(x.amount)).join('; ');
    return ['• Policy exceptions: ' + usdK_(m.flagged) + ' across ' + m.flagLines + ' lines and ' + m.flagPeople + ' people. By type: ' + types + '.'];
  };
  let bullets;
  switch (tpl) {
    case 'budget':      bullets = spendLines(false); break;
    case 'unsubmitted': bullets = unsubLines(); break;
    case 'travel':      bullets = travelLines(); break;
    case 'policy':      bullets = policyLines(); break;
    case 'attention':   bullets = [...spendLines(true), ...unsubLines(), ...travelLines(), ...policyLines()]; break;
    case 'note':        bullets = []; break;
    default: throw new Error('Unknown message type.');
  }
  const head = [q + ' T&E snapshot.'];
  if (note) head.push('> ' + msgEsc_(note).split('\n').join('\n> '));
  const parts = [head.join('\n')];
  if (bullets.length) parts.push(bullets.join('\n'));
  parts.push('———\nFull breakdown: <' + url + '|T&E Command Center>\nRecommend using the Ask T&E agent on the dashboard for quick drill-downs (top spenders, overage drivers, etc.) — happy to discuss further if helpful.');
  const body = parts.join('\n\n');
  return { title: q + ' T&E snapshot', body, footer: '', text: body,
           blocks: [{type:'section', text:{type:'mrkdwn', text: body.slice(0, 2990)}}] };
}

/** Everyone the sender may message: managers (anyone who leads or has people reporting to them) inside the sender's view. */
function msgCandidates_(auth){
  const full=auth.full, D=full.dict, X=full.tx, anc=full.anc;
  if(!anc||!X.ch) return [];
  const managers=new Set(); Object.keys(anc).forEach(n=>anc[n].forEach(a=>managers.add(a)));      // anyone who has someone reporting to them
  const ns=(full.meta.audit&&full.meta.audit.namesakes)||[];
  const ok=n=>n&&MSG_SKIP.indexOf(n)<0&&anc[n]&&ns.indexOf(n)<0;                               // matched in Workday, and not a name shared by two people
  if(auth.isAdmin) return [...managers].filter(ok).sort();
  const leaders=new Set(auth.scopes.filter(s=>s.type==='leader').map(s=>s.name));
  const regions=new Set(auth.scopes.filter(s=>s.type==='region').map(s=>D.ou.indexOf(s.region)));
  const names=new Set();
  if(leaders.size) Object.keys(anc).forEach(p=>{ if(managers.has(p)&&!leaders.has(p)&&anc[p].some(a=>leaders.has(a))) names.add(p); });
  if(regions.size){ const seen=new Set(); for(let i=0;i<X.a.length;i++){ const ci=X.ch[i]; if(ci<0||seen.has(ci)||!regions.has(X.ou[i])) continue; seen.add(ci); full.chains[ci].forEach(n=>{if(managers.has(n))names.add(n);}); } }
  return [...names].filter(ok).sort();
}

function msgStatusMap_(auth){
  const out = {};
  const seen = new Set((auth.view.budget || []).map(b => b.n));
  seen.forEach(n => { const p = msgOwnerPacing_(auth.view, n); if (p) out[n] = p.status; });
  return out;
}

function msgAuth_(){
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!email) throw new Error('We could not confirm your identity. Open the link while signed in to your work Google account.');
  const scopes = readAccess_()[email];
  if (!scopes || !scopes.length) throw new Error('Your account is not on the Access sheet.');
  const full = loadData_();
  if (!full) throw new Error('The dashboard data has not been built yet. Run T&E Dashboard > Rebuild data now.');
  const isAdmin = scopes.some(s => s.type === 'admin');
  return { email, scopes, isAdmin, full, view: viewFor_(full, scopes) };
}

/** Work email for each name, from Employee Hierachy (active records first). Optional override sheet wins. */
function msgEmailMap_(){
  const map = {};
  const put = (n, em) => {
    const c = cleanName_(n) || ''; const k = k1_(c);
    if (k && em && !map[k]) map[k] = em;
    const t = toks_(c);
    if (t.length >= 2 && em && !map['FL:' + t[0] + '|' + t[t.length-1]]) map['FL:' + t[0] + '|' + t[t.length-1]] = em;
  };
  try {
    const hv = hierarchyData_().values, h = header_(hv[0]||[]);
    const ci = h['preferred name'], cw = h['worker'], ce = h['primary work email'], ca = h['currently active'];
    if (ce !== undefined) {
      const rows = hv.slice(1).filter(r => str_(r[ce]).indexOf('@') > 0);
      const active = r => ca === undefined || /^(y|yes|true|1)$/i.test(str_(r[ca]));
      [rows.filter(active), rows.filter(r => !active(r))].forEach(group => group.forEach(r => {
        const em = str_(r[ce]).toLowerCase();
        if (ci !== undefined) put(str_(r[ci]), em);
        if (cw !== undefined) put(str_(r[cw]), em);
      }));
    }
  } catch (e) {}
  try {
    sheetValues_(CONFIG.sheets.emails).slice(1).forEach(r => {
      const n = str_(r[0]), em = str_(r[1]).toLowerCase();
      if (n && em.indexOf('@') > 0) { const k = k1_(cleanName_(n) || ''); if (k) map[k] = em; }
    });
  } catch (e) {}
  return map;
}
function msgEmailFor_(map, name){
  const c = cleanName_(name) || '';
  if (map[k1_(c)]) return map[k1_(c)];
  const t = toks_(c);
  if (t.length >= 2 && map['FL:' + t[0] + '|' + t[t.length-1]]) return map['FL:' + t[0] + '|' + t[t.length-1]];
  for (const a of Object.keys(CONFIG.aliases)) if (cleanName_(CONFIG.aliases[a]) === c && map[k1_(a)]) return map[k1_(a)];
  return null;
}

function msgValidate_(auth, req){
  const tpl = String((req && req.template) || '');
  if (!MSG.templates.some(t => t.id === tpl)) throw new Error('Choose a message type.');
  const names = Array.isArray(req.recipients) ? [...new Set(req.recipients.map(n => String(n || '').trim()).filter(Boolean))] : [];
  if (!names.length) throw new Error('Choose at least one recipient.');
  if (names.length > MSG.maxRecipients) throw new Error('Send to at most ' + MSG.maxRecipients + ' people at a time.');
  const allowed = new Set(msgCandidates_(auth));
  if (names.some(n => !allowed.has(n))) throw new Error('Some recipients are outside your access. Nothing was sent.');
  const note = String(req.note || '').replace(/[\u0000-\u0008\u000B-\u001F]/g,'').trim();
  if (note.length > MSG.maxNote) throw new Error('Notes are limited to ' + MSG.maxNote + ' characters.');
  if (tpl === 'note' && !note) throw new Error('Write a note for the note-only message.');
  return { names, tpl, note };
}

function msgItems_(auth, names, tpl, note, emap){
  const cache = {};
  return names.map(name => {
    cache[name] = cache[name] || subsetPayload_(auth.full, [name]);
    const b = msgBuild_(tpl, name, cache[name], auth.email, note);
    return { name, title: b.title, body: b.body, footer: b.footer, text: b.text, blocks: b.blocks, hasEmail: !!msgEmailFor_(emap, name) };
  });
}

function slackToken_(){
  const t = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!t) throw new Error('Slack is not connected yet. The owner adds the bot token as SLACK_BOT_TOKEN in Project Settings > Script properties.');
  return t;
}
function slackApi_(method, params, isGet){
  const headers = {Authorization: 'Bearer ' + slackToken_()};
  let res;
  if (isGet) {
    const q = Object.keys(params).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
    res = UrlFetchApp.fetch('https://slack.com/api/' + method + '?' + q, {method:'get', headers, muteHttpExceptions:true});
  } else {
    res = UrlFetchApp.fetch('https://slack.com/api/' + method, {method:'post', contentType:'application/json; charset=utf-8', headers, payload:JSON.stringify(params), muteHttpExceptions:true});
  }
  try { return JSON.parse(res.getContentText()); } catch (e) { return {ok:false, error:'bad_response'}; }
}

function msgLogSheet_(create){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.sheets.messageLog);
  if (!sh && create) {
    sh = ss.insertSheet(CONFIG.sheets.messageLog);
    sh.appendRow(['Time','Sent by','Recipient','Message type','Result','Detail']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function msgCountToday_(email){
  const sh = msgLogSheet_(false); if (!sh) return 0;
  const today = new Date().toDateString();
  return sh.getDataRange().getValues().slice(1).filter(r => String(r[1]).toLowerCase() === email && String(r[4]) === 'Sent' && new Date(r[0]).toDateString() === today).length;
}

function msgDeliver_(auth, names, tpl, note, emap){
  slackToken_();                                 // fail before any send if Slack is not connected
  const items = msgItems_(auth, names, tpl, note, emap), out = [];
  items.forEach((it, idx) => {
    const res = {name: it.name, status: 'Failed', detail: ''};
    if (!it.hasEmail) res.detail = 'No work email on file. Add it on the Message Emails sheet.';
    else {
      if (idx > 0) Utilities.sleep(MSG.pauseMs);
      const em = msgEmailFor_(emap, it.name);
      const u = slackApi_('users.lookupByEmail', {email: em}, true);
      if (!u.ok) res.detail = 'Slack lookup: ' + (u.error || 'failed');
      else {
        const o = slackApi_('conversations.open', {users: u.user.id}, false);
        if (!o.ok) res.detail = 'Open chat: ' + (o.error || 'failed');
        else {
          const p = slackApi_('chat.postMessage', {channel: o.channel.id, text: it.text, blocks: it.blocks, unfurl_links:false, unfurl_media:false}, false);
          if (p.ok) { res.status = 'Sent'; res.detail = 'Delivered'; } else res.detail = 'Post: ' + (p.error || 'failed');
        }
      }
    }
    msgLogSheet_(true).appendRow([new Date(), auth.email, it.name, tpl, res.status, res.detail]);
    out.push(res);
  });
  return out;
}

/** Public (called from the page). Returns who the signed-in user can message and the message types. */
function apiMsgContext(){
  try {
    const auth = msgAuth_(), emap = msgEmailMap_(), stat = msgStatusMap_(auth);
    const budgetVisible = new Set((auth.view.budget || []).map(b => b.n));
    const order = {'Over budget':0, 'Projected over':1, 'Watch':2, 'On track':3};
    const rows = msgCandidates_(auth).map(n => ({name:n, kind: budgetVisible.has(n) ? 'Budget holder' : 'Manager', status: stat[n] || null, hasEmail: !!msgEmailFor_(emap, n)}))
      .sort((a, b) => ((order[a.status] ?? 4) - (order[b.status] ?? 4)) || a.name.localeCompare(b.name));
    const role = auth.isAdmin ? 'Admin' : (auth.scopes.some(s => s.type === 'region') ? 'Region' : 'Leader');
    return JSON.stringify({ok:true, role, email:auth.email, recipients:rows, templates:MSG.templates,
      limits:{maxRecipients:MSG.maxRecipients, maxNote:MSG.maxNote}, slackConnected: !!PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'),
      sentToday: msgCountToday_(auth.email), dailyLimit: MSG.dailyPerSender, dashboardUrl: ScriptApp.getService().getUrl()});
  } catch (e) { return JSON.stringify({ok:false, error:String(e.message || e)}); }
}

/** Public. Shows exactly what each recipient will receive. Sends nothing. */
function apiMsgPreview(req){
  try {
    const auth = msgAuth_(), v = msgValidate_(auth, req || {}), emap = msgEmailMap_();
    const items = msgItems_(auth, v.names, v.tpl, v.note, emap).map(it => ({name:it.name, title:it.title, body:it.body, footer:it.footer, hasEmail:it.hasEmail}));
    return JSON.stringify({ok:true, items});
  } catch (e) { return JSON.stringify({ok:false, error:String(e.message || e)}); }
}

/** Public. Sends only when the page passes confirm:true (the user has ticked the confirmation). */
function apiMsgSend(req){
  try {
    const auth = msgAuth_();
    req = req || {};
    if (req.confirm !== true) throw new Error('Nothing was sent. Confirm the recipients and message first.');
    const v = msgValidate_(auth, req);
    if (msgCountToday_(auth.email) + v.names.length > MSG.dailyPerSender) throw new Error('Daily message limit reached (' + MSG.dailyPerSender + ' a day).');
    const results = msgDeliver_(auth, v.names, v.tpl, v.note, msgEmailMap_());
    return JSON.stringify({ok:true, results});
  } catch (e) { return JSON.stringify({ok:false, error:String(e.message || e)}); }
}

/** Public. Admins see every message; everyone else sees their own. */
function apiMsgLog(){
  try {
    const auth = msgAuth_(), sh = msgLogSheet_(false);
    if (!sh) return JSON.stringify({ok:true, rows:[]});
    const rows = sh.getDataRange().getValues().slice(1)
      .filter(r => auth.isAdmin || String(r[1]).toLowerCase() === auth.email)
      .slice(-100).reverse()
      .map(r => [r[0] instanceof Date ? r[0].toISOString() : String(r[0]), r[1], r[2], r[3], r[4], r[5]]);
    return JSON.stringify({ok:true, rows});
  } catch (e) { return JSON.stringify({ok:false, error:String(e.message || e)}); }
}
