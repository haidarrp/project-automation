(function(){
'use strict';
const app=document.getElementById('app');
const access=window.UserAccess;
const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');
const state={loading:true,error:'',users:[],admins:new Set(),records:new Map(),search:'',role:'all',selected:'',module:'cuti',draft:null,revision:0,dirty:false,busy:false,notice:''};
const modules=access.MODULES;
const db=()=>window.FirebaseClient.getDb();
const uid=()=>window.FirebaseClient.getCurrentUser()?.uid||'';
function dateLabel(v){
 if(!v)return 'Belum tercatat';
 const date=typeof v.toDate==='function'?v.toDate():new Date(v);
 return Number.isNaN(date.getTime())?'Belum tercatat':date.toLocaleString('id-ID',{dateStyle:'medium',timeStyle:'short'});
}
function moduleKeys(m){return m.sections.flatMap(s=>[s.id+'.view',...s.actions.map(a=>s.id+'.'+a[0])]);}
function cleaned(moduleId,values){const source={};source[moduleId]=values;return access.normalizeGrants(source)[moduleId];}
function effective(record){return record?access.normalizeGrants(record.grants||{}):access.defaultGrants();}
function roleOf(user){return state.admins.has(user.id)?'Admin':'User';}
function currentRow(){return state.users.find(u=>u.id===state.selected);}
function countOf(g,mid){return Object.values(g?.[mid]||{}).filter(v=>v===true).length;}
function totalOf(g){return modules.reduce((n,m)=>n+countOf(g,m.id),0);}
function initials(name){return String(name||'U').split(/\s+/).slice(0,2).map(s=>s[0]||'').join('').toUpperCase();}
function render(){
 const rows=state.users.filter(u=>{
   const s=state.search.trim().toLowerCase();
   return (!s||[u.displayName,u.email].some(v=>String(v||'').toLowerCase().includes(s)))&&(state.role==='all'||(state.role==='admin')===state.admins.has(u.id));
 });
 const assigned=state.users.filter(u=>state.records.has(u.id)&&!state.admins.has(u.id)).length;
 const metrics=`<div class="um-metrics"><div class="card um-metric"><span>Pengguna Terdaftar</span><strong>${state.users.length}</strong></div><div class="card um-metric"><span>Administrator</span><strong>${state.admins.size}</strong></div><div class="card um-metric"><span>Izin Dikustomisasi</span><strong>${assigned}</strong></div><div class="card um-metric"><span>Akun Dinonaktifkan</span><strong>${state.users.filter(u=>state.records.get(u.id)?.active===false).length}</strong></div></div>`;
 const rowHtml=rows.length?rows.map(u=>{
   const admin=state.admins.has(u.id);
   const record=state.records.get(u.id);
   const grants=admin?{}:effective(record);
   const chips=admin?'<span class="um-chip um-chip-admin">Seluruh modul + Administrasi</span>':modules.filter(m=>countOf(grants,m.id)).map(m=>`<span class="um-chip">${esc(m.label)} · ${countOf(grants,m.id)}/${moduleKeys(m).length}</span>`).join('')||'<span class="um-muted">Belum ada izin fitur</span>';
   const status=admin?'Aktif':record?.active===false?'Nonaktif':'Aktif';
   return `<tr><td><div class="um-person"><span class="um-avatar">${esc(initials(u.displayName||u.email))}</span><span><strong>${esc(u.displayName||u.email?.split('@')[0]||'Pengguna')}</strong><small>${esc(u.email||'Email belum tercatat')}</small></span></div></td><td><span class="um-role ${admin?'admin':''}">${admin?'Admin':'User'}</span></td><td><div class="um-chips">${chips}</div>${!record&&!admin?'<small class="um-legacy">Akses bawaan lama (belum dikustomisasi)</small>':''}</td><td><span class="um-status ${status==='Nonaktif'?'off':''}">${status}</span></td><td><span class="um-date">${esc(dateLabel(u.lastLoginAt))}</span></td><td><button type="button" class="btn btn-secondary btn-sm" data-edit-user="${esc(u.id)}">${admin?'Lihat akses':'Kelola Akses'}</button></td></tr>`;
 }).join(''):'<tr><td colspan="6" class="um-empty">Tidak ada pengguna yang sesuai pencarian.</td></tr>';
 const content=`<div class="um-page"><div class="um-heading"><div><h1>Pengelolaan User</h1><p>Pengaturan akses pengguna hingga tingkat submenu dan tindakan. Hanya administrator yang dapat mengubah izin.</p></div><span class="um-admin-tag">Khusus Administrator</span></div>${metrics}${state.error?`<div class="alert alert-danger">${esc(state.error)}</div>`:''}${state.notice?`<div class="alert alert-success" role="status">${esc(state.notice)}</div>`:''}
 <section class="card um-list"><div class="um-section-head"><div><h2>Daftar Pengguna</h2><p>Menampilkan akun yang pernah masuk dan tercatat di sistem.</p></div><button class="btn btn-secondary btn-sm" type="button" data-refresh>Segarkan</button></div><div class="um-toolbar"><input id="um-search" class="um-search" type="search" placeholder="Cari nama atau email..." value="${esc(state.search)}"><select id="um-role" aria-label="Filter jenis pengguna"><option value="all" ${state.role==='all'?'selected':''}>Semua jenis</option><option value="admin" ${state.role==='admin'?'selected':''}>Admin</option><option value="user" ${state.role==='user'?'selected':''}>User</option></select><span class="um-count">${rows.length} akun</span></div><div class="um-table-wrap"><table class="data-table um-table"><thead><tr><th>Pengguna</th><th>Jenis</th><th>Hak Akses Fitur</th><th>Status</th><th>Login Terakhir</th><th>Aksi</th></tr></thead><tbody>${rowHtml}</tbody></table></div></section>
 ${state.loading?'<div class="um-loading">Mengambil data pengguna dan izin dari Firestore...</div>':''}
 </div>`;
 const overlay=state.selected?editorMarkup():'';
 app.innerHTML=window.AppShell.render({module:'admin',view:'users',viewLabel:'Pengelolaan User',content,overlays:overlay});
}
function editorMarkup(){
 const row=currentRow();if(!row||!state.draft)return '';
 const admin=state.admins.has(row.id);
 const grants=state.draft.grants;
 const m=access.MOD[state.module];
 const selectedCount=admin?moduleKeys(m).length:countOf(grants,m.id);
 const list=modules.map(mod=>`<button type="button" data-module="${mod.id}" class="um-module-item ${state.module===mod.id?'active':''}"><span>${esc(mod.label)}</span><small>${admin?'Penuh':countOf(grants,mod.id)+'/'+moduleKeys(mod).length}</small></button>`).join('');
 const cards=m.sections.map(s=>{
   const selected=Boolean(grants[m.id]?.[s.id+'.view']);
   const parentOK=!s.parent||Boolean(grants[m.id]?.[s.parent+'.view']);
   const actions=s.actions.map(([act,label])=>`<label class="um-check"><input type="checkbox" data-check="${esc(m.id)}|${esc(s.id+'.'+act)}" ${admin||grants[m.id]?.[s.id+'.'+act]?'checked':''} ${admin||!parentOK||!selected?'disabled':''}><span>${esc(label)}</span></label>`).join('');
   return `<div class="um-permission-card ${!parentOK?'disabled':''}"><div class="um-permission-head"><div><strong>${esc(s.label)}</strong>${s.parent?`<small>Subfitur dari ${esc(m.sections.find(x=>x.id===s.parent)?.label||'')}</small>`:''}</div><label class="um-check um-view"><input type="checkbox" data-check="${esc(m.id)}|${esc(s.id+'.view')}" ${selected||admin?'checked':''} ${admin||!parentOK?'disabled':''}><span>Lihat/Akses</span></label></div>${actions?`<div class="um-permission-actions">${actions}</div>`:'<p class="um-readonly">Izin melihat submenu saja.</p>'}</div>`;
 }).join('');
 return `<div class="um-overlay" data-close-editor><div class="um-dialog" role="dialog" aria-modal="true" aria-labelledby="um-dialog-title"><div class="um-editor-header"><div><button class="um-back" type="button" data-close-button>← Kembali ke daftar</button><h2 id="um-dialog-title">Pengaturan Hak Akses</h2><p>Centang izin secara terpisah untuk melihat, input, edit, hapus, dan tindakan lainnya.</p></div><div class="um-editor-actions"><span class="um-dirty">${state.dirty?'Perubahan belum disimpan':''}</span><button class="btn btn-secondary" type="button" data-close-button>Batal</button><button class="btn btn-primary" type="button" data-save-user ${!state.dirty||admin||state.busy?'disabled':''}>${state.busy?'Menyimpan...':'Simpan Perubahan'}</button></div></div>
 <div class="card um-editor-identity"><div class="um-person"><span class="um-avatar um-avatar-large">${esc(initials(row.displayName||row.email))}</span><span><strong>${esc(row.displayName||'Pengguna')}</strong><small>${esc(row.email||'')}</small></span></div><span class="um-role ${admin?'admin':''}">${admin?'Administrator':'User'}</span>${!admin?`<label class="um-account"><input type="checkbox" id="um-active" ${state.draft.active?'checked':''}> Akun dapat menggunakan aplikasi</label>`:'<span class="um-lock">Izin Admin dikunci pada Firebase Console</span>'}</div>
 <div class="um-editor-grid"><aside class="card um-module-panel"><div class="um-module-caption">PILIH MODUL</div>${list}<p class="um-module-help">Setiap perubahan tetap tersimpan dalam draft saat berpindah modul, sampai tombol Simpan Perubahan dipilih.</p></aside>
 <section class="card um-module-editor"><div class="um-permissions-title"><div><small>PENGATURAN IZIN MODUL</small><h3>${esc(m.label)}</h3><p>${selectedCount} dari ${moduleKeys(m).length} izin dipilih</p></div>${!admin?`<label class="um-toggle"><input id="um-module-on" type="checkbox" ${selectedCount?'checked':''}> Aktifkan Modul</label>`:''}</div>
 ${admin?'<p class="um-info">Akun administrator mempunyai akses penuh ke semua menu. Status Admin tidak dapat diedit dari halaman ini.</p>':`<div class="um-presets"><span>Pengaturan cepat</span><button class="btn btn-secondary btn-sm" data-preset="view" type="button">Hanya Lihat</button><button class="btn btn-secondary btn-sm" data-preset="input" type="button">Lihat + Input</button><button class="btn btn-secondary btn-sm" data-preset="all" type="button">Semua Izin</button><button class="btn btn-secondary btn-sm" data-preset="off" type="button">Nonaktifkan Modul</button></div><p class="um-info">Izin tindakan hanya berlaku jika izin <strong>Lihat/Akses</strong> pada submenu dan submenu induknya aktif. Perubahan baru berlaku setelah disimpan.</p>`}
 <div class="um-permission-grid">${cards}</div><div class="um-summary"><span>Keseluruhan: <strong>${totalOf(grants)} izin</strong> dari lima modul</span><span>${state.draft.active?'Akun aktif':'Akun nonaktif'}</span></div></section></div></div></div>`;
}
async function load(){
 state.loading=true;state.error='';render();
 try{
  const [u,a,r]=await Promise.all([db().collection('users').get(),db().collection('admins').get(),db().collection('userPermissions').get()]);
  state.users=u.docs.map(d=>({id:d.id,...d.data()})).sort((x,y)=>String(x.displayName||x.email).localeCompare(String(y.displayName||y.email),'id'));
  state.admins=new Set(a.docs.filter(d=>d.data()?.active===true).map(d=>d.id));
  state.records=new Map(r.docs.map(d=>[d.id,d.data()]));
 }catch(e){state.error='Tidak dapat memuat daftar pengguna: '+(e?.message||e);}
 finally{state.loading=false;render();}
}
function open(id){const row=state.users.find(u=>u.id===id);if(!row)return;
 state.selected=id;state.module='cuti';state.draft={active:state.records.get(id)?.active!==false,grants:effective(state.records.get(id))};
 state.revision=Number(state.records.get(id)?.revision||0);state.dirty=false;render();
}
function close(){if(state.dirty&&!confirm('Perubahan izin belum disimpan. Keluar tanpa menyimpan?'))return;state.selected='';state.draft=null;state.dirty=false;render();}
function setGrant(module,key,checked){
 const m=access.MOD[module];const section=m?.sections.find(s=>key.startsWith(s.id+'.'));if(!section)return;
 const current=state.draft.grants[module];
 if(checked){current[`${section.id}.view`]=true;if(section.parent)current[`${section.parent}.view`]=true;current[key]=true;}
 else {delete current[key];if(key.endsWith('.view')){
   const sectionId=key.slice(0,-5);
   for(const k of Object.keys(current))if(k.startsWith(sectionId+'.'))delete current[k];
   for(const child of m.sections.filter(s=>s.parent===sectionId))for(const k of Object.keys(current))if(k.startsWith(child.id+'.'))delete current[k];
 }}
 state.draft.grants[module]=cleaned(module,current);state.dirty=true;
}
function preset(module,type){
 const m=access.MOD[module];const values={};
 for(const s of m.sections){
   if(type==='off')continue;
   values[s.id+'.view']=true;
   for(const [id] of s.actions){
     const okay=type==='all'||(type==='input'&&['create','upload'].includes(id));
     if(okay)values[s.id+'.'+id]=true;
   }
 }
 state.draft.grants[module]=cleaned(module,values);state.dirty=true;render();
}
async function save(){
 if(!state.dirty||state.busy||!state.selected||state.admins.has(state.selected))return;
 if(!state.draft.active&& !confirm('Nonaktifkan akses aplikasi untuk pengguna ini?'))return;
 const row=currentRow();if(!row)return;
 state.busy=true;render();
 const ref=db().collection('userPermissions').doc(row.id);
 const audit=db().collection('adminAudit').doc();
 try{
   const changes=access.normalizeGrants(state.draft.grants);
   const permissionCount=totalOf(changes);
   const actor=window.FirebaseClient.getCurrentUser();
   await db().runTransaction(async transaction=>{
     const old=await transaction.get(ref);
     const current=Number(old.data()?.revision||0);
     if(current!==state.revision)throw new Error('Izin sudah diubah administrator lain. Muat ulang daftar dan ulangi perubahan.');
     transaction.set(ref,{uid:row.id,active:state.draft.active,grants:changes,revision:current+1,updatedAt:firebase.firestore.FieldValue.serverTimestamp(),updatedByUid:actor.uid,updatedByEmail:actor.email||'',schemaVersion:1},{merge:true});
     transaction.set(audit,{action:'UPDATE_USER_PERMISSIONS',target:`userPermissions/${row.id}`,detail:`Akses ${row.email||row.id}: ${permissionCount} izin; status ${state.draft.active?'aktif':'nonaktif'}; revisi ${current+1}`,actorUid:actor.uid,actorEmail:actor.email||'',createdAt:firebase.firestore.FieldValue.serverTimestamp()});
   });
   state.records.set(row.id,{active:state.draft.active,grants:changes,revision:state.revision+1});
   state.selected='';state.dirty=false;state.draft=null;state.notice='Hak akses berhasil disimpan.';
   render();setTimeout(()=>{state.notice='';render();},3400);
 }catch(e){alert('Gagal menyimpan izin: '+(e?.message||e));}
 finally{state.busy=false;render();}
}
app.addEventListener('click',e=>{
 if(e.target.closest('[data-close-editor]')&&e.target.matches('[data-close-editor]')){close();return;}
 if(e.target.closest('[data-close-button]')){close();return;}
 const edit=e.target.closest('[data-edit-user]');if(edit){open(edit.dataset.editUser);return;}
 const m=e.target.closest('[data-module]');if(m){state.module=m.dataset.module;render();return;}
 const p=e.target.closest('[data-preset]');if(p){preset(state.module,p.dataset.preset);return;}
 if(e.target.closest('[data-save-user]')){save();return;}
 if(e.target.closest('[data-refresh]')){load();return;}
});
app.addEventListener('input',e=>{
 if(e.target?.id==='um-search'){const at=e.target.selectionStart;state.search=e.target.value;render();const next=document.getElementById('um-search');next?.focus();try{next?.setSelectionRange(at,at);}catch(_){} }
});
app.addEventListener('change',e=>{
 if(e.target?.id==='um-role'){state.role=e.target.value;render();return;}
 if(e.target?.id==='um-active'){state.draft.active=e.target.checked;state.dirty=true;render();return;}
 if(e.target?.id==='um-module-on'){
  if(e.target.checked){if(!countOf(state.draft.grants,state.module))setGrant(state.module,'dashboard.view',true);}
  else preset(state.module,'off');render();return;
 }
 if(e.target?.dataset?.check){const [m,key]=e.target.dataset.check.split('|');setGrant(m,key,e.target.checked);render();}
});
async function init(){await window.FirebaseClient.requireAdmin();await access.load();await load();}
init();
})();
