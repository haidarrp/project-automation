'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const root=path.join(__dirname,'..');
const script=(name)=>fs.readFileSync(path.join(root,'assets','js',name),'utf8');
const w={APP_CONFIG:{RULES:{ALLOW_WEEKEND_OVERTIME:true,MIN_OVERTIME_HOURS:1,MAX_OVERTIME_HOURS:4,REQUIRE_WFO_STATUS:true,NORMAL_END_FRIDAY_MINUTES:990,NORMAL_END_MON_THU_MINUTES:960,NORMAL_START_MINUTES:450,LATEST_OVERTIME_ARRIVAL_MINUTES:540,MEAL_ALLOWANCE_MIN_HOURS:2}}};
vm.runInNewContext(script('business-rules.js'),{window:w,Date,Set,Map});
vm.runInNewContext(script('uang-makan-rules.js'),{window:w,Date,Set,Map});
const LM=require('../assets/js/lembur-document-integration.js');
const policy={allows:()=>true};
const employee=(status,date)=>({id:'p1',records:{[date]:{status,date:new Date(...date.split('-').map((v,i)=>i===1?Number(v)-1:Number(v))),overtimeHours:3,originalOvertimeHours:3,manualOverride:true}}});

test('Tukin vs uang makan vs lembur: seluruh Sabtu dan Minggu tidak memberikan uang makan',()=>{
 for(const date of ['2026-11-07','2026-11-08']){
  for(const status of ['Offsite','tidak_hadir','WFO']){
   const day=w.UangMakanRules.classifyDay({key:date,status,inMinutes:450,outMinutes:1140,manualValue:1,reviewed:true},[],[]);
   assert.equal(day.isWeekend,true);
   assert.equal(day.statusDocumentMissing,false);
   assert.equal(day.automatic,0);
   assert.equal(day.value,0);
   assert.equal(day.manualValue,null);
  }
 }
});

test('Uang Makan Sabtu/Minggu tetap nol meskipun dokumen dan review menyetujui 1 hari',()=>{
 for(const date of ['2026-11-07','2026-11-08']){
  const record={key:date,status:'Offsite',inMinutes:450,outMinutes:1140,dateManualValue:1};
  const task={id:'t1',employeeIds:['p1'],startDate:date,endDate:date,accountable:true};
  const output=w.UangMakanRules.applySupporting({id:'p1',days:{[date]:record},supportReviews:{'tugas:t1':{choice:1,signature:w.UangMakanRules.supportSignature('tugas',task)}}},[],[task],2026,11);
  assert.equal(output.days[date].value,0);
  assert.equal(output.days[date].nonWorking,true);
  assert.equal(w.UangMakanRules.overview(output).days,0);
 }
});

test('Lembur Sabtu/Minggu Offsite/Tidak Hadir 0 jam bahkan dengan bukti pendukung',()=>{
 for(const date of ['2026-11-07','2026-11-08']){
  for(const status of ['Offsite','tidak_hadir']){
   const emp=employee(status,date);
   const task={id:'t1',employeeIds:['p1'],startDate:date,endDate:date,accountable:false};
   const result=LM.applyDocuments([emp],[],[task],policy,e=>e.id);
   const rec=emp.records[date];
   assert.equal(result.held,1);
   assert.equal(rec.statusDocumentNeedsReview,false);
   assert.equal(rec.statusWeekendIneligible,true);
   assert.equal(rec.documentBlocksOvertime,true);
   assert.equal(rec.overtimeHours,0);
   const g=LM.reviewGroups([emp])[0];
   LM.verifySource([emp],g.key,g.signature,'allow');
   assert.equal(rec.documentBlocksOvertime,true);
   assert.equal(rec.overtimeHours,0);
   assert.equal(w.APP_CONFIG.RULES.ALLOW_WEEKEND_OVERTIME,true);
   const calc=w.BusinessRules.calculateOvertime(w.BusinessRules.dateFromKey(date),450,1140,status,[]);
   assert.equal(calc.hours,0);
  }
 }
});

test('Lembur WFO pada akhir pekan tetap mengikuti aturan presensi normal',()=>{
 for(const date of ['2026-11-07','2026-11-08']){
  const calc=w.BusinessRules.calculateOvertime(w.BusinessRules.dateFromKey(date),450,1140,'WFO',[]);
  assert.equal(calc.hours,4);
  const emp=employee('WFO',date);
  LM.applyDocuments([emp],[],[],policy,e=>e.id);
  assert.equal(emp.records[date].documentBlocksOvertime,false);
  assert.equal(emp.records[date].overtimeHours,3);
 }
});

test('Lembur Jumat atau hari kerja Offsite tanpa dokumen tetap Perlu Cek',()=>{
 const date='2026-11-06';
 const emp=employee('Offsite',date);
 LM.applyDocuments([emp],[],[],policy,e=>e.id);
 assert.equal(emp.records[date].statusDocumentNeedsReview,true);
 assert.equal(emp.records[date].statusWeekendIneligible,false);
 assert.equal(emp.records[date].overtimeHours,0);
});

test('Lembur arsip lama: ekspor positif Offsite akhir pekan dilarang dan UI menunjukkan blokir',()=>{
 const app=script('app.js');
 assert.match(app,/window\.LemburDocumentIntegration\.isWeekendDate\(date\)/);
 assert.match(app,/statusWeekendIneligible/);
 assert.match(app,/Tidak dapat diklaim - Status Sabtu\/Minggu/);
 assert.match(app,/documentBlocksOvertime && value !== 0/);
 assert.match(app,/refreshStatusRestriction\(record,date,state\.holidays\)/);
});
