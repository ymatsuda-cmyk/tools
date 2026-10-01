const C=require('../core.js');const S=require('./sheets.json');const assert=require('assert');
const sources=[];
for(const [name,v] of Object.entries(S)){ if(C.isSystemSheet(name))continue; const h=C.findHeader(v); if(!h){console.log('skip(no header)',name);continue}
 if(h.isDaily){console.log('skip(daily)',name);continue} const rec=C.parseRecords(v,h); sources.push({name,records:rec}); console.log(name,'header row',h.row+1,'rows',rec.length,'total',rec.reduce((a,r)=>a+r.n,0),'note',h.noteCol);}
assert.equal(sources.length,2);
const p=C.primarySource(sources);assert.equal(p.name,'202601から202608');
const cust=C.collectCustomers(sources,p);console.log('customers',cust.length);
const sug=C.suggestGroups(cust);const gm=sug;
const rows=C.buildInitialGroupRows(cust,sug);console.log('first rows',rows.slice(0,3));
const nm=C.buildNoteMap(sources);
const a=C.aggregate(p.records,gm,nm,'group'),b=C.aggregate(p.records,gm,nm,'customer');
console.log('units',a.units.length,b.units.length,'top10',C.topShare(a,10).toFixed(1),C.topShare(b,10).toFixed(1),'top20',C.topShare(a,20).toFixed(1),C.topShare(b,20).toFixed(1),'linked',a.linked,b.linked);
assert.equal(a.total,24601);assert.equal(b.linked,4503);assert.equal(a.linked,4503);
console.log(a.units.slice(0,6).map(u=>[u.name,u.n,u.members.length,u.linked]));
const j=C.aggregate(sources[1].records,gm,nm,'group');console.log('July total',j.total,'linked',j.linked,'top',j.units[0].name);
// status
const gs=[['請求先CD','請求先','グループ名','受注件数','自動']].concat(rows.map(r=>[r.cd,r.name,r.group,r.n,r.group]));
gs[2][2]='手入力';const st=C.groupStatus(C.parseGroupSheet(gs));console.log(st.auto,st.confirmed,st.blank,st.groups);
console.log(C.groupKey('ＡＮＴＩＣＡ ＩＴＡＬＩＡＮＡ 大阪店'),C.groupKey('U660-R37 株式会社ﾆｯｶﾈ 宇都宮支店'),C.groupKey('㈱ﾄｰﾎｰﾌｰﾄﾞｻｰﾋﾞｽ　佐世保支店'));
const nw=C.suggestForNew([{cd:99999,name:'株式会社マツヤ 札幌支店'}],rows);assert.equal(nw.get(99999),'マツヤ');
console.log('OK');
