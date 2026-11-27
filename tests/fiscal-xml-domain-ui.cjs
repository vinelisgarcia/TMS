const {chromium}=require(process.env.TMS_PLAYWRIGHT||'/Users/vinelis/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const assert=require('node:assert/strict');
(async()=>{const browser=await chromium.launch();try {const page=await browser.newPage();await page.setContent('<!doctype html><title>XML fiscal test</title>');await page.addScriptTag({path:'fiscal-xml-domain.js'});const results=await page.evaluate(()=>{
 const d=window.TMSFiscalXML, failures=[], tests=[];
 const xml=(type='credit',pref='n1',suffix='ONE',values=['',''])=>`<?xml version="1.0" encoding="UTF-8"?><doc xmlns:${pref}="urn:test" type="${type}"><id>123</id><${pref}:DONCFInvoiceRef_${suffix}>${values[0]}</${pref}:DONCFInvoiceRef_${suffix}><${pref}:DONCFInvoiceRefDate_OTHER>${values[1]}</${pref}:DONCFInvoiceRefDate_OTHER><ReferenceBusinessTransactionDocumentID>KEEP</ReferenceBusinessTransactionDocumentID><total>100.00</total><tax>18.00</tax></doc>`;
 function config(s,type='credit'){const doc=d.parse(s), els=[...doc.getElementsByTagName('*')],p=name=>d.path(els.find(e=>e.localName.startsWith(name)));return {type,validated:true,fields:{ncf:{path:p('DONCFInvoiceRef_')},date:{path:p('DONCFInvoiceRefDate_')},reference:{path:p('ReferenceBusinessTransactionDocumentID')}}};}
 function test(name,fn){try{fn();tests.push(name)}catch(e){failures.push(name+': '+e.message)}}
 const check=(x,msg)=>{if(!x)throw Error(msg||'assertion failed')},reject=fn=>{let failed=false;try{fn()}catch(e){failed=true}check(failed,'expected rejection')};
 const data={ncf:'E310000001994',date:'2026-02-20'};
 for(const type of ['credit','debit']) test('nota '+type,()=>{let s=xml(type),r=d.correct(s,config(s,type),data);check(r.after.ncf===data.ncf&&r.after.date===data.date);check(r.xml.includes('<total>100.00</total>')&&r.xml.includes('<tax>18.00</tax>'));check(r.after.reference==='KEEP')});
 test('campos completados',()=>{let s=xml('credit','a','OTHER',['B0100000001','2025-01-01']),r=d.correct(s,config(s),data);check(r.before.ncf==='B0100000001'&&r.before.date==='2025-01-01')});
 test('prefijos y sufijos',()=>{let s=xml('credit','other','DIFFERENT');check(d.correct(s,config(s),data).xml.includes('other:DONCFInvoiceRef_DIFFERENT'))});
 test('inserción con plantilla validada',()=>{const original='<doc xmlns:n="urn:test"><id>1</id><tail>safe</tail></doc>',corrected='<doc xmlns:n="urn:test"><id>1</id><n:DONCFInvoiceRef_X>E310000001994</n:DONCFInvoiceRef_X><n:DONCFInvoiceRefDate_Y>2026-01-01</n:DONCFInvoiceRefDate_Y><tail>safe</tail></doc>',c=d.template(original,corrected,'debit'),r=d.correct(original,c,data);check(r.xml.indexOf('DONCFInvoiceRef_X')<r.xml.indexOf('DONCFInvoiceRefDate_Y'));check(r.xml.includes('<tail>safe</tail>'));reject(()=>d.correct(original.replace('<tail>','<extra/><tail>'),c,data))});
 test('campos ausentes',()=>{let s=xml(),c=config(s);reject(()=>d.correct(s.replace(/<n1:DONCFInvoiceRef_ONE>.*?<\/n1:DONCFInvoiceRef_ONE>/,''),c,data))});
 test('duplicados',()=>{let s=xml(),c=config(s);s=s.replace('</doc>','<n1:DONCFInvoiceRef_TWO>xx</n1:DONCFInvoiceRef_TWO></doc>');reject(()=>d.correct(s,c,data))});
 test('XML mal formado',()=>reject(()=>d.parse('<doc>')));
 test('fecha inválida',()=>{let s=xml();reject(()=>d.correct(s,config(s),{...data,date:'2026-02-30'}))});
 test('NCF inválido',()=>{let s=xml();reject(()=>d.correct(s,config(s),{...data,ncf:'E123'}))});
 test('firma digital',()=>reject(()=>d.parse('<doc xmlns:d="http://www.w3.org/2000/09/xmldsig#"><d:Signature/></doc>')));
 test('DTD y entidades',()=>reject(()=>d.parse('<!DOCTYPE doc [<!ENTITY x SYSTEM "file:///etc/passwd">]><doc>&x;</doc>')));
 test('referencia explícita',()=>{let s=xml(),r=d.correct(s,config(s),{...data,changeReference:true,reference:'SAP123'});check(r.after.reference==='SAP123')});
 test('cambio no autorizado',()=>{let s=xml(),a=d.parse(s),b=d.parse(s.replace('100.00','1.00'));check(!d.compareNodes(a,b,[]))});
 test('plantilla no validada',()=>{let s=xml(),c=config(s);c.validated=false;reject(()=>d.correct(s,c,data))});
 return {tests,failures};});assert.deepEqual(results.failures,[]);console.log(`PASS: ${results.tests.length} comprobaciones XML fiscal`);}finally{await browser.close()}})().catch(e=>{console.error(e);process.exitCode=1});
