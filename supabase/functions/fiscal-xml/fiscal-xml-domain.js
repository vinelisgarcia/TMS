/* Deterministic XML fiscal correction. No network access or SAP assumptions. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TMSFiscalXML = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const MAX_BYTES = 2 * 1024 * 1024;
  const DSIG = 'http://www.w3.org/2000/09/xmldsig#';
  const fail = message => { throw new Error(message); };
  function parse(xml) {
    if (typeof xml !== 'string' || new TextEncoder().encode(xml).length > MAX_BYTES) fail('El XML supera el límite de 2 MB.');
    const declaration = xml.match(/^\uFEFF?\s*<\?xml\s+[^?]*\?>/i);
    const encoding = declaration && declaration[0].match(/\bencoding\s*=\s*['"]([^'"]+)['"]/i);
    if (encoding && !/^utf-8$/i.test(encoding[1])) fail('Solo se admiten XML codificados en UTF-8.');
    if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)) fail('No se permiten DTD ni declaraciones de entidades.');
    let document;
    try {
      // The browser ignores this option; the server parser must fail closed on recovery.
      document = new DOMParser({onError: (_level, message) => { throw new Error(message); }}).parseFromString(xml, 'application/xml');
    } catch (_) { fail('El XML no está correctamente formado.'); }
    if (!document.documentElement || document.getElementsByTagName('parsererror').length) fail('El XML no está correctamente formado.');
    if (document.getElementsByTagNameNS(DSIG, 'Signature').length || Array.from(document.getElementsByTagName('*')).some(e => /^(?:signature|digitalsignature)$/i.test(e.localName))) fail('El XML contiene una firma digital; se requiere un procedimiento de firma autorizado.');
    return document;
  }
  function validNcf(value) {
    const modern=value.match(/^(?:B(01|02|03|04|11|12|13|14|15|16|17)(\d{8})|E(31|32|33|34|41|42|43|44|45|46|47)(\d{10}))$/);
    if (modern) return /[1-9]/.test(modern[2] || modern[4]);
    return /^[AP]\d{18}$/.test(value) && /[1-9]/.test(value.slice(-8));
  }
  function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const [y,m,d] = value.split('-').map(Number);
    if (y < 1900 || m < 1 || m > 12 || d < 1) return false;
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }
  function path(element) {
    const result = [];
    for (let e = element; e && e.nodeType === 1; e = e.parentNode) result.unshift({namespace: e.namespaceURI || '', localName: e.localName});
    return result;
  }
  function samePath(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
  function candidates(doc, field, optional = false) {
    const prefix = field === 'ncf' ? /^DONCFInvoiceRef(?:_.+)?$/ : field === 'date' ? /^DONCFInvoiceRefDate(?:_.+)?$/ : /^ReferenceBusinessTransactionDocumentID$/;
    const matches = Array.from(doc.getElementsByTagName('*')).filter(e => prefix.test(e.localName));
    if (matches.length > 1) fail('Hay varios campos posibles para ' + field + '; no se puede seleccionar uno con seguridad.');
    if (!matches.length && !optional) fail('Falta el campo ' + field + '; no se puede inventar una estructura SAP.');
    return matches[0] || null;
  }
  function find(doc, field, config, optional = false) {
    const node = candidates(doc, field, optional);
    const entry = config && config.fields && config.fields[field];
    if (!node && optional && entry?.insert) return null;
    if (!node) fail('Falta el campo ' + field + '; no existe una plantilla de inserción validada.');
    if (!entry || !samePath(path(node), entry.path)) fail('El campo ' + field + ' no coincide con una plantilla validada para este tipo.');
    if (Array.from(node.childNodes).some(n => n.nodeType !== 3 && n.nodeType !== 4)) fail('El campo ' + field + ' tiene una estructura no admitida.');
    return node;
  }
  function template(originalXml, correctedXml, type) {
    if (!['credit','debit'].includes(type)) fail('Selecciona el tipo de documento.');
    const original = parse(originalXml), corrected = parse(correctedXml);
    const fields = {}, added = [];
    for (const field of ['ncf','date']) {
      const before=candidates(original,field,true), after=candidates(corrected,field);
      fields[field] = {path:path(after)};
      if (before && !samePath(path(before),path(after))) fail('La ubicación de un campo cambió entre las plantillas.');
      if (!before) {
        const parentPath=path(after.parentNode), parents=Array.from(original.getElementsByTagName('*')).filter(e=>samePath(path(e),parentPath));
        if (parents.length!==1) fail('No se puede identificar la ubicación única para insertar el campo fiscal.');
        const parent=parents[0], siblings=Array.from(parent.childNodes), correctedSiblings=Array.from(after.parentNode.childNodes);
        const index=correctedSiblings.indexOf(after), signature=siblings.map(n=>[n.nodeType,n.nodeName,n.namespaceURI||'']);
        const blank=after.cloneNode(true);blank.textContent='';
        fields[field].insert={parentPath,index,signature,xml:new XMLSerializer().serializeToString(blank)};
        added.push(fields[field].path);
      }
    }
    const references = Array.from(original.getElementsByTagName('*')).filter(e => e.localName === 'ReferenceBusinessTransactionDocumentID');
    if (references.length > 1) fail('Hay varias referencias SAP posibles.');
    if (references.length === 1) fields.reference = {path:path(references[0])};
    const config = {type, validated:true, fields};
    const affected = find(corrected,'ncf',config), date = find(corrected,'date',config);
    if (!validNcf(affected.textContent.trim()) || !validDate(date.textContent.trim())) fail('La plantilla corregida necesita un NCF y una fecha válidos.');
    if (!compareNodes(original,corrected,[fields.ncf.path,fields.date.path],added)) fail('El par de plantillas modifica campos no autorizados.');
    // Exercise the complete correction against the validated original before
    // this template can be activated, including any permitted insertion.
    const checked=correct(originalXml,config,{ncf:affected.textContent,date:date.textContent});
    if (checked.after.ncf!==affected.textContent.trim().toUpperCase() || checked.after.date!==date.textContent.trim()) fail('No se pudo reproducir la corrección de referencia.');
    return config;
  }
  function compareNodes(a, b, authorized, inserted = []) {
    if (!a || !b || a.nodeType !== b.nodeType || a.nodeName !== b.nodeName || a.namespaceURI !== b.namespaceURI) return false;
    if (a.nodeType === 1) {
      const attrs = e => Array.from(e.attributes).map(x => [x.namespaceURI || '', x.name, x.value]).sort((x,y) => JSON.stringify(x).localeCompare(JSON.stringify(y)));
      if (JSON.stringify(attrs(a)) !== JSON.stringify(attrs(b))) return false;
      if (authorized.some(p => samePath(path(a), p))) return Array.from(b.childNodes).every(n => n.nodeType === 3 || n.nodeType === 4);
    } else if (a.nodeValue !== b.nodeValue) return false;
    const childrenA=Array.from(a.childNodes), childrenB=Array.from(b.childNodes).filter(n=>n.nodeType!==1||!inserted.some(p=>samePath(path(n),p)));
    if (childrenA.length !== childrenB.length) return false;
    for (let i = 0; i < childrenA.length; i++) if (!compareNodes(childrenA[i], childrenB[i], authorized, inserted)) return false;
    return true;
  }
  function information(doc) {
    const nodes = Array.from(doc.getElementsByTagName('*'));
    const one = names => { const matches = nodes.filter(e => names.some(name => name.test(e.localName))); return matches.length === 1 ? matches[0].textContent.trim() : ''; };
    const direct = names => { const matches = Array.from(doc.documentElement.childNodes).filter(e => e.nodeType === 1 && names.includes(e.localName)); return matches.length === 1 ? matches[0].textContent.trim() : ''; };
    const ownNcf = one([/^DONCF(?:_.+)?$/]);
    const detectedType = /^(?:E34|B04)/.test(ownNcf) ? 'credit' : /^(?:E33|B03)/.test(ownNcf) ? 'debit' : '';
    return {number:direct(['ID','DocumentID']),ownNcf,issueDate:direct(['IssueDate','Date']),client:one([/^BuyerPartyName$/,/^CustomerName$/]),currency:one([/^CurrencyCode$/]),total:one([/^TotalGrossAmount$/,/^GrossAmount$/]),tax:one([/^TotalTaxAmount$/]),detectedType,namespaces:[...new Set(nodes.map(e => e.namespaceURI || '').filter(Boolean))],extensions:[...new Set(nodes.filter(e=>/\/CustomerExtension\//.test(e.namespaceURI||'')).map(e=>e.localName))].slice(0,100)};
  }
  function analyze(xml, config) {
    const doc = parse(xml);
    if (!config || !['credit','debit'].includes(config.type) || config.validated !== true) fail('No hay una plantilla validada para este tipo de documento.');
    const ncf = find(doc, 'ncf', config, true), date = find(doc, 'date', config, true);
    let reference = null;
    if (config.fields.reference) reference = find(doc, 'reference', config);
    const info = information(doc);
    if (info.detectedType && info.detectedType !== config.type) fail('El tipo identificado no coincide con la plantilla seleccionada.');
    return { doc, information:info, fields: {ncf: ncf ? ncf.textContent : '', date: date ? date.textContent : '', reference: reference ? reference.textContent : ''}, paths: {ncf: config.fields.ncf.path, date: config.fields.date.path, reference: reference ? path(reference) : null} };
  }
  function insertField(doc, config, field, inserted, offset) {
    const spec=config.fields[field].insert;
    if (!spec) fail('Falta el campo '+field+' y no existe una plantilla validada para insertarlo.');
    const parents=Array.from(doc.getElementsByTagName('*')).filter(e=>samePath(path(e),spec.parentPath));
    if(parents.length!==1) fail('La ubicación de inserción no es inequívoca.');
    const parent=parents[0], children=Array.from(parent.childNodes), signature=children.filter(n=>n.nodeType!==1||!inserted.some(p=>samePath(path(n),p))).map(n=>[n.nodeType,n.nodeName,n.namespaceURI||'']);
    if(JSON.stringify(signature)!==JSON.stringify(spec.signature)) fail('La ubicación no coincide con la estructura de la plantilla validada.');
    const element=parse(spec.xml).documentElement;
    parent.insertBefore(doc.importNode(element,true),children[spec.index-offset]||null);
    return find(doc,field,config);
  }
  function correct(xml, config, input) {
    const original = analyze(xml, config);
    const ncf = String(input.ncf || '').trim().toUpperCase(), date = String(input.date || '').trim();
    if (!validNcf(ncf)) fail('El NCF o e-NCF no tiene una estructura válida.');
    if (!validDate(date)) fail('La fecha de la factura afectada no es válida.');
    const doc = parse(xml), fields = {ncf: find(doc,'ncf',config,true), date: find(doc,'date',config,true)}, inserted=[];
    const missing=['ncf','date'].filter(field=>!fields[field]);
    missing.sort((a,b)=>(config.fields[b].insert?.index??-1)-(config.fields[a].insert?.index??-1));
    for(const field of missing) {
      const spec=config.fields[field].insert;
      const offset=spec ? missing.filter(other=>other!==field && config.fields[other].insert && samePath(config.fields[other].insert.parentPath,spec.parentPath) && config.fields[other].insert.index<spec.index).length : 0;
      fields[field]=insertField(doc,config,field,inserted,offset);inserted.push(original.paths[field]);
    }
    const authorized = [original.paths.ncf, original.paths.date];
    fields.ncf.textContent = ncf; fields.date.textContent = date;
    let reference = original.fields.reference;
    if (input.changeReference === true) {
      if (!original.paths.reference) fail('La plantilla no autoriza cambiar la referencia SAP.');
      reference = String(input.reference || '').trim();
      if (!reference || reference.length > 100 || /[\x00-\x1f]/.test(reference)) fail('La referencia SAP indicada no es válida.');
      find(doc, 'reference', config).textContent = reference;
      authorized.push(original.paths.reference);
    }
    const serialized = (xml.startsWith('\uFEFF') ? '\uFEFF' : '') + new XMLSerializer().serializeToString(doc);
    const result = parse(serialized);
    if (!compareNodes(original.doc, result, authorized, inserted)) fail('La comparación detectó cambios no autorizados en el XML.');
    return {xml: serialized, before: original.fields, after: {ncf,date,reference}, information:original.information, warnings: ['La validación estructural del TMS no sustituye la aceptación de la DGII.','No se ha validado un esquema XSD.']};
  }
  return Object.freeze({ MAX_BYTES, parse, validNcf, validDate, template, information, analyze, correct, compareNodes, path });
});
