/* Prueba de integración del núcleo OOXML. Requiere Node y @xmldom/xmldom. */
'use strict';
const fs=require('fs'),path=require('path');
const root=path.resolve(__dirname,'..');
let xmldom;
try{xmldom=require('@xmldom/xmldom')}catch(error){
  const adobe='C:/Program Files/Adobe/Adobe Creative Cloud Experience/js/node_modules/@xmldom/xmldom';
  if(!fs.existsSync(adobe))throw error;
  xmldom=require(adobe);
}
global.DOMParser=xmldom.DOMParser;
global.XMLSerializer=xmldom.XMLSerializer;
const JSZip=require(path.join(root,'assets/vendor/jszip.min.js'));
const core=require(path.join(root,'assets/js/app.js'));
const section={numero:'02',titulo:'Sección de prueba',meta:'EVIDENCIA',resumen:'Resumen no vacío.',idea:'Idea clave no vacía.',texto:'Desarrollo no vacío.'};

async function main(){
  let tested=0;
  for(const doc of core.DOCS.filter(item=>!item.static)){
    const templatePaths=doc.templateVariants?Object.values(doc.templateVariants):[doc.file];
    for(const templatePath of templatePaths){
    const source=fs.readFileSync(path.join(root,templatePath));
    const zip=await JSZip.loadAsync(source),names=Object.keys(zip.files).filter(name=>!zip.files[name].dir).sort(),part=zip.file('word/document.xml');
    if(!part)throw new Error(`${doc.id}: falta document.xml`);
    const sourceXml=await part.async('string'),sourceAnchors=(sourceXml.match(/wp:anchor/g)||[]).length;
    const xml=new DOMParser().parseFromString(sourceXml,'application/xml');
    if(doc.rows)core.expandDynamicRows(xml,doc.rows.defaults);
    if(doc.sections)core.expandSections(xml,[section,{...section,numero:'03',titulo:'Segunda sección de prueba'}]);
    const values={};
    for(const name of doc.fields)values[name]=core.FIELDS[name][1];
    values.TEMA=values.tema||core.FIELDS.tema[1];
    values.ciudad_upper=(values.ciudad||core.FIELDS.ciudad[1]).toLocaleUpperCase('es');
    values.titulo_upper=(values.titulo||core.FIELDS.titulo[1]).toLocaleUpperCase('es');
    values.descripcion_actividad_upper=(values.descripcion_actividad||core.FIELDS.descripcion_actividad[1]).toLocaleUpperCase('es');
    const allowedGenerated=new Set(['dynamic_rows','resumen_seccion','idea_clave_seccion','texto_seccion']);
    const templateKeys=[...sourceXml.matchAll(/\{\{\s*([^}]+?)\s*\}\}/g)].map(match=>match[1]);
    const missingFields=[...new Set(templateKeys.filter(key=>!(key in values)&&!allowedGenerated.has(key)))];
    if(missingFields.length)throw new Error(`${doc.id}: el formulario no expone ${missingFields.join(', ')}`);
    core.replaceTokens(xml,values);
    const serialized=new XMLSerializer().serializeToString(xml);
    if(serialized.includes('{{'))throw new Error(`${doc.id}: quedaron marcadores`);
    if(!serialized.trim())throw new Error(`${doc.id}: salida vacía`);
    if((serialized.match(/wp:anchor/g)||[]).length<sourceAnchors)throw new Error(`${doc.id}: perdió objetos anclados del maestro`);
    if(doc.rows&&!serialized.includes(doc.rows.defaults[0][1]))throw new Error(`${doc.id}: no insertó filas configuradas`);
    if(doc.sections&&!serialized.includes('Segunda sección de prueba'))throw new Error(`${doc.id}: no replicó secciones maestras`);
    zip.file('word/document.xml',serialized);
    const output=await zip.generateAsync({type:'nodebuffer'}),reopened=await JSZip.loadAsync(output),outputNames=Object.keys(reopened.files).filter(name=>!reopened.files[name].dir).sort();
    if(JSON.stringify(names)!==JSON.stringify(outputNames)){
      const added=outputNames.filter(name=>!names.includes(name)),removed=names.filter(name=>!outputNames.includes(name));
      throw new Error(`${doc.id}: cambió la topología OOXML; agregadas=${added.join(',')}; eliminadas=${removed.join(',')}`);
    }
    if(doc.rows){
      const outputXml=await reopened.file('word/document.xml').async('string');
      if(outputXml.includes('{{dynamic_rows}}'))throw new Error(`${doc.id}: fila dinámica no resuelta`);
    }
    tested++;
    process.stdout.write(`GEN_OK ${doc.id} ${path.basename(templatePath)}\n`);
    }
  }
  process.stdout.write(`OK: ${tested} salidas DOCX generadas en memoria desde sus plantillas.\n`);
}
main().catch(error=>{process.stderr.write(`FALLO: ${error.stack||error.message}\n`);process.exitCode=1});
