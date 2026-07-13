const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const excelPath = path.join(__dirname, 'bd-matriz-estructura.xlsx');

const wb = XLSX.readFile(excelPath);
const sheetName = wb.SheetNames[0];
const ws = wb.Sheets[sheetName];

// Convertir a matriz de filas
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

if (!rows.length) {
  console.error('❌ La hoja está vacía');
  process.exit(1);
}

// La primera fila es el encabezado
const headers = rows[0].map(h => String(h).trim());
console.log('📋 Encabezados detectados:');
headers.forEach((h, i) => console.log(`   [${i}] ${h}`));

// Detección flexible de columnas por nombre
function findCol(...patterns) {
  for (const p of patterns) {
    const idx = headers.findIndex(h => h.toLowerCase().includes(p.toLowerCase()));
    if (idx !== -1) return idx;
  }
  return -1;
}

const colId       = findCol('id', '#', 'no.');
const colRiesgo   = findCol('riesgo_desc', 'riesgo descrip', 'descripcion del riesgo', 'riesgo');
const colCausa    = findCol('causa_desc', 'causa descrip', 'descripcion de la causa', 'causa');
const colControl  = findCol('control_desc', 'control descrip', 'descripcion del control', 'control');
const colCalifR   = findCol('calif_riesgo', 'calificacion riesgo');
const colCalifC   = findCol('calif_causa', 'calificacion causa');
const colCalifCtrl= findCol('calif_control', 'calificacion control');

console.log('\n🔎 Mapeo de columnas:');
console.log(`   id=${colId}, riesgo=${colRiesgo}, causa=${colCausa}, control=${colControl}`);
console.log(`   califR=${colCalifR}, califC=${colCalifC}, califCtrl=${colCalifCtrl}`);

const data = [];
for (let i = 1; i < rows.length; i++) {
  const r = rows[i];
  // Saltar filas totalmente vacías
  const riesgo = colRiesgo >= 0 ? String(r[colRiesgo] || '').trim() : '';
  const causa = colCausa >= 0 ? String(r[colCausa] || '').trim() : '';
  const control = colControl >= 0 ? String(r[colControl] || '').trim() : '';

  if (!riesgo && !causa && !control) continue;

  const toNum = v => {
    const n = Number(v);
    return (!isNaN(n) && n >= 1 && n <= 3) ? n : null;
  };

  data.push({
    id: colId >= 0 && r[colId] !== '' ? r[colId] : data.length + 1,
    riesgo,
    causa,
    control,
    calif_riesgo: colCalifR >= 0 ? toNum(r[colCalifR]) : null,
    calif_causa: colCalifC >= 0 ? toNum(r[colCalifC]) : null,
    calif_control: colCalifCtrl >= 0 ? toNum(r[colCalifCtrl]) : null
  });
}

const outputPath = path.join(__dirname, 'data.json');
fs.writeFileSync(outputPath, JSON.stringify(data, null, 2), 'utf8');

console.log(`\n✓ Extraídos ${data.length} registros REALES a data.json`);
console.log(`  Primer registro (muestra):`);
if (data[0]) {
  console.log(`    ID: ${data[0].id}`);
  console.log(`    Riesgo: ${data[0].riesgo.substring(0, 80)}...`);
  console.log(`    Causa: ${data[0].causa.substring(0, 80)}...`);
  console.log(`    Control: ${data[0].control.substring(0, 80)}...`);
}
