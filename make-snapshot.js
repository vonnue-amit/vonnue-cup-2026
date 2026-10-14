// Writes data/snapshot.json from the two workbooks, for hosts that can't serve .xlsx files.
// Usage: npm i xlsx@0.18.5 && node make-snapshot.js
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, 'data');
const { readWorkbook } = require('./parser.js');
const slim = (file) => {
  // readWorkbook fills in formula results the file was saved without
  const wb = readWorkbook(XLSX, fs.readFileSync(path.join(dir, file)));
  const Sheets = {};
  wb.SheetNames.forEach((n) => {
    const ws = wb.Sheets[n]; const out = { '!ref': ws['!ref'] };
    Object.keys(ws).forEach((k) => { if (k[0] !== '!') { const c = ws[k]; out[k] = { t: c.t, v: c.v }; if (c.f) out[k].f = c.f; } });
    Sheets[n] = out;
  });
  return { SheetNames: wb.SheetNames, Sheets };
};
const savedAt = fs.statSync(path.join(dir, 'Vonnue_Cup_Knockout_Brackets.xlsx')).mtime.toISOString();
fs.writeFileSync(path.join(dir, 'snapshot.json'), JSON.stringify({
  savedAt, brackets: slim('Vonnue_Cup_Knockout_Brackets.xlsx'), participants: slim('Vonnue_Cup_Participants.xlsx'),
}));
console.log('Wrote data/snapshot.json');
