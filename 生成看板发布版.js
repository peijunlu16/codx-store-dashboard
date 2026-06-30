const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

const root = __dirname;
const sourcePath = path.join(root, '门店运营看板_成品.html');
const outputPath = path.join(root, '门店运营看板_发布版_20260629.html');
let html = fs.readFileSync(sourcePath, 'utf8');

function extractBetween(source, startText, endText) {
  const start = source.indexOf(startText);
  const end = source.indexOf(endText, start);
  if (start < 0 || end < 0) throw new Error(`未找到代码片段: ${startText}`);
  return source.slice(start, end);
}

function cleanCode(v) {
  return v == null ? '' : String(v).trim().replace(/\.0$/, '');
}

function extractSizeCode(v) {
  const s = cleanCode(v).replace(/\D/g, '');
  return s.length >= 16 ? s.slice(13, 16) : '';
}

function readRows(fileName) {
  const buffer = fs.readFileSync(path.join(root, fileName));
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: true });
}

const salesFunction = extractBetween(
  html,
  'function rebuildFromRows(rows){',
  'function normalizeInventoryStore'
);
const inventoryFunctions = extractBetween(
  html,
  'function normalizeInventoryStore',
  'function workbookRowsFromArrayBuffer'
);
const builders = new Function(
  'cleanCode',
  'extractSizeCode',
  'alert',
  `${salesFunction}\n${inventoryFunctions}\nreturn {rebuildFromRows,rebuildInventoryFromRows};`
)(cleanCode, extractSizeCode, message => { throw new Error(message); });

const db = builders.rebuildFromRows(readRows('线下直营数据源.xlsx'));
Object.assign(db, builders.rebuildInventoryFromRows(readRows('门店库存.xlsx'), db));
const latestDaily = db.daily.filter(row => row.date === db.dateMax);
const latestTransactions = latestDaily.reduce((sum, row) => sum + (row.transactions || 0), 0);
const latestNetQty = latestDaily.reduce((sum, row) => sum + (row.qty || 0), 0);
const latestPaidQty = latestDaily.reduce((sum, row) => sum + (row.paidQty || 0), 0);

const embeddedStart = html.indexOf('const EMBEDDED_DATA=');
const stateStart = html.indexOf('let state=', embeddedStart);
const embeddedEnd = html.lastIndexOf(';', stateStart);
if (embeddedStart < 0 || embeddedEnd < 0) throw new Error('未找到内嵌数据区');
html = html.slice(0, embeddedStart) + `const EMBEDDED_DATA=${JSON.stringify(db)}` + html.slice(embeddedEnd);

html = html.replace(
  /<script>\s*if\(location\.protocol==='file:'\)\{[\s\S]*?<\/script>\s*/,
  ''
);
html = html.replace('<script src="xlsx.full.min.js"></script>\n', '');
html = html.replace(
  '</style>\n</head>',
  '.upload-btn,#xlUpload,#invUpload{display:none!important}\n</style>\n</head>'
);
html = html.replace(
  "document.getElementById('data-status').textContent='自动更新需通过「启动看板自动更新.ps1」打开';",
  `document.getElementById('data-status').textContent='发布版 · 销售截至 ${db.dateMax} · 库存 ${db.invDate}';`
);
html = html.replace('<title>门店运营看板</title>', '<title>门店运营看板｜发布版</title>');

const inlineScripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
  .map(match => match[1])
  .filter(Boolean);
inlineScripts.forEach(script => new Function(script));
if (html.includes('location.replace')) throw new Error('发布版仍包含本地服务跳转');
if (!html.includes(`"dateMax":"${db.dateMax}"`)) throw new Error('销售日期写入失败');
if (!html.includes(`"invDate":"${db.invDate}"`)) throw new Error('库存日期写入失败');

fs.writeFileSync(outputPath, html, 'utf8');
console.log(JSON.stringify({
  output: path.basename(outputPath),
  dateMin: db.dateMin,
  dateMax: db.dateMax,
  invDate: db.invDate,
  stores: db.stores.length,
  daily: db.daily.length,
  latestTransactions,
  latestNetQty,
  latestPaidQty,
  latestUpt: latestTransactions ? Number((latestPaidQty / latestTransactions).toFixed(1)) : 0,
  scriptsChecked: inlineScripts.length,
  bytes: fs.statSync(outputPath).size
}, null, 2));
