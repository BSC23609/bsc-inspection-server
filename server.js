const express = require('express');
const cors    = require('cors');
const fetch   = require('node-fetch');
const htmlPdf = require('html-pdf-node');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID     = process.env.TENANT_ID;
const USER_ID       = process.env.USER_ID || 'pdqc@bharatsteels.in';

app.get('/', (req, res) => {
  res.json({ status: 'BSC Inspection Server is running' });
});

app.post('/submit', async (req, res) => {
  try {
    const data     = req.body;
    const folder   = data.form_type;
    const batchNo  = (data.batch_number || 'NOBATCH').replace(/[^a-zA-Z0-9\-_]/g, '_');
    const dateStr  = new Date().toLocaleDateString('en-IN', { day:'2-digit', month:'2-digit', year:'numeric' }).replace(/\//g, '-');
    const suffix   = folder === 'Inward' ? 'Inward' : 'CTL_Inspection';
    const fileName = batchNo + '_(' + dateStr + ')_' + suffix;
    if (!folder) return res.status(400).json({ status: 'error', message: 'Missing form_type' });
    const token = await getToken();
    const pdfBuffer = await generatePDF(data.pdf_content);
    await uploadFile(token, 'BSC Inspections/' + folder + '/PDF/' + fileName + '.pdf', pdfBuffer, 'application/pdf');
    await appendExcelRow(token, folder, data, fileName);
    res.json({ status: 'success', ref: data.ref, filename: fileName });
  } catch (err) {
    console.error('Submit error:', err.message);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

app.post('/download', async (req, res) => {
  try {
    const { folder, batch, date_str } = req.body;
    if (!folder) return res.status(400).json({ status: 'error', message: 'Missing folder' });
    if (!batch && !date_str) return res.status(400).json({ status: 'error', message: 'Provide batch or date' });
    const token      = await getToken();
    const folderPath = 'BSC Inspections/' + folder + '/PDF';
    const listUrl    = 'https://graph.microsoft.com/v1.0/users/' + USER_ID + '/drive/root:/' + encodeURIComponent(folderPath) + ':/children?$select=name,id&$top=200';
    const listResp   = await fetch(listUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    if (!listResp.ok) return res.status(404).json({ status: 'error', message: 'PDF folder not found on OneDrive' });
    const files   = ((await listResp.json()).value || []).map(f => f.name);
    const matches = files.filter(name => {
      const matchBatch = batch    ? name.toLowerCase().includes(batch.toLowerCase()) : true;
      const matchDate  = date_str ? name.includes(date_str) : true;
      return matchBatch && matchDate && name.endsWith('.pdf');
    });
    if (matches.length === 0) return res.status(404).json({ status: 'error', message: 'No report found matching your search' });
    const fileName = matches.sort().pop();
    const fileUrl  = 'https://graph.microsoft.com/v1.0/users/' + USER_ID + '/drive/root:/' + encodeURIComponent(folderPath + '/' + fileName) + ':/content';
    const fileResp = await fetch(fileUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    if (!fileResp.ok) return res.status(404).json({ status: 'error', message: 'File not found' });
    const buffer = await fileResp.buffer();
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fileName + '"');
    res.send(buffer);
  } catch (err) {
    console.error('Download error:', err.message);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

async function generatePDF(htmlContent) {
  return await htmlPdf.generatePdf({ content: htmlContent }, { format: 'A4', printBackground: true, margin: { top: '15mm', bottom: '15mm', left: '15mm', right: '15mm' } });
}

async function getToken() {
  const body = new URLSearchParams({ grant_type: 'client_credentials', client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default' });
  const resp = await fetch('https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token', { method: 'POST', body });
  const json = await resp.json();
  if (!json.access_token) throw new Error('Token error: ' + JSON.stringify(json));
  return json.access_token;
}

async function uploadFile(token, filePath, content, contentType) {
  const resp = await fetch('https://graph.microsoft.com/v1.0/users/' + USER_ID + '/drive/root:/' + encodeURIComponent(filePath) + ':/content', {
    method: 'PUT', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': contentType }, body: content
  });
  if (!resp.ok) throw new Error('Upload failed (' + resp.status + '): ' + await resp.text());
}

async function appendExcelRow(token, folder, data, fileName) {
  const filePath  = 'BSC Inspections/' + folder + '/' + folder + '_Log.xlsx';
  const tableName = folder === 'Inward' ? 'InwardLog' : 'QualityLog';
  const fileResp  = await fetch('https://graph.microsoft.com/v1.0/users/' + USER_ID + '/drive/root:/' + encodeURIComponent(filePath), { headers: { 'Authorization': 'Bearer ' + token } });
  if (!fileResp.ok) throw new Error('Excel file not found: ' + filePath);
  const fileId = (await fileResp.json()).id;
  const pq = data.processed_qty || {};
  const s  = n => (pq['size_' + n] || {});
  const values = folder === 'Inward' ? [[
    fileName, data.timestamp||'', data.vehicle_number||'', data.batch_number||'', data.make_of_coil||'', data.grade||'',
    data.width||'', data.thickness||'', data.coil_weight||'', data.coil_id||'', data.actual_thickness||'', data.actual_width||'',
    data.id_sticker||'', data.edge_inner||'', data.edge_outer||'', data.scratch||'', data.strapping||'', data.rust||'',
    data.other_damages||'', data.inspected_by||'', data.remarks||''
  ]] : [[
    fileName, data.timestamp||'', data.date||'', data.time||'', data.coil_number||'', data.batch_number||'',
    data.make||'', data.coil_thickness||'', data.coil_grade||'', data.coil_width||'', data.coil_weight||'',
    data.first_bit||'', data.last_bit||'', data.defective||'', data.balance_wt||'', data.coil_verified||'',
    data.blade_clearance||'', data.operator||'', data.inspector||'', data.remarks||'',
    s(1).length||'', s(1).nos||'', s(1).weight_t||'', s(2).length||'', s(2).nos||'', s(2).weight_t||'',
    s(3).length||'', s(3).nos||'', s(3).weight_t||'', s(4).length||'', s(4).nos||'', s(4).weight_t||'',
    s(5).length||'', s(5).nos||'', s(5).weight_t||'', s(6).length||'', s(6).nos||'', s(6).weight_t||'',
    s(7).length||'', s(7).nos||'', s(7).weight_t||'', s(8).length||'', s(8).nos||'', s(8).weight_t||'',
    s(9).length||'', s(9).nos||'', s(9).weight_t||'', s(10).length||'', s(10).nos||'', s(10).weight_t||''
  ]];
  const rowResp = await fetch('https://graph.microsoft.com/v1.0/users/' + USER_ID + '/drive/items/' + fileId + '/workbook/tables/' + tableName + '/rows/add', {
    method: 'POST', headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ values })
  });
  if (!rowResp.ok) throw new Error('Excel row failed: ' + JSON.stringify(await rowResp.json()));
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('BSC Server running on port ' + PORT));
