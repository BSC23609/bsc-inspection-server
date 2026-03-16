const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const XLSX = require('xlsx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID     = process.env.TENANT_ID;
const USER_ID       = process.env.USER_ID || 'pdqc@bharatsteels.in';

// ── Health check ─────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'BSC Inspection Server is running ✅' });
});

// ── Main submit endpoint ──────────────────────────────────────
app.post('/submit', async (req, res) => {
  try {
    const data = req.body;
    const folder = data.form_type; // 'Inward' or 'Quality'
    const ref    = data.ref;
    const pdf    = data.pdf_content;

    if (!folder || !ref) {
      return res.status(400).json({ status: 'error', message: 'Missing form_type or ref' });
    }

    const token = await getToken();

    // 1. Save PDF to OneDrive
    const pdfPath = `BSC Inspections/${folder}/PDF/${folder}_${ref}.html`;
    await uploadFile(token, pdfPath, pdf, 'text/html');

    // 2. Append row to Excel
    await appendExcelRow(token, folder, data, ref);

    res.json({ status: 'success', ref });

  } catch (err) {
    console.error('Submit error:', err.message);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

// ── Get Microsoft token (app-level, no user login needed) ─────
async function getToken() {
  const url  = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default'
  });
  const resp = await fetch(url, { method: 'POST', body });
  const json = await resp.json();
  if (!json.access_token) throw new Error('Token error: ' + JSON.stringify(json));
  return json.access_token;
}

// ── Upload any file to OneDrive ───────────────────────────────
async function uploadFile(token, filePath, content, contentType) {
  const url  = `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root:/${encodeURIComponent(filePath)}:/content`;
  const resp = await fetch(url, {
    method:  'PUT',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': contentType },
    body:    content
  });
  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Upload failed (${resp.status}): ${err}`);
  }
}

// ── Append row to Excel table ─────────────────────────────────
async function appendExcelRow(token, folder, data, ref) {
  const filePath  = `BSC Inspections/${folder}/${folder}_Log.xlsx`;
  const tableName = folder === 'Inward' ? 'InwardLog' : 'QualityLog';

  // Get file ID
  const fileUrl  = `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/root:/${encodeURIComponent(filePath)}`;
  const fileResp = await fetch(fileUrl, { headers: { 'Authorization': `Bearer ${token}` } });
  if (!fileResp.ok) throw new Error(`Excel file not found: ${filePath}`);
  const fileJson = await fileResp.json();
  const fileId   = fileJson.id;

  // Build row
  let values;
  if (folder === 'Inward') {
    values = [[
      ref,
      data.timestamp        || '',
      data.vehicle_number   || '',
      data.batch_number     || '',
      data.make_of_coil     || '',
      data.grade            || '',
      data.width            || '',
      data.thickness        || '',
      data.coil_weight      || '',
      data.coil_id          || '',
      data.actual_thickness || '',
      data.actual_width     || '',
      data.id_sticker       || '',
      data.edge_inner       || '',
      data.edge_outer       || '',
      data.scratch          || '',
      data.strapping        || '',
      data.rust             || '',
      data.other_damages    || '',
      data.inspected_by     || '',
      data.remarks          || ''
    ]];
  } else {
    const pq = data.processed_qty || {};
    const s  = (n) => (pq['size_'+n] || {});
    values = [[
      ref,
      data.timestamp      || '',
      data.date           || '',
      data.time           || '',
      data.coil_number    || '',
      data.batch_number   || '',
      data.make           || '',
      data.coil_thickness || '',
      data.coil_grade     || '',
      data.coil_width     || '',
      data.coil_weight    || '',
      data.first_bit      || '',
      data.last_bit       || '',
      data.defective      || '',
      data.balance_wt     || '',
      data.coil_verified  || '',
      data.blade_clearance|| '',
      data.operator       || '',
      data.inspector      || '',
      data.remarks        || '',
      s(1).length||'', s(1).nos||'', s(1).weight_t||'',
      s(2).length||'', s(2).nos||'', s(2).weight_t||'',
      s(3).length||'', s(3).nos||'', s(3).weight_t||'',
      s(4).length||'', s(4).nos||'', s(4).weight_t||'',
      s(5).length||'', s(5).nos||'', s(5).weight_t||'',
      s(6).length||'', s(6).nos||'', s(6).weight_t||'',
      s(7).length||'', s(7).nos||'', s(7).weight_t||'',
      s(8).length||'', s(8).nos||'', s(8).weight_t||'',
      s(9).length||'', s(9).nos||'', s(9).weight_t||'',
      s(10).length||'', s(10).nos||'', s(10).weight_t||''
    ]];
  }

  // POST row to Excel table
  const rowUrl  = `https://graph.microsoft.com/v1.0/users/${USER_ID}/drive/items/${fileId}/workbook/tables/${tableName}/rows/add`;
  const rowResp = await fetch(rowUrl, {
    method:  'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify({ values })
  });
  if (!rowResp.ok) {
    const err = await rowResp.json();
    throw new Error(`Excel row failed: ${JSON.stringify(err)}`);
  }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`BSC Server running on port ${PORT}`));
