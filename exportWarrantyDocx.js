// Uses global JSZip and saveToDiskUtility from index.html/fileSystem.js

const WARRANTY_TEMPLATE_FILE = './Гарантійний талон монтаж.docx';

const warrantyXmlEscape = (value = '') => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&apos;');

const warrantyToNumber = (value, fallback = 0) => {
  if (value === null || value === undefined) return fallback;
  if (typeof value === 'string') {
    const n = Number(value.replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : fallback;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
};

const warrantyFormatDate = (value) => {
  const raw = String(value || '').trim();
  if (!raw) return '___ . ___ . _______';
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(raw)) return raw;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? raw : parsed.toLocaleDateString('uk-UA');
};

const warrantyFormatQty = (qty, unit = 'шт') => {
  const value = warrantyToNumber(qty, 0);
  const cleanUnit = String(unit || 'шт').trim() || 'шт';
  const text = Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100).replace('.', ',');
  return `${text} ${cleanUnit}`;
};

const warrantySplitTables = (xml) => xml.match(/<w:tbl[\s\S]*?<\/w:tbl>/g) || [];
const warrantySplitRows = (tableXml) => tableXml.match(/<w:tr[\s\S]*?<\/w:tr>/g) || [];
const warrantySplitCells = (rowXml) => rowXml.match(/<w:tc[\s\S]*?<\/w:tc>/g) || [];

const warrantyTextFromXml = (xml = '') => (xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [])
  .map((part) => part.replace(/^<w:t[^>]*>/, '').replace(/<\/w:t>$/, ''))
  .join('');

const warrantySetCellText = (cellXml, value) => {
  const tcPr = (cellXml.match(/<w:tcPr[\s\S]*?<\/w:tcPr>/) || [''])[0];
  const pPr = (cellXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/) || [''])[0];
  const rPr = (cellXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/) || [''])[0];
  const text = warrantyXmlEscape(value);
  return `<w:tc>${tcPr}<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
};

const warrantySetRowCells = (rowXml, values) => {
  const cells = warrantySplitCells(rowXml);
  if (!cells.length) return rowXml;
  let nextRow = rowXml;
  values.forEach((value, idx) => {
    if (!cells[idx]) return;
    const nextCell = warrantySetCellText(cells[idx], value);
    nextRow = nextRow.replace(cells[idx], nextCell);
  });
  return nextRow;
};

const warrantyFindTable = (xml, predicate) => warrantySplitTables(xml).find(predicate) || '';

const warrantyGetItemType = (item = {}, groupKey = '') => {
  const type = String(item?.type || '').trim();
  if (type) return type;
  return String(groupKey || '').trim();
};

const warrantyClassifyItem = (item = {}, groupKey = '') => {
  const type = warrantyGetItemType(item, groupKey).toLowerCase();
  const name = String(item?.name || '').toLowerCase();
  const group = String(groupKey || '').toLowerCase();
  if (type.includes('феп') || name.includes('solar') || name.includes('панел') || name.includes('модул')) {
    return { equipment: 'Сонячна панель', warranty: '10 років', priority: 1 };
  }
  if (type.includes('інвертор') || name.includes('inverter') || name.includes('інвертор')) {
    return { equipment: 'Інвертор', warranty: '5 років', priority: 2 };
  }
  if (type.includes('акб') || name.includes('акб') || name.includes('battery') || name.includes('lifepo') || name.includes('батар')) {
    return { equipment: 'Акумуляторна батарея', warranty: '3 роки', priority: 3 };
  }
  if (type.includes('bms') || name.includes('bms')) {
    return { equipment: 'BMS / контролер АКБ', warranty: '3 роки', priority: 4 };
  }
  if (type.includes('mppt') || name.includes('mppt') || name.includes('контролер')) {
    return { equipment: 'Контролер', warranty: '3 роки', priority: 5 };
  }
  if (group.includes('кріплення') || name.includes('кріплен') || name.includes('профіль')) {
    return { equipment: 'Монтажний комплект', warranty: '1 рік', priority: 20 };
  }
  return null;
};

const buildWarrantyRows = (calculations = {}) => {
  const rows = [];
  const groups = calculations?.groups && typeof calculations.groups === 'object' ? calculations.groups : {};
  Object.keys(groups).forEach((groupKey) => {
    (Array.isArray(groups[groupKey]) ? groups[groupKey] : []).forEach((item) => {
      const qty = warrantyToNumber(item?.quantity, 0);
      const sumUsd = warrantyToNumber(item?.sumUsd, 0);
      const sumUah = warrantyToNumber(item?.sumUah, 0);
      if (!String(item?.name || '').trim() || qty <= 0 || (sumUsd <= 0 && sumUah <= 0)) return;
      const classified = warrantyClassifyItem(item, groupKey);
      if (!classified) return;
      rows.push({
        ...classified,
        model: String(item.name || '').trim(),
        serial: '_______________',
        qty: warrantyFormatQty(qty, item.unit || 'шт')
      });
    });
  });

  if (warrantyToNumber(calculations?.sums?.installPercentAmountUsd, 0) > 0 || rows.some((row) => row.equipment === 'Монтажний комплект')) {
    const hasMounting = rows.some((row) => row.equipment === 'Монтажний комплект');
    if (!hasMounting) {
      rows.push({
        equipment: 'Монтажний комплект',
        model: 'Профіль + кріплення',
        serial: '—',
        qty: '1 к-т',
        warranty: '1 рік',
        priority: 20
      });
    }
  }

  if (!rows.length) {
    rows.push({
      equipment: 'Обладнання',
      model: '________________',
      serial: '_______________',
      qty: '___',
      warranty: '___'
    });
  }

  return rows
    .sort((a, b) => (a.priority || 99) - (b.priority || 99))
    .slice(0, 10);
};

const warrantyReplaceInfoTable = (documentXml, clientInfo = {}, documentDetails = {}) => {
  const table = warrantyFindTable(documentXml, (tbl) => tbl.includes('Замовник:') && tbl.includes("Адреса об'єкту:"));
  if (!table) return documentXml;
  const rows = warrantySplitRows(table);
  const values = [
    clientInfo?.name || '',
    clientInfo?.address || '',
    documentDetails?.clientPhone || '',
    documentDetails?.warrantyInstallDate || ''
  ];
  let nextTable = table;
  rows.forEach((row, idx) => {
    if (idx >= values.length) return;
    const nextRow = warrantySetRowCells(row, [warrantyTextFromXml(warrantySplitCells(row)[0] || ''), values[idx]]);
    nextTable = nextTable.replace(row, nextRow);
  });
  return documentXml.replace(table, nextTable);
};

const warrantyReplaceEquipmentTable = (documentXml, warrantyRows) => {
  const table = warrantyFindTable(documentXml, (tbl) => tbl.includes('Обладнання') && tbl.includes('Серійний №') && tbl.includes('Термін гарантії'));
  if (!table) return documentXml;
  const rows = warrantySplitRows(table);
  if (rows.length < 2) return documentXml;
  const templateRow = rows[1];
  const generatedRows = warrantyRows.map((row) => warrantySetRowCells(templateRow, [
    row.equipment,
    row.model,
    row.serial,
    row.qty,
    row.warranty
  ])).join('');
  const nextTable = table.replace(rows.slice(1).join(''), generatedRows);
  return documentXml.replace(table, nextTable);
};

async function exportWarrantyDocxFile({
  clientInfo = {},
  documentDetails = {},
  calculations = {},
  workspaceHandle,
  projectFolderName = ''
}) {
  if (typeof window.JSZip === 'undefined') {
    alert('Бібліотека JSZip не завантажена. Оновіть сторінку та спробуйте ще раз.');
    return;
  }

  try {
    const response = await fetch(WARRANTY_TEMPLATE_FILE, { cache: 'no-store' });
    if (!response.ok) throw new Error(`template_http_${response.status}`);
    const buffer = await response.arrayBuffer();
    const zip = await window.JSZip.loadAsync(buffer);
    let documentXml = await zip.file('word/document.xml').async('string');

    const specNumber = String(documentDetails?.specNumber || '').trim() || '_______';
    const warrantyDate = warrantyFormatDate(documentDetails?.warrantyInstallDate || documentDetails?.contractDate || new Date().toISOString().slice(0, 10));
    const headerLine = `№ ${specNumber}   від  ${warrantyDate} р.`;
    documentXml = documentXml.replace(/№ _______[\s\S]*?___ \. ___ \. _______ р\./, warrantyXmlEscape(headerLine));

    documentXml = warrantyReplaceInfoTable(documentXml, clientInfo, {
      ...documentDetails,
      warrantyInstallDate: warrantyFormatDate(documentDetails?.warrantyInstallDate || '')
    });
    documentXml = warrantyReplaceEquipmentTable(documentXml, buildWarrantyRows(calculations));

    zip.file('word/document.xml', documentXml);
    const outBlob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    const safeClient = String(clientInfo?.name || 'Клієнт').trim().replace(/\s+/g, '_') || 'Клієнт';
    const fileName = `Гарантійний_талон_${safeClient}.docx`;
    await saveToDiskUtility(
      workspaceHandle,
      clientInfo,
      calculations,
      fileName,
      outBlob,
      'Гарантійний талон',
      projectFolderName
    );
  } catch (error) {
    console.error('Warranty DOCX export error', error);
    alert('Помилка при створенні гарантійного талона. Перевірте консоль браузера.');
  }
}

window.exportWarrantyDocxFile = exportWarrantyDocxFile;
