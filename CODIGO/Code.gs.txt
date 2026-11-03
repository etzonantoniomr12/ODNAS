const CFG = Object.freeze({
  ROOT_FOLDER_ID: '1hjiTz7LNl1L2gYDZ14Po1iZAQ4Kr67nl',
  STATE_FILE: '_APP_ODNA_ESTADO.json',
  MIME_XLSX: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  MIME_SHEET: 'application/vnd.google-apps.spreadsheet',
  FIRST_ITEM_ROW: 8,
  ACTIVITY_SPREADSHEET_ID: '1BLfuhEulYSoQX7IAohFTEPf_l0Dk2q1Lk7bY7CQEm34',
  ACTIVITY_SHEET: 'ACT',
  ACTIVITY_RANGE: 'C2:C52'
});

function doGet() {
  return json_({ ok: true, data: { service: 'ODNA API', version: 1 } });
}

function doPost(event) {
  try {
    const request = JSON.parse(event && event.postData ? event.postData.contents : '{}');
    const actions = {
      createProject: () => createProject_(request),
      processNext: () => processNext_(request),
      retryErrors: () => retryErrors_(request),
      getProject: () => getProject_(request),
      getReview: () => getReview_(request),
      applyReview: () => applyReview_(request),
      generatePdf: () => generatePdf_(request)
    };
    if (!actions[request.action]) throw new Error('Accion no permitida.');
    return json_({ ok: true, data: actions[request.action]() });
  } catch (error) {
    console.log(`ODNA API ERROR: ${error.stack || error.message || error}`);
    return json_({ ok: false, error: String(error.message || error) });
  }
}

function createProject_(request) {
  const sourceUrl = clean_(request.sourceUrl);
  const destinationUrl = clean_(request.destinationUrl);
  if (!sourceUrl) throw new Error('Pega el enlace de la carpeta origen.');
  if (!destinationUrl) throw new Error('Pega el enlace de la carpeta destino.');

  const source = getFolder_(driveRef_(sourceUrl));
  const folder = getFolder_(driveRef_(destinationUrl));
  assertFolderInsideRoot_(folder);
  if (folder.getFilesByName(CFG.STATE_FILE).hasNext()) throw new Error('La carpeta destino ya contiene un proyecto ODNA iniciado.');
  const files = [];
  collectXlsx_(source, source.getName(), files);
  if (!files.length) throw new Error('No se encontraron archivos .xlsx.');

  const project = {
    id: folder.getId(), name: folder.getName(), sourceUrl, destinationUrl,
    folderUrl: folder.getUrl(), createdAt: new Date().toISOString(),
    total: files.length, pending: files.length,
    documents: files.map(file => ({ sourceId: file.id, sourceKey: file.key, name: file.name.replace(/\.xlsx$/i, ''), path: file.path, status: 'PENDING', outputId: '', error: '', applied: false }))
  };
  saveProject_(folder, project);
  return publicProject_(project);
}

function processNext_(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folder = getProjectFolder_(request.projectId);
    const project = loadProject_(folder);
    const document = project.documents.find(item => item.status === 'PENDING' || item.status === 'PROCESSING');
    if (!document) return { project: publicProject_(project) };
    document.status = 'PROCESSING';
    saveProject_(folder, project);
    try {
      if (!document.outputId) document.outputId = copyAndConvert_(document, folder).id;
      const book = openSheetRetry_(document.outputId);
      transformWorkbook_(book);
      document.status = 'DONE';
      document.error = '';
    } catch (error) {
      document.status = 'ERROR';
      document.error = clean_(error.message || error).slice(0, 500);
    }
    project.pending = project.documents.filter(item => item.status === 'PENDING' || item.status === 'PROCESSING').length;
    saveProject_(folder, project);
    return { project: publicProject_(project), document };
  } finally { lock.releaseLock(); }
}

function retryErrors_(request) {
  const folder = getProjectFolder_(request.projectId);
  const project = loadProject_(folder);
  let count = 0;
  project.documents.forEach(document => {
    if (document.status !== 'ERROR') return;
    document.status = 'PENDING';
    document.error = '';
    count++;
  });
  saveProject_(folder, project);
  return { count, project: publicProject_(project) };
}

function getProject_(request) {
  const folder = getProjectFolder_(request.projectId);
  const project = loadProject_(folder);
  return { project: publicProject_(project), documents: project.documents };
}

function getReview_(request) {
  const projectFolder = getProjectFolder_(request.projectId);
  const project = loadProject_(projectFolder);
  const sheetId = clean_(request.sheetId);
  const document = project.documents.find(item => item.outputId === sheetId);
  if (!document) throw new Error('El documento no pertenece a este proyecto.');
  const book = SpreadsheetApp.openById(sheetId);
  const main = findMainSheet_(book);
  const lastRow = lastItemRow_(main);
  if (lastRow < CFG.FIRST_ITEM_ROW) throw new Error('No se encontraron items desde la fila 8.');
  const values = main.getRange(CFG.FIRST_ITEM_ROW, 1, lastRow - CFG.FIRST_ITEM_ROW + 1, 6).getDisplayValues();
  const items = values.map((row, index) => ({
    key: String(CFG.FIRST_ITEM_ROW + index),
    row: CFG.FIRST_ITEM_ROW + index,
    item: clean_(row[0]), defect: clean_(row[1]), startKm: clean_(row[2]), endKm: clean_(row[3]),
    deadline: clean_(row[4]), dueDate: clean_(row[5]), photoUrl: ''
  })).filter(item => item.item);
  const photoMap = ensurePhotoMap_(document, projectFolder, items.length);
  items.forEach(item => { item.photoUrl = photoMap[item.item] || ''; });
  saveProject_(projectFolder, project);
  const activities = getActivities_();
  return { sheetId, name: document.name, items, activities, decisions: document.decisions || {} };
}

function applyReview_(request) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const folder = getProjectFolder_(request.projectId);
    const project = loadProject_(folder);
    const document = project.documents.find(item => item.outputId === clean_(request.sheetId));
    if (!document) throw new Error('Documento no encontrado en el proyecto.');
    const decisions = request.decisions || {};
    const book = SpreadsheetApp.openById(document.outputId);
    const main = findMainSheet_(book);
    const allowedActivities = new Set(getActivities_().map(activity => activity.text));
    const lastRow = lastItemRow_(main);
    if (lastRow < CFG.FIRST_ITEM_ROW) throw new Error('No se encontraron items desde la fila 8.');
    const items = main.getRange(CFG.FIRST_ITEM_ROW, 1, lastRow - CFG.FIRST_ITEM_ROW + 1, 1).getDisplayValues();
    const formatRange = main.getRange(CFG.FIRST_ITEM_ROW, 1, items.length, 6);
    const backgrounds = formatRange.getBackgrounds();
    const fontColors = formatRange.getFontColors();
    document.originalFormats = document.originalFormats || {};
    let formatsAdded = false;
    for (let i = 0; i < items.length; i++) {
      const row = CFG.FIRST_ITEM_ROW + i;
      if (clean_(items[i][0]) && !document.originalFormats[String(row)]) {
        document.originalFormats[String(row)] = { backgrounds: backgrounds[i], fontColors: fontColors[i] };
        formatsAdded = true;
      }
    }
    if (formatsAdded) saveProject_(folder, project);
    let normal = 0, caution = 0;
    for (let i = 0; i < items.length; i++) {
      const row = CFG.FIRST_ITEM_ROW + i;
      const item = clean_(items[i][0]);
      if (!item) continue;
      const decision = decisions[String(row)] || decisions[item] || {};
      if (!decision.activity && !decision.caution) throw new Error(`Falta revisar el item ${item}.`);
      if (!decision.caution && !allowedActivities.has(clean_(decision.activity))) throw new Error(`Actividad no valida para el item ${item}.`);
      const rowRange = main.getRange(row, 1, 1, 6);
      const cell = main.getRange(row, 2);
      const original = cleanOriginalDefect_(cell.getDisplayValue());
      const originalFormat = document.originalFormats[String(row)];
      if (decision.caution) {
        cell.setValue(original);
        rowRange.setBackground('#D32F2F').setFontColor('#FFFFFF');
        caution++;
      } else {
        rowRange.setBackgrounds([originalFormat.backgrounds]).setFontColors([originalFormat.fontColors]);
        const suffix = `\nACTIVIDAD OFERTADA: ${clean_(decision.activity)}`;
        const full = original + suffix;
        const rich = SpreadsheetApp.newRichTextValue().setText(full)
          .setTextStyle(original.length, full.length, SpreadsheetApp.newTextStyle().setForegroundColor('#D00000').setBold(true).build()).build();
        cell.setRichTextValue(rich).setWrap(true).setVerticalAlignment('middle');
        normal++;
      }
      cell.setWrap(true);
    }
    main.autoResizeRows(CFG.FIRST_ITEM_ROW, items.length);
    document.decisions = decisions;
    document.applied = true;
    document.appliedAt = new Date().toISOString();
    saveProject_(folder, project);
    console.log(`ACTIVIDAD OFERTADA | normales: ${normal} | cautelares: ${caution}`);
    return { normal, caution };
  } finally { lock.releaseLock(); }
}

function generatePdf_(request) {
  const folder = getProjectFolder_(request.projectId);
  const project = loadProject_(folder);
  const document = project.documents.find(item => item.outputId === clean_(request.sheetId));
  if (!document) throw new Error('Documento no encontrado en el proyecto.');
  if (!document.applied) throw new Error('Revisa y aplica todos los items antes de generar el PDF.');
  const pdfName = `${document.name}.pdf`;
  const existing = folder.getFilesByName(pdfName);
  const url = `https://docs.google.com/spreadsheets/d/${document.outputId}/export?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=true&printtitle=false&pagenumbers=true&gridlines=false&fzr=true`;
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error(`La exportacion PDF respondio ${response.getResponseCode()}.`);
  const pdfBlob = response.getBlob().setName(pdfName);
  let pdfId, updated = false;
  if (existing.hasNext()) {
    pdfId = existing.next().getId();
    const update = UrlFetchApp.fetch(`https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(pdfId)}?uploadType=media`, {
      method: 'patch', contentType: 'application/pdf', headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      payload: pdfBlob.getBytes(), muteHttpExceptions: true
    });
    if (update.getResponseCode() < 200 || update.getResponseCode() >= 300) throw new Error(`Drive no pudo actualizar el PDF: ${update.getContentText().slice(0, 300)}`);
    updated = true;
  } else {
    pdfId = folder.createFile(pdfBlob).getId();
  }
  document.pdfId = pdfId;
  saveProject_(folder, project);
  return { id: pdfId, updated };
}

function copyAndConvert_(document, targetFolder) {
  const metadata = { title: document.name, mimeType: CFG.MIME_SHEET, parents: [{ id: targetFolder.getId() }] };
  const headers = { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` };
  if (document.sourceKey) headers['X-Goog-Drive-Resource-Keys'] = `${document.sourceId}/${document.sourceKey}`;
  const url = `https://www.googleapis.com/drive/v2/files/${encodeURIComponent(document.sourceId)}/copy?convert=true&supportsAllDrives=true`;
  const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers, payload: JSON.stringify(metadata), muteHttpExceptions: true });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error(`Drive no pudo convertir '${document.name}': ${response.getContentText().slice(0, 300)}`);
  return JSON.parse(response.getContentText());
}

function transformWorkbook_(book) {
  const main = findMainSheet_(book);
  const width = main.getColumnWidth(2) + main.getColumnWidth(3) + main.getColumnWidth(4);
  main.getRange('C2').copyTo(main.getRange('E2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  main.deleteColumns(3, 2);
  main.setColumnWidth(2, width);
  const last = lastItemRow_(main);
  if (last < CFG.FIRST_ITEM_ROW) throw new Error('No se encontraron items desde la fila 8.');
  const range = main.getRange(CFG.FIRST_ITEM_ROW, 1, last - CFG.FIRST_ITEM_ROW + 1, 6);
  range.setValues(range.getValues());
  SpreadsheetApp.flush();
}

function findMainSheet_(book) {
  const sheet = book.getSheets().find(s => {
    if (s.isSheetHidden() || s.getMaxRows() < 8 || s.getMaxColumns() < 8) return false;
    const headers = s.getRange('A7:B7').getDisplayValues()[0].map(normalizeHeader_);
    return /ITEM/.test(headers[0]) && /DEFECTO/.test(headers[1]);
  });
  if (!sheet) throw new Error('No se encontro la hoja principal con encabezados de ITEM y DEFECTO en la fila 7.');
  return sheet;
}

function ensurePhotoMap_(document, projectFolder, itemCount) {
  const saved = document.photoMap || {};
  if (Object.keys(saved).filter(key => saved[key]).length >= itemCount) return saved;

  const source = document.sourceKey
    ? DriveApp.getFileByIdAndResourceKey(document.sourceId, document.sourceKey)
    : DriveApp.getFileById(document.sourceId);
  const sizeMb = source.getSize() / 1024 / 1024;
  let extraction;
  try {
    extraction = extractPicturesFromXlsx_(source, document.sourceKey);
  } catch (error) {
    throw new Error(`No se pudieron extraer las fotos del XLSX (${sizeMb.toFixed(1)} MB) mediante lectura parcial: ${clean_(error.message || error)}`);
  }
  const pictures = extraction.pictures;
  if (!pictures.length) throw new Error('El XLSX no contiene fotografias mapeables en hojas de sustento.');

  const folderName = `_FOTOS_${document.name}`.replace(/[\\/:*?"<>|]/g, '_').slice(0, 180);
  const folders = projectFolder.getFoldersByName(folderName);
  const photoFolder = folders.hasNext() ? folders.next() : projectFolder.createFolder(folderName);
  const result = {};
  for (let i = 0; i < Math.min(itemCount, pictures.length); i++) {
    const item = String(i + 1);
    const extension = imageExtension_(pictures[i].path);
    const fileName = `ITEM_${String(i + 1).padStart(3, '0')}.${extension}`;
    const existing = photoFolder.getFilesByName(fileName);
    const pictureBlob = readZipEntryBlob_(extraction.archive, pictures[i].path);
    const file = existing.hasNext() ? existing.next() : photoFolder.createFile(
      pictureBlob.copyBlob().setName(fileName).setContentType(imageMimeType_(extension))
    );
    result[item] = drivePhotoUrl_(file);
  }
  document.photoMap = result;
  document.photoFolderId = photoFolder.getId();
  console.log(`FOTOS XLSX | archivo: ${document.name} | extraidas: ${pictures.length} | vinculadas: ${Object.keys(result).length}`);
  return result;
}

function extractPicturesFromXlsx_(sourceFile, resourceKey) {
  const archive = openRemoteZip_(sourceFile, resourceKey);
  const workbookPath = 'xl/workbook.xml';
  const workbook = xmlRoot_(archive, workbookPath);
  const workbookRels = relationshipMap_(xmlRoot_(archive, 'xl/_rels/workbook.xml.rels'));
  const spreadsheetNs = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  const officeRelNs = XmlService.getNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  const sheetsNode = workbook.getChild('sheets', spreadsheetNs);
  if (!sheetsNode) throw new Error('El XLSX no contiene la estructura de hojas esperada.');
  const pictures = [];
  sheetsNode.getChildren('sheet', spreadsheetNs).forEach((sheetNode, sheetIndex) => {
    const sheetName = sheetNode.getAttribute('name').getValue();
    if (!/SUSTENTO|FOTO|PANEL/i.test(sheetName)) return;
    const relId = sheetNode.getAttribute('id', officeRelNs).getValue();
    const sheetTarget = workbookRels[relId];
    if (!sheetTarget) return;
    const sheetPath = resolveZipPath_(workbookPath, sheetTarget);
    const sheetRoot = xmlRoot_(archive, sheetPath);
    const sheetRelsPath = `${zipDir_(sheetPath)}/_rels/${zipBase_(sheetPath)}.rels`;
    if (!archive.entries[sheetRelsPath]) return;
    const sheetRels = relationshipMap_(xmlRoot_(archive, sheetRelsPath));
    sheetRoot.getChildren('drawing', spreadsheetNs).forEach(drawingNode => {
      const drawingId = drawingNode.getAttribute('id', officeRelNs).getValue();
      const drawingTarget = sheetRels[drawingId];
      if (!drawingTarget) return;
      const drawingPath = resolveZipPath_(sheetPath, drawingTarget);
      readDrawingPictures_(archive, drawingPath, sheetIndex, pictures);
    });
  });
  pictures.sort((a, b) => a.sheetIndex - b.sheetIndex || a.row - b.row || a.col - b.col);
  return { archive, pictures };
}

function readDrawingPictures_(archive, drawingPath, sheetIndex, output) {
  const root = xmlRoot_(archive, drawingPath);
  const drawingRelsPath = `${zipDir_(drawingPath)}/_rels/${zipBase_(drawingPath)}.rels`;
  if (!archive.entries[drawingRelsPath]) return;
  const rels = relationshipMap_(xmlRoot_(archive, drawingRelsPath));
  const xdrNs = XmlService.getNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
  const drawingNs = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
  const officeRelNs = XmlService.getNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  const anchors = root.getChildren('twoCellAnchor', xdrNs).concat(root.getChildren('oneCellAnchor', xdrNs));
  anchors.forEach(anchor => {
    const from = anchor.getChild('from', xdrNs);
    const pic = anchor.getChild('pic', xdrNs);
    if (!from || !pic) return;
    const row = Number(from.getChildText('row', xdrNs)) + 1;
    const col = Number(from.getChildText('col', xdrNs)) + 1;
    if (row < 5 || (row - 5) % 4 !== 0) return;
    const blipFill = pic.getChild('blipFill', xdrNs);
    const blip = blipFill && blipFill.getChild('blip', drawingNs);
    const embed = blip && blip.getAttribute('embed', officeRelNs);
    if (!embed || !rels[embed.getValue()]) return;
    const mediaPath = resolveZipPath_(drawingPath, rels[embed.getValue()]);
    if (!archive.entries[mediaPath]) return;
    output.push({ sheetIndex, row, col, path: mediaPath });
  });
}

function relationshipMap_(root) {
  const ns = XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships');
  const result = {};
  root.getChildren('Relationship', ns).forEach(rel => {
    result[rel.getAttribute('Id').getValue()] = rel.getAttribute('Target').getValue();
  });
  return result;
}

function xmlRoot_(archive, path) {
  const normalized = normalizeZipName_(path);
  if (!archive.entries[normalized]) throw new Error(`Falta el archivo interno ${path}.`);
  archive.xmlCache = archive.xmlCache || {};
  if (!archive.xmlCache[normalized]) {
    archive.xmlCache[normalized] = XmlService.parse(readZipEntryBlob_(archive, normalized).getDataAsString()).getRootElement();
  }
  return archive.xmlCache[normalized];
}

function openRemoteZip_(sourceFile, resourceKey) {
  const size = Number(sourceFile.getSize());
  if (!size) throw new Error('El archivo XLSX esta vacio.');
  const tailStart = Math.max(0, size - 65557);
  const tail = fetchDriveRange_(sourceFile.getId(), resourceKey, tailStart, size - 1);
  let eocd = -1;
  for (let i = tail.length - 22; i >= 0; i--) {
    if (readU32_(tail, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('No se encontro el directorio ZIP del XLSX.');
  const totalEntries = readU16_(tail, eocd + 10);
  const centralSize = readU32_(tail, eocd + 12);
  const centralOffset = readU32_(tail, eocd + 16);
  if (totalEntries === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error('El XLSX usa ZIP64, formato no compatible con el lector parcial.');
  }
  const central = fetchDriveRange_(sourceFile.getId(), resourceKey, centralOffset, centralOffset + centralSize - 1);
  const entries = {};
  let pos = 0;
  for (let index = 0; index < totalEntries && pos + 46 <= central.length; index++) {
    if (readU32_(central, pos) !== 0x02014b50) throw new Error('El directorio ZIP esta danado o incompleto.');
    const nameLength = readU16_(central, pos + 28);
    const extraLength = readU16_(central, pos + 30);
    const commentLength = readU16_(central, pos + 32);
    const entryLength = 46 + nameLength + extraLength + commentLength;
    const name = normalizeZipName_(decodeZipText_(central.slice(pos + 46, pos + 46 + nameLength)));
    entries[name] = {
      name,
      flags: readU16_(central, pos + 8),
      method: readU16_(central, pos + 10),
      compressedSize: readU32_(central, pos + 20),
      uncompressedSize: readU32_(central, pos + 24),
      localOffset: readU32_(central, pos + 42),
      centralBytes: central.slice(pos, pos + entryLength)
    };
    pos += entryLength;
  }
  return { fileId: sourceFile.getId(), resourceKey: resourceKey || '', entries };
}

function readZipEntryBlob_(archive, path) {
  const normalized = normalizeZipName_(path);
  const entry = archive.entries[normalized];
  if (!entry) throw new Error(`No existe ${normalized} dentro del XLSX.`);
  if (entry.flags & 1) throw new Error(`La entrada ${normalized} esta cifrada.`);
  if (entry.method !== 0 && entry.method !== 8) throw new Error(`Compresion ZIP no compatible en ${normalized}.`);
  const header = fetchDriveRange_(archive.fileId, archive.resourceKey, entry.localOffset, entry.localOffset + 29);
  if (readU32_(header, 0) !== 0x04034b50) throw new Error(`Cabecera ZIP invalida en ${normalized}.`);
  const localNameLength = readU16_(header, 26);
  const localExtraLength = readU16_(header, 28);
  const dataEnd = entry.localOffset + 30 + localNameLength + localExtraLength + entry.compressedSize;
  let descriptorLength = 0;
  if (entry.flags & 8) {
    const descriptor = fetchDriveRange_(archive.fileId, archive.resourceKey, dataEnd, dataEnd + 15);
    descriptorLength = readU32_(descriptor, 0) === 0x08074b50 ? 16 : 12;
  }
  const recordLength = 30 + localNameLength + localExtraLength + entry.compressedSize + descriptorLength;
  const localRecord = fetchDriveRange_(archive.fileId, archive.resourceKey, entry.localOffset, entry.localOffset + recordLength - 1);
  const centralRecord = entry.centralBytes.slice();
  writeU32_(centralRecord, 42, 0);
  const eocd = [0x50,0x4b,0x05,0x06,0,0,0,0,1,0,1,0,0,0,0,0,0,0,0,0,0,0];
  writeU32_(eocd, 12, centralRecord.length);
  writeU32_(eocd, 16, localRecord.length);
  const zipBlob = Utilities.newBlob(toSignedBytes_(localRecord.concat(centralRecord, eocd)), 'application/zip', 'entrada.zip');
  const unpacked = Utilities.unzip(zipBlob);
  if (!unpacked.length) throw new Error(`No se pudo descomprimir ${normalized}.`);
  return unpacked[0];
}

function fetchDriveRange_(fileId, resourceKey, start, end) {
  const headers = { Authorization: `Bearer ${ScriptApp.getOAuthToken()}`, Range: `bytes=${start}-${end}` };
  if (resourceKey) headers['X-Goog-Drive-Resource-Keys'] = `${fileId}/${resourceKey}`;
  const response = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
    headers, muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code !== 206 && code !== 200) throw new Error(`Drive respondio ${code} al leer bytes ${start}-${end}.`);
  return Array.from(response.getContent(), value => value & 255);
}

function readU16_(bytes, offset) { return (bytes[offset] & 255) + (bytes[offset + 1] & 255) * 256; }
function readU32_(bytes, offset) { return readU16_(bytes, offset) + readU16_(bytes, offset + 2) * 65536; }
function writeU32_(bytes, offset, value) { bytes[offset]=value&255; bytes[offset+1]=(value>>>8)&255; bytes[offset+2]=(value>>>16)&255; bytes[offset+3]=(value>>>24)&255; }
function toSignedBytes_(bytes) { return bytes.map(value => { const byte = value & 255; return byte > 127 ? byte - 256 : byte; }); }
function decodeZipText_(bytes) { return Utilities.newBlob(toSignedBytes_(bytes)).getDataAsString('UTF-8'); }

function resolveZipPath_(baseFile, target) {
  if (/^\//.test(target)) return normalizeZipName_(target.replace(/^\/+/, ''));
  const parts = zipDir_(baseFile).split('/').filter(Boolean).concat(String(target).split('/'));
  const resolved = [];
  parts.forEach(part => {
    if (!part || part === '.') return;
    if (part === '..') resolved.pop(); else resolved.push(part);
  });
  return normalizeZipName_(resolved.join('/'));
}

function normalizeZipName_(name) { return String(name || '').replace(/\\/g, '/').replace(/^\/+/, ''); }
function zipDir_(path) { const value = normalizeZipName_(path); return value.includes('/') ? value.slice(0, value.lastIndexOf('/')) : ''; }
function zipBase_(path) { const value = normalizeZipName_(path); return value.slice(value.lastIndexOf('/') + 1); }
function imageExtension_(name) { const match = String(name || '').toLowerCase().match(/\.(png|jpe?g|gif|bmp|webp)$/); return match ? (match[1] === 'jpeg' ? 'jpg' : match[1]) : 'png'; }
function imageMimeType_(extension) { return ({ png:'image/png', jpg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', webp:'image/webp' })[extension] || 'image/png'; }
function drivePhotoUrl_(file) { const key = file.getResourceKey(); return `https://drive.google.com/thumbnail?id=${encodeURIComponent(file.getId())}&sz=w1600${key ? `&resourcekey=${encodeURIComponent(key)}` : ''}`; }

function getActivities_() {
  const catalog = SpreadsheetApp.openById(CFG.ACTIVITY_SPREADSHEET_ID);
  const sheet = catalog.getSheetByName(CFG.ACTIVITY_SHEET);
  if (!sheet) throw new Error(`No existe la hoja '${CFG.ACTIVITY_SHEET}' en el catalogo central.`);
  const range = sheet.getRange(CFG.ACTIVITY_RANGE);
  return range.getDisplayValues().flat()
    .map((value, index) => ({ row: range.getRow() + index, text: clean_(value) }))
    .filter(activity => activity.text);
}

function collectXlsx_(folder, path, output) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== CFG.MIME_XLSX && !/\.xlsx$/i.test(file.getName())) continue;
    output.push({ id: file.getId(), key: file.getResourceKey() || '', name: file.getName(), path: `${path}/${file.getName()}` });
  }
  const folders = folder.getFolders();
  while (folders.hasNext()) { const child = folders.next(); collectXlsx_(child, `${path}/${child.getName()}`, output); }
}

function saveProject_(folder, project) {
  project.pending = project.documents.filter(item => item.status === 'PENDING' || item.status === 'PROCESSING').length;
  const files = folder.getFilesByName(CFG.STATE_FILE);
  if (files.hasNext()) files.next().setContent(JSON.stringify(project));
  else folder.createFile(CFG.STATE_FILE, JSON.stringify(project), MimeType.PLAIN_TEXT);
}
function loadProject_(folder) { const files = folder.getFilesByName(CFG.STATE_FILE); if (!files.hasNext()) throw new Error('No se encontro el estado del proyecto.'); return JSON.parse(files.next().getBlob().getDataAsString()); }
function publicProject_(p) { return { id:p.id,name:p.name,folderUrl:p.folderUrl,total:p.total,pending:p.pending,createdAt:p.createdAt }; }
function lastItemRow_(sheet) {
  const start = CFG.FIRST_ITEM_ROW;
  const count = sheet.getLastRow() - start + 1;
  if (count <= 0) return start - 1;
  const rows = sheet.getRange(start, 1, count, 6).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].every(value => !clean_(value))) return start + i - 1;
  }
  return start + rows.length - 1;
}
function getProjectFolder_(projectId) { const id=clean_(projectId); if(!id)throw new Error('Falta el ID del proyecto.'); const folder=DriveApp.getFolderById(id); assertFolderInsideRoot_(folder); return folder; }
function assertFolderInsideRoot_(folder) { const rootId=CFG.ROOT_FOLDER_ID; const pending=[folder]; const visited={}; while(pending.length){const current=pending.pop(); const id=current.getId(); if(visited[id])continue; visited[id]=true; if(id===rootId)return; const parents=current.getParents(); while(parents.hasNext())pending.push(parents.next());} throw new Error('La carpeta destino no pertenece a la raiz ODNA autorizada.'); }
function cleanOriginalDefect_(text) { return String(text == null ? '' : text).replace(/\nACTIVIDAD OFERTADA:[\s\S]*$/i, '').trim(); }
function openSheetRetry_(id) { let error; for(let i=0;i<5;i++){try{return SpreadsheetApp.openById(id);}catch(e){error=e;Utilities.sleep(2000);}} throw error; }
function driveRef_(text) { const value=clean_(text); const id=((value.match(/\/folders\/([\w-]+)/)||[])[1]||(value.match(/\/d\/([\w-]+)/)||[])[1]||(value.match(/[?&]id=([\w-]+)/)||[])[1]); if(!id)throw new Error('No se pudo extraer el ID de Drive.'); return {id,key:((value.match(/[?&]resourcekey=([\w-]+)/i)||[])[1]||'')}; }
function getFolder_(ref) { return ref.key ? DriveApp.getFolderByIdAndResourceKey(ref.id,ref.key) : DriveApp.getFolderById(ref.id); }
function clean_(value) { return String(value == null ? '' : value).replace(/\u00a0/g,' ').replace(/ {2,}/g,' ').trim(); }
function normalizeHeader_(value) { return clean_(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase(); }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }

