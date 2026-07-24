import { ORDER_OF_WORSHIP_EXPORT_FOLDER_URL, SERVICES_COL, SERVICES_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';
import { getOrder, type OrderItem } from './order';
import { esvPassage, lblaPassage } from './services';

type ExportOrderOfWorshipDocInput = {
  serviceId?: string;
  folderUrl?: string;
  previousFileId?: string;
};

type ServiceRecord = {
  id: string;
  date: string;
  time: string;
  type: string;
  leader: string;
  preacher: string;
  scripture: string;
  scriptureText: string;
};

type ReadingContent = {
  heading: string;
  intro: string;
  reader: string;
  reference: string;
  englishText: string;
  spanishText: string;
  firstTranslation: 'ESV' | 'LBLA';
};

type ParagraphOptions = {
  align?: 'center' | 'left';
  bold?: boolean;
  italic?: boolean;
  size?: number;
  spacingAfter?: number;
};

type TextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
};

type TableCellSpec = {
  span?: number;
  paragraph: {
    segments: TextRun[];
    size?: number;
  };
};

const ITEM_CALL_TO_WORSHIP = 'call to worship';
const ITEM_SECOND_SCRIPTURE = '2nd scripture';
const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export function exportOrderOfWorshipDoc(input?: ExportOrderOfWorshipDocInput) {
  const serviceId = String(input?.serviceId || '').trim();
  if (!serviceId) throw new Error('serviceId is required.');

  const service = getServiceById(serviceId);
  if (!service) throw new Error(`Service not found: ${serviceId}`);

  const order = getOrder(serviceId);
  const items = Array.isArray(order?.items) ? order.items : [];
  if (!items.length) throw new Error('Save the order of worship before exporting the document.');

  const callToWorship = findOrderItem(items, ITEM_CALL_TO_WORSHIP);
  const secondScripture = findOrderItem(items, ITEM_SECOND_SCRIPTURE);
  if (!callToWorship?.detail) throw new Error('Add a Call to Worship reference before exporting.');
  if (!secondScripture?.detail) throw new Error('Add a 2nd Scripture reference before exporting.');

  const folderUrl = resolveFolderUrl(input?.folderUrl);
  const folderId = extractFolderId(folderUrl);
  if (!folderId) throw new Error('The Order of Worship export folder URL is invalid.');
  const folder = DriveApp.getFolderById(folderId);

  const callReading = buildReadingContent('call', callToWorship, service);
  const secondReading = buildReadingContent('second', secondScripture, service);
  const docxName = buildDocxFileName(service);
  const docxBlob = buildDocxBlob(service, items, callReading, secondReading);
  const previousFileId = String(input?.previousFileId || '').trim();
  trashExistingExports(folder, folderId, docxName, previousFileId);

  docxBlob.setName(docxName);
  const saved = folder.createFile(docxBlob);
  return {
    ok: true,
    fileId: saved.getId(),
    name: saved.getName(),
    url: saved.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function trashExistingExports(
  folder: GoogleAppsScript.Drive.Folder,
  folderId: string,
  desiredName: string,
  previousFileId: string
) {
  const trashedIds: Record<string, true> = {};
  if (previousFileId) {
    trashExistingExportById(previousFileId, folderId);
    trashedIds[previousFileId] = true;
  }

  const familyPattern = exportFileNamePattern(desiredName);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileId = String(file.getId() || '').trim();
    const fileName = String(file.getName() || '').trim();
    if (!fileId || trashedIds[fileId]) continue;
    if (!familyPattern.test(fileName)) continue;
    file.setTrashed(true);
  }
}

function trashExistingExportById(fileId: string, folderId: string) {
  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    let inTargetFolder = false;
    while (parents.hasNext()) {
      const parent = parents.next();
      if (String(parent.getId() || '').trim() === folderId) {
        inTargetFolder = true;
        break;
      }
    }
    if (!inTargetFolder) throw new Error('Existing export is not in the configured folder.');
    file.setTrashed(true);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err || '');
    throw new Error(`Unable to replace the existing DOCX export. ${message}`.trim());
  }
}

function exportFileNamePattern(desiredName: string) {
  const dot = desiredName.lastIndexOf('.');
  const stem = dot > 0 ? desiredName.slice(0, dot) : desiredName;
  const ext = dot > 0 ? desiredName.slice(dot) : '';
  return new RegExp(`^${escapeRegex(stem)}(?: \\(\\d+\\))?${escapeRegex(ext)}$`, 'i');
}

function resolveFolderUrl(folderUrl?: string) {
  const raw = String(folderUrl || '').trim();
  if (raw) return raw;
  const prop = String(PropertiesService.getScriptProperties().getProperty('ORDER_OF_WORSHIP_EXPORT_FOLDER_URL') || '').trim();
  return prop || ORDER_OF_WORSHIP_EXPORT_FOLDER_URL;
}

function extractFolderId(folderUrl: string) {
  const direct = /\/folders\/([A-Za-z0-9_-]+)/.exec(String(folderUrl || ''));
  if (direct) return direct[1];
  const query = /[?&]id=([A-Za-z0-9_-]+)/.exec(String(folderUrl || ''));
  if (query) return query[1];
  return '';
}

function getServiceById(serviceId: string): ServiceRecord | null {
  const sh = getSheetByName(SERVICES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  if (idIdx < 0) return null;

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (const row of rows) {
    const id = String(row[idIdx] ?? '').trim();
    if (id !== serviceId) continue;
    return {
      id,
      date: toIsoDate(row[col(SERVICES_COL.date)]),
      time: String(row[col(SERVICES_COL.time)] ?? '').trim(),
      type: String(row[col(SERVICES_COL.type)] ?? '').trim(),
      leader: String(row[col(SERVICES_COL.leader)] ?? '').trim(),
      preacher: String(row[col(SERVICES_COL.preacher)] ?? '').trim(),
      scripture: String(row[col(SERVICES_COL.scripture)] ?? '').trim(),
      scriptureText: String(row[col(SERVICES_COL.scriptureText)] ?? '').trim()
    };
  }
  return null;
}

function toIsoDate(value: unknown) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(value ?? '').trim();
}

function findOrderItem(items: OrderItem[], itemType: string) {
  const target = normalizeItemType(itemType);
  return items.find(item => normalizeItemType(item?.itemType) === target) || null;
}

function normalizeItemType(value: unknown) {
  return String(value || '').trim().toLowerCase();
}

function buildReadingContent(kind: 'call' | 'second', item: OrderItem, service: ServiceRecord): ReadingContent {
  const reference = String(item.detail || '').trim();
  const englishText = getEnglishReadingText(item, service);
  const spanishText = getLblaPassageText(reference);
  const reader = String(item.leader || '').trim();

  if (!englishText) {
    throw new Error(`Unable to load ESV text for ${reference}. Save the scripture text or set ESV_API_TOKEN.`);
  }
  if (!spanishText) {
    throw new Error(`Unable to load LBLA text for ${reference}.`);
  }

  if (kind === 'call') {
    return {
      heading: 'Call To Worship',
      intro: `Please stand for the reading of our call to worship from ${reference}`,
      reader,
      reference,
      englishText,
      spanishText,
      firstTranslation: 'ESV'
    };
  }

  return {
    heading: 'Scripture Reading',
    intro: `Our second Scripture reading comes from ${reference}`,
    reader,
    reference,
    englishText,
    spanishText,
    firstTranslation: 'LBLA'
  };
}

function getEnglishReadingText(item: OrderItem, service: ServiceRecord) {
  const saved = String(item.scriptureText || '').trim();
  if (saved) return saved;
  const ref = String(item.detail || '').trim();
  if (ref) {
    const fetched = esvPassage({ reference: ref, html: false });
    const text = String(fetched?.text || '').trim();
    if (text) return text;
  }
  if (normalizeReference(ref) === normalizeReference(service.scripture)) {
    return String(service.scriptureText || '').trim();
  }
  return '';
}

function normalizeReference(value: unknown) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function getLblaPassageText(reference: string) {
  const ref = String(reference || '').trim();
  if (!ref) return '';
  return String(lblaPassage({ reference: ref })?.text || '').trim();
}

function buildDocxFileName(service: ServiceRecord) {
  const dateSource = String(service.date || '').trim() || deriveDateFromServiceId(service.id);
  const parts = dateSource.split('-');
  const shortDate = parts.length === 3
    ? `${parts[0].slice(-2)}.${parts[1]}.${parts[2]}`
    : 'service';
  const suffix = sanitizeFileName(service.type || 'Service');
  return `${shortDate} Order of Worship (${suffix}).docx`;
}

function deriveDateFromServiceId(serviceId: string) {
  const match = String(serviceId || '').match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : '';
}

function sanitizeFileName(value: string) {
  return String(value || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function nextAvailableFileName(folder: GoogleAppsScript.Drive.Folder, desiredName: string) {
  const dot = desiredName.lastIndexOf('.');
  const stem = dot > 0 ? desiredName.slice(0, dot) : desiredName;
  const ext = dot > 0 ? desiredName.slice(dot) : '';
  let nextName = desiredName;
  let index = 2;
  while (folder.getFilesByName(nextName).hasNext()) {
    nextName = `${stem} (${index})${ext}`;
    index += 1;
  }
  return nextName;
}

function escapeRegex(value: string) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function buildDocxBlob(
  service: ServiceRecord,
  items: OrderItem[],
  callReading: ReadingContent,
  secondReading: ReadingContent
) {
  const files = [
    xmlBlob('[Content_Types].xml', contentTypesXml()),
    xmlBlob('_rels/.rels', rootRelsXml()),
    xmlBlob('word/_rels/document.xml.rels', documentRelsXml()),
    xmlBlob('word/document.xml', documentXml(service, items, callReading, secondReading)),
    xmlBlob('word/header1.xml', headerXml(formatLongDate(service.date)))
  ];
  return Utilities.zip(files, 'order-of-worship.docx').setContentType(DOCX_MIME_TYPE);
}

function documentXml(
  service: ServiceRecord,
  items: OrderItem[],
  callReading: ReadingContent,
  secondReading: ReadingContent
) {
  const body = [
    paragraph('Order of Worship', { align: 'center', bold: true, size: 32, spacingAfter: 120 }),
    service.leader || service.preacher
      ? paragraph(compactMetaLine(service), { align: 'center', size: 20, spacingAfter: 160 })
      : '',
    orderTable(items),
    pageBreak(),
    readingPage(service, callReading),
    pageBreak(),
    readingPage(service, secondReading),
    sectionProperties()
  ].join('');

  return xmlDocument([
    '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:body>',
    body,
    '</w:body>',
    '</w:document>'
  ].join(''));
}

function headerXml(text: string) {
  return xmlDocument([
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    paragraph(text, { align: 'center', size: 22 }),
    '</w:hdr>'
  ].join(''));
}

function readingPage(service: ServiceRecord, reading: ReadingContent) {
  const parts = [
    paragraph(`${reading.heading}${reading.reader ? ` (${reading.reader})` : ''}`, { bold: true, italic: true, size: 40, spacingAfter: 260 }),
    reading.heading === 'Call To Worship'
      ? paragraphRuns([
        { text: 'Please stand', bold: true },
        { text: ' for the reading of our call to worship' },
        { text: `\nfrom ${reading.reference}` }
      ], { size: 32, spacingAfter: 260 })
      : paragraph(reading.intro, { size: 32, spacingAfter: 260 })
  ];

  if (reading.firstTranslation === 'ESV') {
    parts.push(translationBlock('ESV', reading.englishText));
    parts.push(paragraph('', { spacingAfter: 120 }));
    parts.push(translationBlock('LBLA', reading.spanishText));
  } else {
    parts.push(translationBlock('LBLA', reading.spanishText));
    parts.push(paragraph('', { spacingAfter: 120 }));
    parts.push(translationBlock('ESV', reading.englishText));
  }
  return parts.join('');
}

function translationBlock(label: 'ESV' | 'LBLA', text: string) {
  const parts = [paragraph(`${label}:`, { bold: true, size: 26, spacingAfter: 80 })];
  const blocks = String(text || '').split(/\n{2,}/).map(part => part.trim()).filter(Boolean);
  blocks.forEach(block => parts.push(paragraph(block, { size: 26, spacingAfter: 120 })));
  return parts.join('');
}

function orderTable(items: OrderItem[]) {
  const rows = [
    [
      headerCell('ITEM'),
      headerCell('DETAIL'),
      headerCell('LEADER/NOTES')
    ],
    ...items.map(item => buildOrderRow(item))
  ];
  const grid = [2800, 2900, 3660];
  const rowXml = rows.map((row) => {
    let columnIndex = 0;
    const cells = row.map((cell) => {
      const span = Math.max(1, Number(cell.span || 1));
      const width = grid.slice(columnIndex, columnIndex + span).reduce((sum, value) => sum + value, 0);
      columnIndex += span;
      return tableCell(cell, width);
    });
    return `<w:tr>${cells.join('')}</w:tr>`;
  }).join('');

  return [
    '<w:tbl>',
    '<w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="999999"/><w:left w:val="single" w:sz="4" w:color="999999"/><w:bottom w:val="single" w:sz="4" w:color="999999"/><w:right w:val="single" w:sz="4" w:color="999999"/><w:insideH w:val="single" w:sz="4" w:color="999999"/><w:insideV w:val="single" w:sz="4" w:color="999999"/></w:tblBorders><w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>',
    `<w:tblGrid>${grid.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`,
    rowXml,
    '</w:tbl>'
  ].join('');
}

function headerCell(text: string): TableCellSpec {
  return {
    paragraph: {
      segments: [{ text, bold: true }],
      size: 20
    }
  };
}

function buildOrderRow(item: OrderItem): TableCellSpec[] {
  return [
    textCell(tableItemFor(item)),
    textCell(tableDetailFor(item), { bold: shouldBoldDetail(item) }),
    leaderNotesCell(item)
  ];
}

function textCell(text: string, opts?: { bold?: boolean }): TableCellSpec {
  return {
    paragraph: {
      segments: [{ text: String(text || '').trim(), bold: !!opts?.bold }],
      size: 22
    }
  };
}

function tableCell(cell: TableCellSpec, width: number) {
  const span = Math.max(1, Number(cell.span || 1));
  const gridSpan = span > 1 ? `<w:gridSpan w:val="${span}"/>` : '';
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${gridSpan}</w:tcPr>${paragraphRuns(cell.paragraph.segments, { size: cell.paragraph.size || 22, spacingAfter: 0 })}</w:tc>`;
}

function leaderNotesCell(item: OrderItem): TableCellSpec {
  const leader = tableLeaderFor(item);
  const notes = tableNotesFor(item);
  const segments: TextRun[] = [];
  if (leader) segments.push({ text: leader, bold: shouldBoldLeader(item) });
  if (leader && notes) segments.push({ text: ' - ' });
  if (notes) segments.push({ text: notes });
  if (!segments.length) segments.push({ text: '' });
  return {
    paragraph: {
      segments,
      size: 22
    }
  };
}

function tableItemFor(item: OrderItem) {
  const itemType = String(item.itemType || '').trim();
  if (normalizeItemType(itemType) === ITEM_CALL_TO_WORSHIP) return 'Call to Worship';
  if (normalizeItemType(itemType) === ITEM_SECOND_SCRIPTURE) return 'Scripture Reading';
  return itemType;
}

function tableDetailFor(item: OrderItem) {
  return String(item.detail || '').trim();
}

function tableLeaderFor(item: OrderItem) {
  return String(item.leader || '').trim();
}

function tableNotesFor(item: OrderItem) {
  return String(item.notes || '').trim();
}

function shouldBoldDetail(item: OrderItem) {
  const value = normalizeItemType(item.itemType);
  return value.includes('song') || value.includes('sound');
}

function shouldBoldLeader(item: OrderItem) {
  const leader = String(item.leader || '').trim();
  if (leader) return true;
  return normalizeItemType(item.itemType).includes('leader');
}

function compactMetaLine(service: ServiceRecord) {
  const parts: string[] = [];
  if (service.type) parts.push(service.type);
  if (service.leader) parts.push(`Leader: ${service.leader}`);
  if (service.preacher) parts.push(`Preacher: ${service.preacher}`);
  return parts.join(' | ');
}

function formatLongDate(isoDate: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(isoDate || ''))) return String(isoDate || '');
  const [y, m, d] = isoDate.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'America/Chicago', 'EEEE, MMMM d, yyyy');
}

function paragraph(text: string, opts?: ParagraphOptions) {
  return paragraphRuns([{ text: String(text || ''), bold: opts?.bold, italic: opts?.italic }], opts);
}

function paragraphRuns(runs: TextRun[], opts?: ParagraphOptions) {
  const align = opts?.align && opts.align !== 'left' ? `<w:jc w:val="${opts.align}"/>` : '';
  const spacing = typeof opts?.spacingAfter === 'number' ? `<w:spacing w:after="${opts.spacingAfter}"/>` : '';
  const paragraphProps = (align || spacing) ? `<w:pPr>${spacing}${align}</w:pPr>` : '';
  const size = opts?.size || 22;
  const xml = runs.map(run => runXml(run, size)).join('');
  return `<w:p>${paragraphProps}${xml}</w:p>`;
}

function runXml(run: TextRun, defaultSize: number) {
  const text = String(run?.text || '');
  const lines = text.split('\n');
  const runProps = [
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>',
    run?.bold ? '<w:b/>' : '',
    run?.italic ? '<w:i/>' : '',
    `<w:sz w:val="${defaultSize}"/>`,
    `<w:szCs w:val="${defaultSize}"/>`
  ].join('');
  const textXml = lines.map((line, idx) => `${idx ? '<w:br/>' : ''}<w:t xml:space="preserve">${xmlEscape(line)}</w:t>`).join('');
  return `<w:r><w:rPr>${runProps}</w:rPr>${textXml}</w:r>`;
}

function pageBreak() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

function sectionProperties() {
  return '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
}

function xmlBlob(name: string, content: string) {
  return Utilities.newBlob(content, 'text/xml', name);
}

function xmlDocument(content: string) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${content}`;
}

function xmlEscape(value: string) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function contentTypesXml() {
  return xmlDocument('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
}

function rootRelsXml() {
  return xmlDocument('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
}

function documentRelsXml() {
  return xmlDocument('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
}
