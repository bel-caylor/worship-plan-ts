import { ORDER_OF_WORSHIP_EXPORT_FOLDER_URL, SERVICES_COL, SERVICES_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';
import { getOrder, type OrderItem } from './order';
import { esvPassage } from './services';

type ExportOrderOfWorshipDocInput = {
  serviceId?: string;
  folderUrl?: string;
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

  let tempDocId = '';
  try {
    const tempDoc = DocumentApp.create(`TMP ${docxName}`);
    tempDocId = tempDoc.getId();
    buildDocument(tempDoc, service, items, callReading, secondReading);
    tempDoc.saveAndClose();

    const exported = DriveApp.getFileById(tempDocId).getAs(DOCX_MIME_TYPE);
    exported.setName(nextAvailableFileName(folder, docxName));
    const saved = folder.createFile(exported);

    return {
      ok: true,
      name: saved.getName(),
      url: saved.getUrl(),
      folderUrl: folder.getUrl()
    };
  } finally {
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (_) { /* ignore */ }
    }
  }
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
  const url = `https://www.biblegateway.com/passage/?search=${encodeURIComponent(ref)}&version=LBLA`;
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; WorshipPlanExporter/1.0)'
    }
  });
  const html = response.getContentText();
  if (!html) return '';

  const match = html.match(/<div class="passage-text">[\s\S]*?<div class='passage-content[\s\S]*?<div class="version-LBLA[^"]*">([\s\S]*?)<a class="full-chap-link"/i);
  if (!match?.[1]) return '';

  let content = match[1];
  content = content.replace(/<h\d[\s\S]*?<\/h\d>/gi, '');
  content = content.replace(/<sup[\s\S]*?<\/sup>/gi, '');
  content = content.replace(/<span class="chapternum">[\s\S]*?<\/span>/gi, '');
  content = content.replace(/<span class="versenum">[\s\S]*?<\/span>/gi, '');
  content = content.replace(/<br\s*\/?>/gi, '\n');
  content = content.replace(/<\/p>/gi, '\n\n');
  content = content.replace(/<\/div>/gi, '\n\n');
  content = content.replace(/<[^>]+>/g, '');
  content = decodeHtmlEntities(content);
  return normalizePassageText(content);
}

function decodeHtmlEntities(value: string) {
  return String(value || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, '\'')
    .replace(/&rsquo;/g, '\'')
    .replace(/&lsquo;/g, '\'')
    .replace(/&ldquo;/g, '"')
    .replace(/&rdquo;/g, '"')
    .replace(/&ndash;/g, '–')
    .replace(/&mdash;/g, '—')
    .replace(/&hellip;/g, '...')
    .replace(/&aacute;/g, 'á')
    .replace(/&eacute;/g, 'é')
    .replace(/&iacute;/g, 'í')
    .replace(/&oacute;/g, 'ó')
    .replace(/&uacute;/g, 'ú')
    .replace(/&Aacute;/g, 'Á')
    .replace(/&Eacute;/g, 'É')
    .replace(/&Iacute;/g, 'Í')
    .replace(/&Oacute;/g, 'Ó')
    .replace(/&Uacute;/g, 'Ú')
    .replace(/&ntilde;/g, 'ñ')
    .replace(/&Ntilde;/g, 'Ñ')
    .replace(/&iexcl;/g, '¡')
    .replace(/&iquest;/g, '¿')
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
}

function normalizePassageText(value: string) {
  return String(value || '')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function buildDocxFileName(service: ServiceRecord) {
  const parts = String(service.date || '').split('-');
  const shortDate = parts.length === 3
    ? `${parts[0].slice(-2)}.${parts[1]}.${parts[2]}`
    : String(service.date || 'service');
  const suffix = sanitizeFileName(service.type || 'Service');
  return `${shortDate} Order of Worship (${suffix}).docx`;
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

function buildDocument(
  doc: GoogleAppsScript.Document.Document,
  service: ServiceRecord,
  items: OrderItem[],
  callReading: ReadingContent,
  secondReading: ReadingContent
) {
  const body = doc.getBody();
  clearContainer(body);

  const header = doc.getHeader() || doc.addHeader();
  clearContainer(header);
  const headerPara = header.appendParagraph(formatLongDate(service.date));
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerPara.setFontFamily('Arial');
  headerPara.setFontSize(11);

  const title = body.appendParagraph('Order of Worship');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setBold(true);

  const subtitle = body.appendParagraph(`${formatLongDate(service.date)}${service.time ? ` • ${service.time}` : ''}`);
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitle.setFontFamily('Arial');
  subtitle.setFontSize(11);

  if (service.leader || service.preacher) {
    const meta = body.appendParagraph(compactMetaLine(service));
    meta.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    meta.setFontFamily('Arial');
    meta.setFontSize(10);
  }

  body.appendParagraph('');
  appendOrderTable(body, items);
  body.appendPageBreak();
  appendReadingPage(body, service, callReading);
  body.appendPageBreak();
  appendReadingPage(body, service, secondReading);
}

function clearContainer(container: GoogleAppsScript.Document.Body | GoogleAppsScript.Document.HeaderSection) {
  for (let i = container.getNumChildren() - 1; i >= 0; i -= 1) {
    container.removeChild(container.getChild(i));
  }
}

function compactMetaLine(service: ServiceRecord) {
  const parts: string[] = [];
  if (service.type) parts.push(service.type);
  if (service.leader) parts.push(`Leader: ${service.leader}`);
  if (service.preacher) parts.push(`Preacher: ${service.preacher}`);
  return parts.join(' • ');
}

function formatLongDate(isoDate: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(isoDate || ''))) return String(isoDate || '');
  const [y, m, d] = isoDate.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'America/Chicago', 'EEEE, MMMM d, yyyy');
}

function appendOrderTable(body: GoogleAppsScript.Document.Body, items: OrderItem[]) {
  const rows = [
    ['LIGHTING / SOUND', 'AUDIO INPUT', 'EVENT'],
    ...items.map(item => [
      '',
      tableLeaderFor(item),
      tableEventFor(item)
    ])
  ];
  const table = body.appendTable(rows);
  for (let r = 0; r < table.getNumRows(); r += 1) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c += 1) {
      const cell = row.getCell(c);
      cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
      const text = cell.editAsText();
      text.setFontFamily('Arial');
      text.setFontSize(r === 0 ? 10 : 11);
      if (r === 0) text.setBold(true);
    }
  }
}

function tableLeaderFor(item: OrderItem) {
  const leader = String(item.leader || '').trim();
  if (leader) return leader;
  return isSongItem(item.itemType) ? 'Worship Team' : '';
}

function tableEventFor(item: OrderItem) {
  const itemType = String(item.itemType || '').trim();
  const detail = String(item.detail || '').trim();
  if (!itemType && !detail) return '';
  if (!detail) return itemType;
  if (normalizeItemType(itemType) === ITEM_CALL_TO_WORSHIP) return `Call to Worship – ${detail}`;
  if (normalizeItemType(itemType) === ITEM_SECOND_SCRIPTURE) return `Scripture Reading – ${detail}`;
  if (isSongItem(itemType)) return detail;
  return `${itemType} – ${detail}`;
}

function isSongItem(itemType: unknown) {
  const value = normalizeItemType(itemType);
  return value.includes('song') || value === 'opening song' || value === 'closing song';
}

function appendReadingPage(body: GoogleAppsScript.Document.Body, service: ServiceRecord, reading: ReadingContent) {
  const dateLine = body.appendParagraph(formatLongDate(service.date));
  dateLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  dateLine.setBold(true);
  dateLine.setFontFamily('Arial');
  dateLine.setFontSize(12);

  const title = body.appendParagraph(`${reading.heading}${reading.reader ? ` (${reading.reader})` : ''}`);
  title.setBold(true);
  title.setItalic(true);
  title.setFontFamily('Arial');
  title.setFontSize(20);

  body.appendParagraph('');

  const intro = body.appendParagraph(reading.intro);
  intro.setFontFamily('Arial');
  intro.setFontSize(16);

  body.appendParagraph('');

  if (reading.firstTranslation === 'ESV') {
    appendTranslationBlock(body, 'ESV', reading.englishText);
    body.appendParagraph('');
    appendTranslationBlock(body, 'LBLA', reading.spanishText);
    return;
  }

  appendTranslationBlock(body, 'LBLA', reading.spanishText);
  body.appendParagraph('');
  appendTranslationBlock(body, 'ESV', reading.englishText);
}

function appendTranslationBlock(body: GoogleAppsScript.Document.Body, label: 'ESV' | 'LBLA', text: string) {
  const labelPara = body.appendParagraph(`${label}:`);
  labelPara.setBold(true);
  labelPara.setFontFamily('Arial');
  labelPara.setFontSize(13);

  const blocks = String(text || '').split(/\n{2,}/).map(part => part.trim()).filter(Boolean);
  if (!blocks.length) {
    const empty = body.appendParagraph('');
    empty.setFontFamily('Arial');
    empty.setFontSize(13);
    return;
  }
  blocks.forEach((block) => {
    const para = body.appendParagraph(block);
    para.setFontFamily('Arial');
    para.setFontSize(13);
  });
}
