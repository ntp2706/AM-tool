// Node >= 18. Chạy: npm install && npm start  → http://<IP-máy-chạy-server>:3000
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as cheerio from 'cheerio';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';

// Cert của gdt.gov.vn thường thiếu chuỗi trung gian → Node báo lỗi. Chỉ dùng cho local.
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
const TARGET = 'https://tracuunnt.gdt.gov.vn/tcnnt/mstdn.jsp';
const CONTRACT_TEMPLATE_IDS = {
  'web-hosting': [
    '1hycPqNuQQwfHRv-0_gXLAKDNI5SIzfv4',
    '1vwUAKGG7pGvb7oMozx_-qjcKtfDMKxYW',
  ],
  'ca-don-vi': ['1029H9MuXITcFi_1mauBTMSaxHiTh7xGr'],
  'ca-bhxh': ['1thLwiljnkJiqVo1UbSCkmO0OnDhbKOUp'],
  'combo-hkd': ['1KbsVIo77oZz7yJADObXuwEnPPqvopD7-'],
  'hkd-cu': ['1V1_MVb4R0E_5ttsEmhqAyM4SlCaGKcsJ'],
  pharmacy: ['12Eg6KsLOyOjZZRo4Y0HO_rTAJ9O0xfrU', '1k8O-FlEBDZJ-EaMR-y3Kd9zJ5LTmRjbs'],
  'ca-ca-nhan': [
    '1CFnb-4dcacA2Lqg_L5BG5uuQe8G8N7x2',
    '16e8lTtH52fGaF_xrkC5PDX_GWLI4kpnC',
  ],
  'hoadon-dien-tu': [
    '1NsQRv9EN_HIfQpmh6hsKgUoRdTKGHX9T',
    '1gRulvNA30WnEEkkHCFKLiT5kr0_6k5zz',
  ],
  nq66: [
    '1yeLcyEAGcdm62oXiapFIEK1Y95f_whsw',
    '1Rguo9Ikuv5KK8jqJxG4UQYdiXpSRuWQr',
  ],
};
const ECONTRACT_TEMPLATE_IDS = {
  cadonvi: '1Pb06tFY7BZ9Y_cp-HioITHti9SXh_Y2i',
  cacanhan: '1lRFtM7xzXimr7BrOGnXWGHEcp6a7Stxw',
  hoadondientu: '11lrYiWrsFHPQAAQ8vJxByG75oN6OYUe7',
};
const ECONTRACT_CHECKBOX_FIELDS = new Set([
  'capMoi', 'giaHan', 'TokenHSM', 'SmartCA', '1nam', '2nam', '3nam', 'coBan', 'nangCao',
]);
const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const TEMPLATE_FIELDS = [
  'soHopDong', 'ngayKy', 'thangKy', 'namKy', 'ngayThangNamKy',
  'tenDoanhNghiepA', 'diaChiA', 'maSoThueA', 'nguoiDaiDienA',
  'chucVuDaiDienA', 'soCccdDaiDienA', 'noiCapCccdA', 'ngayCapCccdA', 'loaiHinhDN', 'xaPhuong',
  'nguoiLienHeA', 'nguoiLienHeB', 'chucVuNguoiLienHeA', 'chucVuNguoiLienHeB', 'sdtLienHeA',
  'emailLienHeA', 'nguoiNghiemThuA', 'nguoiKyXacNhan', 'tenMien', 'goiCuoc',
  'thoiHan', 'soThangBangChu',
  'capMoi', 'giaHan', 'TokenHSM', 'SmartCA', '1nam', '2nam', '3nam', 'coBan', 'nangCao',
  'donGiaSauThue', 'donGiaTruocThue', 'cuocDVBoSung', 'soTB', 'thanhTienTruocThue',
  'soLuongHD', 'soHDBBNT', 'tongTienTruocThue', 'thueVAT', 'thanhTienSauThue',
  'tongGiaTriHopDong', 'soTienBangChu',
];
const TOKEN_TEMPLATE_FIELDS = [
  'stt', 'soTB', 'cuocDVBoSung', 'tongCuocDVBoSung', 'thueDVBoSung', 'cuocDVBoSungSauThue',
];
const CONTRACT_RUN_FORMAT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>';
const MAX_TEMPLATE_SIZE = 10 * 1024 * 1024;
const MAX_IMPORT_SIZE = 10 * 1024 * 1024;
const MAX_CA_ROWS = 5000;
const PORT = process.env.PORT || 3000;
const __dir = path.dirname(fileURLToPath(import.meta.url));
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36';

// Cookie jar (1 phiên - dùng cá nhân/local). Captcha gắn với JSESSIONID nên phải giữ cookie.
const jar = new Map();
const cookieHeader = () => [...jar].map(([k, v]) => `${k}=${v}`).join('; ');
const saveCookies = (res) => {
  for (const c of res.headers.getSetCookie?.() || []) {
    const [kv] = c.split(';'); const i = kv.indexOf('=');
    jar.set(kv.slice(0, i).trim(), kv.slice(i + 1).trim());
  }
};
async function upstream(url, opts = {}) {
  const res = await fetch(url, {
    redirect: 'follow', ...opts,
    headers: { 'User-Agent': UA, Cookie: cookieHeader(), Referer: TARGET, ...(opts.headers || {}) },
  });
  saveCookies(res);
  return res;
}

// Chỉ giữ: input[name=mst], input[name=captcha], img captcha, .subBtn (trong #module3Content)
const esc = (v) => String(v ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
function extract(html, baseUrl) {
  const $ = cheerio.load(html);
  let box = $('#module3Content');
  if (!box.length) box = $('[id*="module3" i]').first();
  if (!box.length) {
    const ids = $('[id]').map((i, e) => e.attribs.id).get().slice(0, 60);
    return { html: '', data: [], debug: { length: html.length, title: $('title').text().trim(), ids, head: html.slice(0, 600) } };
  }
  box.find('script').remove();
  const data = parseResults($, box);          // bắt kết quả từ response (không hiển thị)

  const mst = box.find('input[name="mst"]').first();
  if (!mst.length) return { html: '', data };  // response kết quả không chứa form → giữ form cũ ở client

  const form = mst.closest('form');
  const action = new URL(form.attr('action') || '', baseUrl).href;
  const method = (form.attr('method') || 'get').toLowerCase();

  // Các input khác giữ lại dưới dạng hidden để submit vẫn đúng, nhưng không hiển thị
  let hidden = '';
  form.find('input,select,textarea').each((_, el) => {
    const n = el.attribs.name, t = (el.attribs.type || '').toLowerCase();
    if (!n || n === 'mst' || n === 'captcha') return;
    if (['submit', 'button', 'image', 'reset', 'file'].includes(t)) return;
    if ((t === 'radio' || t === 'checkbox') && !('checked' in el.attribs)) return;
    hidden += `<input type="hidden" name="${esc(n)}" value="${esc($(el).val())}">`;
  });

  const img = box.find('img[src*="captcha.png"]').first();
  const imgSrc = img.attr('src') ? '/proxy/res?u=' + encodeURIComponent(new URL(img.attr('src'), baseUrl).href) : '';
  const btn = box.find('.subBtn').first();
  const label = btn.attr('value') || btn.text().trim() || 'Tra cứu';

  const out = `<form data-action="${esc(action)}" data-method="${method}">${hidden}
  <input type="text" id="tax" name="mst" maxlength="14" placeholder="Mã số thuế" autocomplete="off">
  <input type="text" name="captcha" maxlength="5" placeholder="Mã xác nhận" autocomplete="off">
  ${imgSrc ? `<img id="cap" src="${imgSrc}" data-src="${imgSrc}" alt="captcha" style="cursor:pointer">` : ''}
  <button type="submit" class="subBtn">${esc(label)}</button></form>`;
  return { html: out, data };
}

// Trích bảng kết quả → mảng object theo tiêu đề cột
function parseResults($, box) {
  const out = [];
  box.find('table').each((_, t) => {
    const rows = $(t).find('tr');
    let headers = null;
    rows.each((__, tr) => {
      const th = $(tr).find('th').map((i, e) => $(e).text().replace(/\s+/g, ' ').trim()).get();
      const td = $(tr).find('> td').map((i, e) => $(e).text().replace(/\s+/g, ' ').trim()).get();
      if (th.length > 2) { headers = th; return; }
      if (headers && td.length === headers.length && td.some(Boolean)) {
        out.push(Object.fromEntries(headers.map((h, i) => [h, td[i]])));
      }
    });
  });
  return out;
}

const send = (res, code, body, type = 'application/json; charset=utf-8', headers = {}) => {
  res.writeHead(code, { 'Content-Type': type, ...headers }); res.end(body);
};
function attachmentDisposition(filename) {
  const asciiFilename = filename.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
  const encodedFilename = encodeURIComponent(filename).replace(/['()*]/g, (character) => (
    `%${character.charCodeAt(0).toString(16).toUpperCase()}`
  ));
  return `attachment; filename="${asciiFilename}"; filename*=UTF-8''${encodedFilename}`;
}
const readBody = (req) => new Promise((r) => { let b = ''; req.on('data', (c) => (b += c)); req.on('end', () => r(b)); });
async function readBuffer(req, maxSize) {
  const chunks = [];
  let size = 0;
  let tooLarge = false;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > maxSize) tooLarge = true;
    else chunks.push(chunk);
  }
  if (tooLarge) throw new Error('Tệp import vượt quá giới hạn 10 MB.');
  return Buffer.concat(chunks, size);
}
function parseCaWorkbook(buffer) {
  const zip = new PizZip(buffer);
  const readXml = (name) => {
    const file = zip.file(name);
    if (!file) throw new Error(`File Excel thiếu thành phần ${name}.`);
    return cheerio.load(file.asText(), { xmlMode: true });
  };
  const workbook = readXml('xl/workbook.xml');
  const relationships = readXml('xl/_rels/workbook.xml.rels');
  const firstSheet = workbook('sheet').first();
  const relationshipId = firstSheet.attr('r:id');
  if (!relationshipId) throw new Error('File Excel không có sheet dữ liệu.');
  const target = relationships('Relationship')
    .filter((_, element) => relationships(element).attr('Id') === relationshipId)
    .first()
    .attr('Target');
  if (!target) throw new Error('Không tìm thấy nội dung sheet đầu tiên.');
  const sheetPath = target.startsWith('/')
    ? target.slice(1)
    : path.posix.normalize(path.posix.join('xl', target));
  if (!sheetPath.startsWith('xl/worksheets/')) throw new Error('Đường dẫn sheet Excel không hợp lệ.');
  const worksheet = readXml(sheetPath);
  const sharedStringsFile = zip.file('xl/sharedStrings.xml');
  const sharedStrings = sharedStringsFile
    ? cheerio.load(sharedStringsFile.asText(), { xmlMode: true })('si').toArray()
      .map((element) => cheerio.load(element, { xmlMode: true })('si').text())
    : [];
  const dateFormatIds = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  const dateStyleIds = new Set();
  if (zip.file('xl/styles.xml')) {
    const styles = readXml('xl/styles.xml');
    const customFormats = new Map(styles('numFmt').toArray().map((element) => {
      const item = styles(element);
      return [Number(item.attr('numFmtId')), item.attr('formatCode') || ''];
    }));
    styles('cellXfs > xf').each((index, element) => {
      const numFmtId = Number(styles(element).attr('numFmtId') || 0);
      const format = (customFormats.get(numFmtId) || '')
        .replace(/"[^"]*"/g, '')
        .replace(/\\./g, '')
        .replace(/\[[^\]]*\]/g, '');
      if (dateFormatIds.has(numFmtId) || /[dy]/i.test(format)) dateStyleIds.add(index);
    });
  }
  const columnIndex = (reference) => {
    const letters = /^([A-Z]+)/i.exec(reference || '')?.[1]?.toUpperCase();
    if (!letters) return 0;
    return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
  };
  const formatDate = (serial) => {
    if (!Number.isFinite(serial)) return '';
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
    return `${String(date.getUTCDate()).padStart(2, '0')}/${String(date.getUTCMonth() + 1).padStart(2, '0')}/${date.getUTCFullYear()}`;
  };
  const rowsInSheet = worksheet('sheetData > row');
  if (rowsInSheet.length > MAX_CA_ROWS + 1) {
    throw new Error(`File vượt quá giới hạn ${MAX_CA_ROWS} dòng dữ liệu.`);
  }
  const rows = [];
  const readCell = (row, column) => {
    const cell = row.children('c').toArray().find((element) => columnIndex(worksheet(element).attr('r')) === column);
    if (!cell) return '';
    const selected = worksheet(cell);
    const type = selected.attr('t');
    if (type === 'inlineStr') return selected.find('is').text().trim();
    const value = selected.children('v').first().text().trim();
    if (type === 's') return String(sharedStrings[Number(value)] ?? '').trim();
    if (type === 'b') return value === '1' ? 'TRUE' : 'FALSE';
    if (!value) return '';
    const styleId = Number(selected.attr('s') || 0);
    if (type === 'd') {
      const date = new Date(value);
      return Number.isNaN(date.getTime()) ? value : `${String(date.getUTCDate()).padStart(2, '0')}/${String(date.getUTCMonth() + 1).padStart(2, '0')}/${date.getUTCFullYear()}`;
    }
    if (dateStyleIds.has(styleId)) return formatDate(Number(value));
    return value;
  };
  rowsInSheet.each((index, element) => {
    const row = worksheet(element);
    const rowNumber = Number(row.attr('r') || index + 1);
    if (rowNumber < 2) return;
    const hoVaTen = readCell(row, 1);
    const soCCCDBM = readCell(row, 5);
    const chucVuBM = readCell(row, 6);
    const phuongXa = readCell(row, 7);
    const diaChi = readCell(row, 8);
    const ngayCapCCCDBM = readCell(row, 12);
    const noiCapCCCDBM = readCell(row, 13);
    if (![hoVaTen, soCCCDBM, chucVuBM, phuongXa, diaChi, ngayCapCCCDBM, noiCapCCCDBM].some(Boolean)) return;
    rows.push({
      hoVaTen,
      diaChiBM: [diaChi, phuongXa].filter(Boolean).join(', '),
      chucVuBM,
      soCCCDBM,
      quocTich: 'Việt Nam',
      ngayCapCCCDBM,
      noiCapCCCDBM,
    });
  });
  if (!rows.length) throw new Error('Không tìm thấy dữ liệu CA từ hàng 2 trở đi.');
  return rows;
}

function findWorkbookSheetPath(zip, sheetName) {
  const readXml = (name) => {
    const file = zip.file(name);
    if (!file) throw new Error(`File Excel thiếu thành phần ${name}.`);
    return cheerio.load(file.asText(), { xmlMode: true });
  };
  const workbook = readXml('xl/workbook.xml');
  const relationships = readXml('xl/_rels/workbook.xml.rels');
  const sheet = workbook('sheet').toArray().find((element) => workbook(element).attr('name') === sheetName);
  if (!sheet) throw new Error(`Không tìm thấy sheet "${sheetName}" trong mẫu eContract.`);
  const relationshipId = workbook(sheet).attr('r:id');
  const target = relationships('Relationship')
    .filter((_, element) => relationships(element).attr('Id') === relationshipId)
    .first()
    .attr('Target');
  if (!target) throw new Error(`Không tìm thấy nội dung sheet "${sheetName}".`);
  const sheetPath = target.startsWith('/')
    ? target.slice(1)
    : path.posix.normalize(path.posix.join('xl', target));
  if (!sheetPath.startsWith('xl/worksheets/')) throw new Error(`Đường dẫn sheet "${sheetName}" không hợp lệ.`);
  return sheetPath;
}

function replaceEcontractPlaceholders(zip, sheetName, cellReferences, values) {
  const sharedStringsFile = zip.file('xl/sharedStrings.xml');
  const sharedStrings = sharedStringsFile
    ? cheerio.load(sharedStringsFile.asText(), { xmlMode: true })('si').toArray()
      .map((element) => cheerio.load(element, { xmlMode: true })('si').text())
    : [];
  const sheetPath = findWorkbookSheetPath(zip, sheetName);
  const worksheetFile = zip.file(sheetPath);
  if (!worksheetFile) throw new Error(`File Excel thiếu nội dung sheet "${sheetName}".`);
  const worksheet = cheerio.load(worksheetFile.asText(), { xmlMode: true });
  const readCellText = (cell) => {
    const type = cell.attr('t');
    let text = type === 'inlineStr'
      ? cell.find('is').text()
      : cell.children('v').first().text();
    if (type === 's') text = sharedStrings[Number(text)] ?? '';
    return text;
  };
  for (const reference of cellReferences) {
    const cell = worksheet(`c[r="${reference}"]`).first();
    if (!cell.length) continue;
    const text = readCellText(cell);
    const checkboxMatch = /^\$\{([^}]+)\}$/.exec(text);
    const checkboxValue = checkboxMatch && ECONTRACT_CHECKBOX_FIELDS.has(checkboxMatch[1])
      ? values[checkboxMatch[1]]
      : null;
    const rowMatch = /^C(\d+)$/.exec(reference);
    const rowVariable = rowMatch && /^\$\{([^}]+)\}$/.exec(readCellText(worksheet(`c[r="B${rowMatch[1]}"]`).first()));
    const hasRowValue = rowVariable && Object.prototype.hasOwnProperty.call(values, rowVariable[1]);
    if (!text.includes('${') && !hasRowValue && text.trim()) continue;
    let replacement = text.includes('${')
      ? text.replace(/\$\{([^}]+)\}/g, (_, key) => String(values[key] ?? ''))
      : hasRowValue ? String(values[rowVariable[1]] ?? '') : text;
    if (!replacement.trim()) replacement = ' ';
    if (checkboxValue === '\u2612' || checkboxValue === '\u2610') {
      cell.empty().removeAttr('t').append(`<v>${checkboxValue === '\u2612' ? '1' : '0'}</v>`);
    } else {
      cell.empty().attr('t', 'inlineStr')
        .append(`<is><t xml:space="preserve">${replacement.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</t></is>`);
    }
  }
  zip.file(sheetPath, worksheet.xml());
}

function createEcontractWorkbook(buffer, values) {
  const zip = new PizZip(buffer);
  replaceEcontractPlaceholders(
    zip,
    'Thông tin hợp đồng',
    Array.from({ length: 94 }, (_, index) => `C${index + 2}`),
    values,
  );
  if (values.businessType === 'hoadondientu') {
    replaceEcontractPlaceholders(zip, 'Bảng #dv', ['B2', 'C2', 'D2', 'E2', 'F2', 'G2'], values);
  }
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function useTimesNewRoman(xml) {
  return xml.replace(/(<w:r(?=[\s>])[^>]*>)([\s\S]*?)<\/w:r>/g, (run, openingTag, content) => {
    const properties = /<w:rPr(?=[\s/>])(?:[^>]*\/>|[^>]*>[\s\S]*?<\/w:rPr>)/.exec(content);
    if (!properties) return `${openingTag}<w:rPr>${CONTRACT_RUN_FORMAT}</w:rPr>${content}</w:r>`;

    const updatedProperties = properties[0].startsWith('<w:rPr')
      && /\/>$/.test(properties[0])
      ? `<w:rPr>${CONTRACT_RUN_FORMAT}</w:rPr>`
      : properties[0]
        .replace(/<w:rFonts(?=[\s/>])[^>]*(?:\/>|>[\s\S]*?<\/w:rFonts>)/g, '')
        .replace(/<w:sz(?:Cs)?(?=[\s/>])[^>]*(?:\/>|>[\s\S]*?<\/w:sz(?:Cs)?>)/g, '')
        .replace(/(<w:rPr(?=[\s/>])[^>]*>)/, `$1${CONTRACT_RUN_FORMAT}`);
    return `${openingTag}${content.replace(properties[0], updatedProperties)}</w:r>`;
  });
}

function setParagraphLineSpacing(xml) {
  const setSpacing = (properties) => {
    const spacing = /<w:spacing(?=[\s/>])([^<>]*?)(?:\/\s*>|>[\s\S]*?<\/w:spacing>)/;
    const current = spacing.exec(properties);
    if (!current) {
      const afterSpacing = /<w:(?:ind|contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)(?=[\s/>])/;
      const followingProperty = afterSpacing.exec(properties);
      const insertionIndex = followingProperty ? followingProperty.index : properties.lastIndexOf('</w:pPr>');
      return `${properties.slice(0, insertionIndex)}<w:spacing w:line="317" w:lineRule="auto"/>${properties.slice(insertionIndex)}`;
    }
    const attributes = current[1]
      .replace(/\s+w:line="[^"]*"/g, '')
      .replace(/\s+w:lineRule="[^"]*"/g, '');
    return properties.replace(current[0], `<w:spacing${attributes} w:line="317" w:lineRule="auto"/>`);
  };
  return xml.replace(/<w:p(?=[\s/>])([^<>]*?)(\/\s*>|>([\s\S]*?)<\/w:p>)/g, (paragraph, attributes, ending, content = '') => {
    if (ending.startsWith('/')) return paragraph;
    const propertiesPattern = /<w:pPr(?=[\s/>])([^<>]*?)(\/\s*>|>([\s\S]*?)<\/w:pPr>)/;
    const properties = propertiesPattern.exec(content);
    if (!properties) {
      return `<w:p${attributes}><w:pPr><w:spacing w:line="317" w:lineRule="auto"/></w:pPr>${content}</w:p>`;
    }
    const updatedProperties = properties[2].startsWith('/')
      ? `<w:pPr${properties[1]}><w:spacing w:line="317" w:lineRule="auto"/></w:pPr>`
      : setSpacing(properties[0]);
    return `<w:p${attributes}>${content.replace(properties[0], updatedProperties)}</w:p>`;
  });
}

function setSectionPageLayout(xml) {
  const replaceElements = (source, name, transform) => {
    const pattern = new RegExp(`<w:${name}(?=[\\s/>])([^<>]*?)(\\/\\s*>|>([\\s\\S]*?)<\\/w:${name}>)`, 'g');
    return source.replace(pattern, (match, attributes, ending, content = '') =>
      transform({ attributes, content }));
  };
  const setPropertyAttributes = (element, name, values, insertionPattern) => {
    const pattern = new RegExp(`<w:${name}(?=[\\s/>])([^<>]*?)(\\/\\s*>|>([\\s\\S]*?)<\\/w:${name}>)`);
    const existing = pattern.exec(element);
    if (existing) {
      let attributes = existing[1];
      for (const [attribute, value] of Object.entries(values)) {
        attributes = attributes.replace(new RegExp(`\\s+w:${attribute}="[^"]*"`, 'g'), '');
        attributes += ` w:${attribute}="${value}"`;
      }
      return element.replace(existing[0], `<w:${name}${attributes}/>`); 
    }
    const insertion = insertionPattern.exec(element);
    const index = insertion ? insertion.index : element.lastIndexOf('</w:sectPr>');
    const property = `<w:${name}${Object.entries(values).map(([attribute, value]) => ` w:${attribute}="${value}"`).join('')}/>`;
    return `${element.slice(0, index)}${property}${element.slice(index)}`;
  };
  return replaceElements(xml, 'sectPr', ({ attributes, content }) => {
    let section = `<w:sectPr${attributes}>${content}</w:sectPr>`;
    section = setPropertyAttributes(
      section,
      'pgSz',
      { w: 11906, h: 16838, orient: 'portrait' },
      /<w:(?:pgMar|paperSrc|pgBorders|lnNumType|pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)(?=[\s/>])/,
    );
    return setPropertyAttributes(
      section,
      'pgMar',
      { top: 850, bottom: 850, left: 1417, right: 1020 },
      /<w:(?:paperSrc|pgBorders|lnNumType|pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)(?=[\s/>])/,
    );
  });
}

function setBlackTableBordersToHalfPoint(xml) {
  const replaceElements = (source, name, transform) => {
    const pattern = new RegExp(`<w:${name}(?=[\\s/>])([^<>]*?)(\\/\\s*>|>([\\s\\S]*?)<\\/w:${name}>)`, 'g');
    return source.replace(pattern, (match, attributes, ending, content = '') =>
      transform({ attributes, content, selfClosing: ending.startsWith('/') }));
  };
  const narrowBlackBorders = (borderXml) => {
    const borderName = /^<w:(tblBorders|tcBorders)/.exec(borderXml)?.[1];
    if (!borderName) return borderXml;
    return replaceElements(borderXml, borderName, ({ attributes, content }) => {
      const updatedContent = content.replace(/<w:(top|left|bottom|right|insideH|insideV)(?=[\s/>])([^>]*?)(\/?)>/g, (edge, name, attrs, selfClosingEdge) => {
        const color = /\bw:color="([^"]+)"/i.exec(attrs)?.[1]?.toLowerCase();
        const themeColor = /\bw:themeColor="([^"]+)"/i.exec(attrs)?.[1]?.toLowerCase();
        const value = /\bw:val="([^"]+)"/i.exec(attrs)?.[1]?.toLowerCase();
        const isBlack = ['000000', '000', 'auto'].includes(color) || (!color && !themeColor);
        if (!isBlack || ['nil', 'none'].includes(value)) return edge;
        const updatedAttributes = attrs.replace(/\s+w:sz="[^"]*"/i, '');
        return `<w:${name}${updatedAttributes} w:sz="4"${selfClosingEdge}>`;
      });
      return `<w:${borderName}${attributes}>${updatedContent}</w:${borderName}>`;
    });
  };
  let output = replaceElements(xml, 'tblBorders', ({ attributes, content }) =>
    narrowBlackBorders(`<w:tblBorders${attributes}>${content}</w:tblBorders>`));
  output = replaceElements(output, 'tcBorders', ({ attributes, content }) =>
    narrowBlackBorders(`<w:tcBorders${attributes}>${content}</w:tcBorders>`));
  return output;
}

http.createServer(async (req, res) => {
  const u = new URL(req.url, `http://${req.headers.host}`);
  try {
    if (u.pathname === '/') return send(res, 200, fs.readFileSync(path.join(__dir, 'public/index.html')), 'text/html; charset=utf-8', { 'Cache-Control': 'no-store' });
    if (u.pathname === '/extract.html') return send(res, 200, fs.readFileSync(path.join(__dir, 'public/extract.html')), 'text/html; charset=utf-8');
    if (u.pathname === '/sotien.gs') return send(res, 200, fs.readFileSync(path.join(__dir, 'public/sotien.gs')), 'text/javascript; charset=utf-8');
    if (u.pathname === '/img/stamp.png') return send(res, 200, fs.readFileSync(path.join(__dir, 'img/stamp.png')), 'image/png');

    if (u.pathname === '/api/import-ca' && req.method === 'POST') {
      const buffer = await readBuffer(req, MAX_IMPORT_SIZE);
      if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
        return send(res, 400, JSON.stringify({ error: 'File không đúng định dạng .xlsx.' }));
      }
      let rows;
      try {
        rows = await parseCaWorkbook(buffer);
      } catch (error) {
        return send(res, 400, JSON.stringify({ error: `Không đọc được file Excel: ${error.message}` }));
      }
      return send(res, 200, JSON.stringify({ rows }));
    }

    if (u.pathname === '/api/export-econtract' && req.method === 'POST') {
      const payload = JSON.parse(await readBody(req));
      const templateId = ECONTRACT_TEMPLATE_IDS[payload.businessType];
      if (typeof templateId !== 'string') {
        return send(res, 400, JSON.stringify({ error: 'Loại hình không hỗ trợ xuất Excel eContract.' }));
      }
      const templateUrl = `https://docs.google.com/spreadsheets/d/${templateId}/export?format=xlsx`;
      const templateResponse = await fetch(templateUrl, { redirect: 'follow' });
      if (!templateResponse.ok) throw new Error(`Không tải được mẫu Excel eContract từ Google Drive (HTTP ${templateResponse.status}).`);
      const contentLength = Number(templateResponse.headers.get('content-length') || 0);
      if (contentLength > MAX_TEMPLATE_SIZE) throw new Error('Tệp mẫu Excel eContract vượt quá giới hạn 10 MB.');
      const template = Buffer.from(await templateResponse.arrayBuffer());
      if (template.length > MAX_TEMPLATE_SIZE) throw new Error('Tệp mẫu Excel eContract vượt quá giới hạn 10 MB.');
      if (template[0] !== 0x50 || template[1] !== 0x4b) {
        throw new Error('Tệp Google Drive không phải mẫu Excel eContract hợp lệ hoặc không có quyền tải xuống.');
      }

      const output = createEcontractWorkbook(template, {
        ...payload,
        dienThoaiA: payload.sdtLienHeA,
        faxA: '',
        soTaiKhoanA: '',
        diaChiTaiKhoanA: '',
        soGiayUyQuyenA: '',
        ngayGiayUyQuyenA: '',
        donViUyQuyenA: '',
      });
      const filename = String(payload.filename || 'eContract.xlsx')
        .replace(/[<>:"\\|?*\x00-\x1f]/g, '_')
        .replace(/\//g, '_');
      res.writeHead(200, {
        'Content-Type': XLSX_MIME_TYPE,
        'Content-Disposition': attachmentDisposition(filename),
        'Content-Length': output.length,
        'X-Content-Type-Options': 'nosniff',
      });
      return res.end(output);
    }

    if (u.pathname === '/api/export-web-hosting' && req.method === 'POST') {
      const payload = JSON.parse(await readBody(req));
      const { templateIndex, templateType = 'web-hosting' } = payload;
      const templateIds = CONTRACT_TEMPLATE_IDS[templateType];
      if (!templateIds || !Number.isInteger(templateIndex) || !templateIds[templateIndex]) {
        return send(res, 400, JSON.stringify({ error: 'Mẫu DOCX không hợp lệ.' }));
      }
      const templateId = templateIds[templateIndex];
      const templateUrl = `https://drive.usercontent.google.com/download?id=${templateId}&export=download&confirm=t`;
      const templateResponse = await fetch(templateUrl, { redirect: 'follow' });
      if (!templateResponse.ok) throw new Error(`Không tải được mẫu DOCX từ Google Drive (HTTP ${templateResponse.status}).`);
      const contentLength = Number(templateResponse.headers.get('content-length') || 0);
      if (contentLength > MAX_TEMPLATE_SIZE) throw new Error('Tệp mẫu DOCX vượt quá giới hạn 10 MB.');
      const template = Buffer.from(await templateResponse.arrayBuffer());
      if (template.length > MAX_TEMPLATE_SIZE) throw new Error('Tệp mẫu DOCX vượt quá giới hạn 10 MB.');
      if (template[0] !== 0x50 || template[1] !== 0x4b) {
        throw new Error('Tệp Google Drive không phải DOCX hợp lệ hoặc không có quyền tải xuống.');
      }

      const zip = new PizZip(template);
      const isCaPersonalTemplate = templateType === 'ca-ca-nhan' && templateIndex === 1;
      const supportsTokenRows = ['ca-don-vi', 'ca-bhxh', 'ca-ca-nhan'].includes(templateType);
      if (supportsTokenRows) {
        for (const [name, file] of Object.entries(zip.files)) {
          if (!name.startsWith('word/') || !name.endsWith('.xml') || file.dir) continue;
          const originalXml = file.asText();
          let xml = originalXml.replace(/\{#token\}/g, '${#token}${stt}').replace(/\{\/token\}/g, '${/token}');
          if (isCaPersonalTemplate) {
            xml = xml.replace(/\{#STT\}/g, '${#STT}${STT}').replace(/\{\/STT\}/g, '${/STT}');
          }
          if (xml !== originalXml) zip.file(name, xml);
        }
      }
      const data = Object.fromEntries(TEMPLATE_FIELDS.map((key) => [key, String(payload[key] ?? '')]));
      if (supportsTokenRows) data.token = Array.isArray(payload.token)
        ? payload.token.map((row) => Object.fromEntries(
          TOKEN_TEMPLATE_FIELDS.map((key) => [key, String(row?.[key] ?? '')]),
        ))
        : [];
      if (isCaPersonalTemplate) {
        data.STT = Array.isArray(payload.caRows) ? payload.caRows.map((row, index) => ({
          STT: String(index + 1),
          hoVaTen: String(row?.hoVaTen ?? ''),
          diaChiBM: String(row?.diaChiBM ?? ''),
          chucVuBM: String(row?.chucVuBM ?? ''),
          soCCCDBM: String(row?.soCCCDBM ?? ''),
          quocTich: String(row?.quocTich ?? ''),
          ngayCapCCCDBM: String(row?.ngayCapCCCDBM ?? ''),
          noiCapCCCDBM: String(row?.noiCapCCCDBM ?? ''),
        })) : [];
      }
      const doc = new Docxtemplater(zip, {
        delimiters: { start: '${', end: '}' },
        nullGetter: () => '',
      });
      doc.render(data);
      const renderedZip = doc.getZip();
      for (const [name, file] of Object.entries(renderedZip.files)) {
        if (!name.startsWith('word/') || !name.endsWith('.xml') || file.dir) continue;
        let xml = useTimesNewRoman(file.asText());
        if (payload.pdfSource === true) {
          xml = setParagraphLineSpacing(xml);
          xml = setSectionPageLayout(xml);
          xml = setBlackTableBordersToHalfPoint(xml);
        }
        renderedZip.file(name, xml);
      }
      const output = renderedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
      const filename = String(payload.filename || `${templateType}-${templateIndex + 1}.docx`)
        .replace(/[<>:"\\|?*\x00-\x1f]/g, '_')
        .replace(/\//g, '_');
      res.writeHead(200, {
        'Content-Type': DOCX_MIME_TYPE,
        'Content-Disposition': attachmentDisposition(filename),
        'Content-Length': output.length,
        'X-Content-Type-Options': 'nosniff',
      });
      return res.end(output);
    }

    if (u.pathname === '/proxy/page') {
      jar.clear();
      const r = await upstream(TARGET);
      const out = extract(await r.text(), TARGET);
      if (out.debug) out.debug.status = r.status;
      return send(res, 200, JSON.stringify(out));
    }
    if (u.pathname === '/proxy/submit' && req.method === 'POST') {
      const { action, method, fields } = JSON.parse(await readBody(req));
      const body = new URLSearchParams(fields);
      const r = method === 'post'
        ? await upstream(action, { method: 'POST', body, headers: { 'Content-Type': 'application/x-www-form-urlencoded' } })
        : await upstream(action + (action.includes('?') ? '&' : '?') + body);
      const out = extract(await r.text(), r.url || action);
      console.log(`[submit] ${out.data.length} dòng kết quả`);
      return send(res, 200, JSON.stringify(out));
    }
    if (u.pathname === '/proxy/res') {
      const r = await upstream(u.searchParams.get('u'));
      res.writeHead(200, { 'Content-Type': r.headers.get('content-type') || 'image/jpeg', 'Cache-Control': 'no-store' });
      return res.end(Buffer.from(await r.arrayBuffer()));
    }
    send(res, 404, '{"error":"not found"}');
  } catch (e) { send(res, 500, JSON.stringify({ error: String(e) })); }
}).listen(PORT, '0.0.0.0', () => {
  const addresses = Object.values(os.networkInterfaces())
    .flat()
    .filter((info) => info && info.family === 'IPv4' && !info.internal)
    .map((info) => info.address);

  if (addresses.length) {
    console.log('Truy cập từ các máy cùng mạng bằng:');
    for (const address of addresses) console.log(`  http://${address}:${PORT}`);
  } else {
    console.log(`Server đang lắng nghe trên cổng ${PORT}, nhưng không tìm thấy IP LAN.`);
  }
});
