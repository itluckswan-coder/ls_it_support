/* Luck Swan Sales — minimal spreadsheet I/O: read .xlsx / .csv, write a simple .xlsx (no dependencies) */
(function (g) {
  'use strict';

  /* ---------- zip ---------- */
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

  async function inflateRaw(data) {
    if (typeof DecompressionStream === 'undefined') throw new Error('เบราว์เซอร์นี้อ่านไฟล์ Excel ไม่ได้ กรุณาบันทึกเป็น CSV แล้วลองใหม่');
    const ds = new DecompressionStream('deflate-raw');
    const out = new Response(new Blob([data]).stream().pipeThrough(ds));
    return new Uint8Array(await out.arrayBuffer());
  }

  async function unzip(buf) {
    const b = new Uint8Array(buf);
    let eocd = -1;
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (u32(b, i) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('ไฟล์ไม่ใช่ Excel (.xlsx) ที่ถูกต้อง');
    const n = u16(b, eocd + 10);
    let p = u32(b, eocd + 16);
    const files = {};
    for (let i = 0; i < n; i++) {
      if (u32(b, p) !== 0x02014b50) break;
      const method = u16(b, p + 10), csize = u32(b, p + 20), nlen = u16(b, p + 28), xlen = u16(b, p + 30), clen = u16(b, p + 32), off = u32(b, p + 42);
      const name = new TextDecoder().decode(b.subarray(p + 46, p + 46 + nlen));
      const dataStart = off + 30 + u16(b, off + 26) + u16(b, off + 28);
      const raw = b.subarray(dataStart, dataStart + csize);
      files[name] = { method, raw };
      p += 46 + nlen + xlen + clen;
    }
    return {
      has: name => !!files[name],
      names: () => Object.keys(files),
      async text(name) {
        const f = files[name]; if (!f) return null;
        const data = f.method === 0 ? f.raw : f.method === 8 ? await inflateRaw(f.raw) : null;
        if (!data) throw new Error('รูปแบบการบีบอัดไฟล์ไม่รองรับ');
        return new TextDecoder().decode(data);
      }
    };
  }

  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = d => { let c = 0xffffffff; for (let i = 0; i < d.length; i++) c = CRC[(c ^ d[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

  function zip(entries) { // stored (no compression) zip
    const enc = new TextEncoder(), parts = [], central = [];
    let off = 0;
    for (const [name, text] of entries) {
      const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nb.length, true);
      parts.push(new Uint8Array(h.buffer), nb, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), nb);
      off += 30 + nb.length + data.length;
    }
    const csize = central.reduce((a, x) => a + x.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, entries.length, true); e.setUint16(10, entries.length, true); e.setUint32(12, csize, true); e.setUint32(16, off, true);
    return new Blob([...parts, ...central, new Uint8Array(e.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ---------- xlsx ---------- */
  const colIdx = ref => { const m = /^([A-Z]+)/.exec(ref || ''); if (!m) return -1; let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const byTag = (node, tag) => Array.from(node.getElementsByTagNameNS('*', tag));

  async function readXlsx(buf) {
    const z = await unzip(buf);
    // first worksheet listed in the workbook
    let sheetPath = 'xl/worksheets/sheet1.xml';
    try {
      const wb = xml(await z.text('xl/workbook.xml')), rels = xml(await z.text('xl/_rels/workbook.xml.rels'));
      const first = byTag(wb, 'sheet')[0];
      const rid = first && (first.getAttribute('r:id') || first.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
      const rel = byTag(rels, 'Relationship').find(r => r.getAttribute('Id') === rid);
      if (rel) { const t = rel.getAttribute('Target').replace(/^\//, ''); sheetPath = t.startsWith('xl/') ? t : 'xl/' + t; }
    } catch (e) { /* fall back to sheet1 */ }
    if (!z.has(sheetPath)) sheetPath = z.names().find(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n));
    if (!sheetPath) throw new Error('ไม่พบแผ่นงานในไฟล์');
    const shared = [];
    const ss = await z.text('xl/sharedStrings.xml');
    if (ss) for (const si of byTag(xml(ss), 'si')) shared.push(byTag(si, 't').map(t => t.textContent).join(''));
    const doc = xml(await z.text(sheetPath));
    const rows = [];
    for (const r of byTag(doc, 'row')) {
      const idx = Number(r.getAttribute('r')) - 1, row = [];
      for (const c of byTag(r, 'c')) {
        const t = c.getAttribute('t'), ci = colIdx(c.getAttribute('r'));
        const v = byTag(c, 'v')[0], val = v ? v.textContent : '';
        let out;
        if (t === 's') out = shared[Number(val)] || '';
        else if (t === 'inlineStr') out = byTag(c, 't').map(x => x.textContent).join('');
        else if (t === 'str' || t === 'e') out = val;
        else if (t === 'b') out = val === '1' ? 'TRUE' : 'FALSE';
        else out = val === '' ? '' : Number(val);
        row[ci < 0 ? row.length : ci] = out;
      }
      rows[isNaN(idx) ? rows.length : idx] = row;
    }
    return Array.from(rows, r => Array.from(r || [], x => x == null ? '' : x));
  }

  function writeXlsx(rows, opts = {}) {
    const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const col = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
    const widths = opts.widths || [];
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
${widths.length ? '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>' : ''}<sheetData>${rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) =>
      typeof v === 'number' ? `<c r="${col(ci)}${ri + 1}"${ri === 0 ? ' s="1"' : ''}><v>${v}</v></c>` : `<c r="${col(ci)}${ri + 1}" t="inlineStr"${ri === 0 ? ' s="1"' : ' s="2"'}><is><t xml:space="preserve">${esc(v)}</t></is></c>`).join('')}</row>`).join('')}</sheetData></worksheet>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Tahoma"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Tahoma"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFA8262D"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    return zip([
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${esc(opts.sheetName || 'Sheet1')}" sheetId="1" r:id="rId1"/></sheets></workbook>`],
      ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml', styles],
      ['xl/worksheets/sheet1.xml', sheet]
    ]);
  }

  /* ---------- public ---------- */
  async function readFile(file) {
    const name = (file.name || '').toLowerCase();
    if (/\.xls$/.test(name)) throw new Error('ไฟล์ .xls รุ่นเก่ายังไม่รองรับ กรุณาบันทึกเป็น .xlsx หรือ .csv');
    if (/\.(xlsx|xlsm)$/.test(name)) return readXlsx(await file.arrayBuffer());
    if (/\.(csv|txt|tsv)$/.test(name) || !name.includes('.')) {
      let text = await file.text();
      text = text.replace(/^﻿/, '');
      if (/\.tsv$/.test(name) || (!text.includes(',') && text.includes('\t'))) text = text.split(/\r?\n/).map(l => l.split('\t').map(c => /[",]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c).join(',')).join('\n');
      return g.U.parseCSV(text);
    }
    throw new Error('รองรับเฉพาะไฟล์ .xlsx และ .csv');
  }

  g.Sheet = { readFile, readXlsx, writeXlsx };
})(window);
