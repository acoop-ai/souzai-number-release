// xlsx を書く部品（外の部品なし・zip は縮めない stored・CRC32 は自前）。畜産の order-book.js の型に、書式・字の飾り・タブの色・印刷の設定・絵を足した。
//   入れる物はセルの格子（sheet-model.js）だけ。式として書くのは格子で f を持つセルだけ（文字の = は式に化けない）。
(function (root) {
  'use strict';
  const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(bytes) { let c = 0xFFFFFFFF; for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function dosTime(d) { return { time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2), date: ((Math.max(1980, d.getFullYear()) - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate() }; }
  /** files: [{ name, text | bytes }] → zip の中身 */
  function zipStored(files, when) {
    const enc = new TextEncoder(); const t = dosTime(when || new Date());
    const parts = []; const central = []; let offset = 0;
    files.forEach((f) => {
      const name = enc.encode(f.name); const data = f.bytes ? f.bytes : enc.encode(f.text); const crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true); local.setUint16(8, 0, true);
      local.setUint16(10, t.time, true); local.setUint16(12, t.date, true); local.setUint32(14, crc, true); local.setUint32(18, data.length, true); local.setUint32(22, data.length, true); local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true); cd.setUint16(10, 0, true);
      cd.setUint16(12, t.time, true); cd.setUint16(14, t.date, true); cd.setUint32(16, crc, true); cd.setUint32(20, data.length, true); cd.setUint32(24, data.length, true); cd.setUint16(28, name.length, true); cd.setUint32(42, offset, true);
      parts.push(new Uint8Array(local.buffer), name, data); central.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + data.length;
    });
    const cdSize = central.reduce((n, p) => n + p.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = parts.concat(central, [new Uint8Array(end.buffer)]);
    const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0)); let at = 0;
    all.forEach((p) => { out.set(p, at); at += p.length; });
    return out;
  }
  const xml = (v) => String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colName = (c) => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - m - 1) / 26); } return s; };
  const BUILTIN_FMT = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '@': 49 };

  /** 書式の表（同じ物は1つに）。cellStyle(書式の行, 字の色, 塗り) → cellXfs の番号 */
  function StyleTable(stdFont, stdSize) {
    const fonts = ['<font><sz val="' + stdSize + '"/><name val="' + xml(stdFont) + '"/><family val="3"/><charset val="128"/></font>'];
    const fontIdx = new Map([[fonts[0], 0]]);
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const fillIdx = new Map([[fills[0], 0], [fills[1], 1]]);
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    const borderIdx = new Map([[borders[0], 0]]);
    const fmts = []; const fmtIdx = new Map();
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const xfIdx = new Map([[xfs[0], 0]]);
    const memo = new Map();
    const intern = (list, map, s) => { let i = map.get(s); if (i === undefined) { i = list.length; list.push(s); map.set(s, i); } return i; };
    function cellStyle(st, key, fc, fl) {
      if (!st) return 0;
      const memoKey = key + '|' + (fc === undefined ? '~' : fc) + '|' + (fl === undefined ? '~' : fl);
      const hit = memo.get(memoKey); if (hit !== undefined) return hit;
      const color = fc !== undefined ? fc : (st['字の色'] || '');
      const fill = fl !== undefined ? fl : (st['塗り'] || '');
      const fontXml = '<font>' + (st['太字'] === '○' ? '<b/>' : '') + (st['斜体'] === '○' ? '<i/>' : '') + (st['大きさ'] ? '<sz val="' + st['大きさ'] + '"/>' : '') + (color ? '<color rgb="FF' + color + '"/>' : '') + (st['字'] ? '<name val="' + xml(st['字']) + '"/>' : '') + '<family val="3"/><charset val="128"/></font>';
      const fontId = intern(fonts, fontIdx, fontXml);
      const fillId = fill ? intern(fills, fillIdx, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + fill + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
      const side = (name, v) => (v ? '<' + name + ' style="' + v + '"><color auto="1"/></' + name + '>' : '<' + name + '/>');
      const borderId = intern(borders, borderIdx, '<border>' + side('left', st['罫線（左）']) + side('right', st['罫線（右）']) + side('top', st['罫線（上）']) + side('bottom', st['罫線（下）']) + '<diagonal/></border>');
      const code = st['表示形式'] || 'General';
      let numFmtId = BUILTIN_FMT[code];
      if (numFmtId === undefined) { numFmtId = fmtIdx.get(code); if (numFmtId === undefined) { numFmtId = 164 + fmts.length; fmts.push('<numFmt numFmtId="' + numFmtId + '" formatCode="' + xml(code) + '"/>'); fmtIdx.set(code, numFmtId); } }
      const al = (st['横の位置'] ? ' horizontal="' + st['横の位置'] + '"' : '') + (st['縦の位置'] ? ' vertical="' + st['縦の位置'] + '"' : '') + (st['折り返し'] === '○' ? ' wrapText="1"' : '') + (st['縮小して全体を表示'] === '○' ? ' shrinkToFit="1"' : '');
      const xf = '<xf numFmtId="' + numFmtId + '" fontId="' + fontId + '" fillId="' + fillId + '" borderId="' + borderId + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"' + (al ? ' applyAlignment="1"><alignment' + al + '/></xf>' : '/>');
      const idx = intern(xfs, xfIdx, xf);
      memo.set(memoKey, idx);
      return idx;
    }
    function toXml() {
      return XML_HEAD + '<styleSheet xmlns="' + MAIN_NS + '">' + (fmts.length ? '<numFmts count="' + fmts.length + '">' + fmts.join('') + '</numFmts>' : '')
        + '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills><borders count="' + borders.length + '">' + borders.join('') + '</borders>'
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }
    return { cellStyle, toXml };
  }

  /** 字の飾りの1区切り（rPr） */
  function runProps(p) {
    return '<rPr>' + (p.b ? '<b/>' : '') + (p.i ? '<i/>' : '') + (p.u ? '<u/>' : '') + (p.s ? '<sz val="' + p.s + '"/>' : '') + (p.c ? '<color rgb="FF' + p.c + '"/>' : '') + (p.f ? '<rFont val="' + xml(p.f) + '"/>' : '') + '<family val="3"/><charset val="128"/></rPr>';
  }
  /** 共有の文字（同じ文字は1つに。字の飾りつきは別） */
  function SharedStrings() {
    const items = []; const idx = new Map(); let total = 0;
    function add(text, runs) {
      total++;
      let body;
      if (runs && runs.length) {
        const s = String(text); let at = 0; body = '';
        runs.slice().sort((a, b) => a[0] - b[0]).forEach((r) => {
          const from = Math.max(at, Math.min(r[0], s.length)), to = Math.max(from, Math.min(r[1], s.length));
          if (from > at) body += '<r><t xml:space="preserve">' + xml(s.slice(at, from)) + '</t></r>';
          body += '<r>' + runProps(r[2] || {}) + '<t xml:space="preserve">' + xml(s.slice(from, to)) + '</t></r>';
          at = to;
        });
        if (at < s.length) body += '<r><t xml:space="preserve">' + xml(s.slice(at)) + '</t></r>';
      } else body = '<t xml:space="preserve">' + xml(text) + '</t>';
      let i = idx.get(body);
      if (i === undefined) { i = items.length; items.push(body); idx.set(body, i); }
      return i;
    }
    function toXml() { return XML_HEAD + '<sst xmlns="' + MAIN_NS + '" count="' + total + '" uniqueCount="' + items.length + '">' + items.map((b) => '<si>' + b + '</si>').join('') + '</sst>'; }
    return { add, toXml };
  }

  function sheetXml(sheet, index, book, st, sst, hasDrawing) {
    const styles = book.styles;
    const rowsMap = new Map();
    sheet.grid.cells.forEach((cell) => { if (!rowsMap.has(cell.r)) rowsMap.set(cell.r, []); rowsMap.get(cell.r).push(cell); });
    Object.keys(sheet.rowHeights).forEach((r) => { if (!rowsMap.has(Number(r))) rowsMap.set(Number(r), []); });
    (sheet.hiddenRows || []).forEach((r) => { if (!rowsMap.has(r)) rowsMap.set(r, []); });
    const hidden = new Set(sheet.hiddenRows || []);
    const rowNos = Array.from(rowsMap.keys()).sort((a, b) => a - b);
    const parts = [];
    rowNos.forEach((r) => {
      const cells = rowsMap.get(r).sort((a, b) => a.c - b.c);
      const h = sheet.rowHeights[r];
      let cx = '';
      cells.forEach((c) => {
        const sIdx = c.s ? st.cellStyle(styles[c.s], c.s, c.fc, c.fl) : 0;
        const ref = colName(c.c) + r;
        const sAttr = sIdx ? ' s="' + sIdx + '"' : '';
        if (c.f !== undefined) {
          const cached = typeof c.cached === 'number' ? '<v>' + c.cached + '</v>' : '';
          const isStr = c.cachedText !== undefined && !cached;
          cx += '<c r="' + ref + '"' + sAttr + (isStr ? ' t="str"' : '') + '><f>' + xml(c.f) + '</f>' + (isStr ? '<v></v>' : cached) + '</c>';
        } else if (c.v === undefined) {
          if (sIdx) cx += '<c r="' + ref + '"' + sAttr + '/>';
        } else if (typeof c.v === 'number') {
          cx += '<c r="' + ref + '"' + sAttr + '><v>' + c.v + '</v></c>';
        } else if (typeof c.v === 'boolean') {
          cx += '<c r="' + ref + '"' + sAttr + ' t="b"><v>' + (c.v ? 1 : 0) + '</v></c>';
        } else {
          cx += '<c r="' + ref + '"' + sAttr + ' t="s"><v>' + sst.add(c.v, c.runs) + '</v></c>';
        }
      });
      parts.push('<row r="' + r + '"' + (h !== undefined ? ' ht="' + h + '" customHeight="1"' : '') + (hidden.has(r) ? ' hidden="1"' : '') + '>' + cx + '</row>');
    });
    const cols = sheet.cols.slice().sort((a, b) => a.c - b.c).filter((c) => c.w !== null || c.hidden).map((c) => '<col min="' + c.c + '" max="' + c.c + '"' + (c.w !== null ? ' width="' + c.w + '" customWidth="1"' : ' width="8.43"') + (c.hidden ? ' hidden="1"' : '') + '/>').join('');
    const pg = sheet.page;
    const m = pg.margins;
    const hf = (o) => (o.l ? '&L' + o.l : '') + (o.c ? '&C' + o.c : '') + (o.r ? '&R' + o.r : '');
    const head = hf(pg.header), foot = hf(pg.footer);
    return XML_HEAD + '<worksheet xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '">'
      + '<sheetPr>' + (sheet.tabColor ? '<tabColor rgb="FF' + sheet.tabColor + '"/>' : '') + '</sheetPr>'
      + '<dimension ref="A1:' + colName(sheet.lastCol) + sheet.lastRow + '"/>'
      + '<sheetViews><sheetView workbookViewId="0"' + (pg.zoom ? ' zoomScale="' + pg.zoom + '"' : '') + (index === 0 ? ' tabSelected="1"' : '') + '/></sheetViews>'
      + '<sheetFormatPr defaultRowHeight="' + (pg.defaultRowHeight || 15) + '"' + (pg.defaultRowHeight ? ' customHeight="1"' : '') + '/>'
      + (cols ? '<cols>' + cols + '</cols>' : '') + '<sheetData>' + parts.join('') + '</sheetData>'
      + (sheet.merges.length ? '<mergeCells count="' + sheet.merges.length + '">' + sheet.merges.map((x) => '<mergeCell ref="' + x + '"/>').join('') + '</mergeCells>' : '')
      + ((pg.hCenter || pg.vCenter) ? '<printOptions' + (pg.hCenter ? ' horizontalCentered="1"' : '') + (pg.vCenter ? ' verticalCentered="1"' : '') + '/>' : '')
      + '<pageMargins left="' + (m.left ?? 0.7) + '" right="' + (m.right ?? 0.7) + '" top="' + (m.top ?? 0.75) + '" bottom="' + (m.bottom ?? 0.75) + '" header="' + (m.header ?? 0.3) + '" footer="' + (m.footer ?? 0.3) + '"/>'
      + '<pageSetup' + (pg.paper ? ' paperSize="' + pg.paper + '"' : '') + (pg.scale ? ' scale="' + pg.scale + '"' : '') + ' orientation="' + pg.orientation + '"/>'
      + ((head || foot) ? '<headerFooter>' + (head ? '<oddHeader>' + xml(head) + '</oddHeader>' : '') + (foot ? '<oddFooter>' + xml(foot) + '</oddFooter>' : '') + '</headerFooter>' : '')
      + (hasDrawing ? '<drawing r:id="rId1"/>' : '')
      + '</worksheet>';
  }
  function base64Bytes(s) {
    if (typeof atob === 'function') { const b = atob(s); const out = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) out[i] = b.charCodeAt(i); return out; }
    return new Uint8Array(Buffer.from(s, 'base64'));
  }

  /** book（sheet-model.js）→ xlsx の中身（Uint8Array）。when を同じにすれば、同じ入力から同じ中身 */
  function writeBook(book, when) {
    const stdShape = book.shape[Object.keys(book.shape)[0]] || {};
    const st = StyleTable(stdShape['標準の字'] || 'ＭＳ Ｐゴシック', stdShape['標準の字の大きさ'] || 11);
    const sst = SharedStrings();
    const files = [];
    const sheets = book.sheets;
    const types = [];
    let picNo = 0; const media = [];
    const sheetFiles = sheets.map((sheet, i) => {
      const pics = (sheet.pictures || []).filter((p) => p.data);
      return { xml: sheetXml(sheet, i, book, st, sst, pics.length > 0), pics };
    });
    types.push('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
    sheets.forEach((s, i) => types.push('<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'));
    const quote = (n) => "'" + String(n).replace(/'/g, "''") + "'";
    const printName = (sheet, i) => { const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(sheet.page.printArea); return '<definedName name="_xlnm.Print_Area" localSheetId="' + i + '">' + xml(quote(sheet.name)) + '!$' + m[1] + '$' + m[2] + ':$' + m[3] + '$' + m[4] + '</definedName>'; };
    sheetFiles.forEach((f, i) => {
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', text: f.xml });
      if (f.pics.length) {
        const dn = ++picNo;
        files.push({ name: 'xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL_NS + '/drawing" Target="../drawings/drawing' + dn + '.xml"/></Relationships>' });
        let anchors = '', rels = '';
        f.pics.forEach((p, k) => {
          const mediaName = 'image' + (media.length + 1) + '.png';
          media.push({ name: 'xl/media/' + mediaName, bytes: base64Bytes(p.data) });
          const a = p.anchor;
          const mark = (tag, x) => '<xdr:' + tag + '><xdr:col>' + x.col + '</xdr:col><xdr:colOff>' + x.colOff + '</xdr:colOff><xdr:row>' + x.row + '</xdr:row><xdr:rowOff>' + x.rowOff + '</xdr:rowOff></xdr:' + tag + '>';
          anchors += '<xdr:twoCellAnchor>' + mark('from', a.from) + mark('to', a.to) + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (k + 2) + '" name="図 ' + (k + 1) + '"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="' + REL_NS + '" r:embed="rId' + (k + 1) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';
          rels += '<Relationship Id="rId' + (k + 1) + '" Type="' + REL_NS + '/image" Target="../media/' + mediaName + '"/>';
        });
        files.push({ name: 'xl/drawings/drawing' + dn + '.xml', text: XML_HEAD + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + anchors + '</xdr:wsDr>' });
        files.push({ name: 'xl/drawings/_rels/drawing' + dn + '.xml.rels', text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>' });
        types.push('<Override PartName="/xl/drawings/drawing' + dn + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
      }
    });
    media.forEach((m) => files.push(m));
    files.unshift({ name: 'xl/sharedStrings.xml', text: sst.toXml() });
    files.unshift({ name: 'xl/styles.xml', text: st.toXml() });
    const n = sheets.length;
    files.unshift({ name: 'xl/_rels/workbook.xml.rels', text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + REL_NS + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rId' + (n + 1) + '" Type="' + REL_NS + '/styles" Target="styles.xml"/><Relationship Id="rId' + (n + 2) + '" Type="' + REL_NS + '/sharedStrings" Target="sharedStrings.xml"/></Relationships>' });
    files.unshift({ name: 'xl/workbook.xml', text: XML_HEAD + '<workbook xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' + sheets.map((s, i) => '<sheet name="' + xml(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets><definedNames>' + sheets.map(printName).join('') + '</definedNames><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>' });
    files.unshift({ name: '_rels/.rels', text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL_NS + '/officeDocument" Target="xl/workbook.xml"/></Relationships>' });
    files.unshift({ name: '[Content_Types].xml', text: XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' + types.join('') + '</Types>' });
    return zipStored(files, when);
  }
  /** ファイル名（商品呼出番号表_2026年10月.xlsx） */
  function fileName(month) { const m = /^(\d{4})-(\d{2})$/.exec(String(month)); return '商品呼出番号表_' + (m ? m[1] + '年' + Number(m[2]) + '月' : String(month)) + '.xlsx'; }

  root.SouzaiXlsx = { writeBook, zipStored, crc32, fileName };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
