'use strict';
// 取り込むファイルを読む作業場所（画面の外＝大きな .xls を読んでいる間も、画面は止まらない）。読むだけで、どこへも送らない。
//   何の物か（手作りの番号表・サマリー）は、ファイルの名前でなく中身（見えるシートの名前）で決める。
importScripts('core.js?v=0.2.4', 'search-norm.js?v=0.2.4', 'book-read.js?v=0.2.4', 'import-table.js?v=0.2.4', 'import-summary.js?v=0.2.4');
function kindOf(book) {
  const names = book.sheets.filter((s) => s.state === 'visible').map((s) => s.name);
  if (names.some((n) => /寿司・弁当|焼き物/.test(n))) return 'table';
  if (names.some((n) => /サマリー/.test(n))) return 'summary';
  return 'unknown';
}
self.onmessage = async (e) => {
  const { id, name, bytes, typeTables } = e.data || {};
  const t0 = Date.now();
  try {
    const book = await self.SouzaiBook.readBook(bytes);
    const kind = kindOf(book);
    let result = null;
    if (kind === 'table') result = self.SouzaiImportTable.readTableBook(book, { fileName: name, typeTables }); // 書式の鍵は、アプリの表の型から選ぶ
    else if (kind === 'summary') result = self.SouzaiImportSummary.readSummaryBook(book, { fileName: name });
    self.postMessage({ id, ok: true, kind, format: book.format, result, ms: Date.now() - t0 });
  } catch (error) {
    self.postMessage({ id, ok: false, message: String((error && error.message) || error), ms: Date.now() - t0 });
  }
};
