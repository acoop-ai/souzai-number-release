'use strict';
// パックの値段の絵（焼き物の表の「≪パック販売用値段特≫」）。数（core.js の packOf）から、元の書類の絵と同じ形を描く。
//   Excel の絵（PNG）と番号表の画面の絵は、ここで描いた同じ物（2つに分けない）。窓口は数だけを持つ。描けない所（試験の node）では null を返し、元の絵を使う。
(function (root) {
  const C = root.SouzaiCore;
  const W = 690, H = 412, SCALE = 2; // 元の絵の大きさ（Excel の置き場はこの縦横の比）。印刷で粗くならないよう2倍で描く
  const TITLE_H = 68;
  const FONT = '"MS PGothic", "ＭＳ Ｐゴシック", "Meiryo", "Yu Gothic", sans-serif';
  const memo = new Map();
  function draw(rows) {
    const key = C.packText(rows);
    if (memo.has(key)) return memo.get(key);
    let out = null;
    try {
      const cv = root.document.createElement('canvas'); cv.width = W * SCALE; cv.height = H * SCALE;
      const g = cv.getContext('2d');
      if (g) {
        g.scale(SCALE, SCALE);
        g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
        g.fillStyle = '#FFC000'; g.fillRect(0, 0, W, TITLE_H);
        g.fillStyle = '#000'; g.textBaseline = 'middle';
        g.font = '36px ' + FONT; g.textAlign = 'center'; g.fillText(C.PACK_TITLE, W / 2, TITLE_H / 2 + 2, W - 24);
        // 段: 1個 109円 ⇒ ｜ 2個 199円 ｜ 3個 280円（段が多いほど字を小さく）
        const n = rows.length, rh = (H - TITLE_H) / n, fs = Math.min(34, Math.floor(rh * 0.52));
        g.font = fs + 'px ' + FONT;
        const put = (text, x, align, max) => { g.textAlign = align; g.fillText(text, x, 0, max); };
        rows.forEach((r, i) => {
          g.save(); g.translate(0, TITLE_H + rh * i + rh / 2 + 1);
          put('1個', 8, 'left'); put(r[0] + '円', 214, 'right', 138); put('⇒', 243, 'center');
          put('2個', 277, 'left'); put(r[1] + '円', 471, 'right', 140);
          put('3個', 488, 'left'); put(r[2] + '円', 682, 'right', 140);
          g.restore();
        });
        // 線: 段の間は点線・縦の区切りは細い実線・題の下は太い線・外は太い枠
        g.strokeStyle = '#000';
        const line = (x1, y1, x2, y2) => { g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
        g.lineWidth = 1.5; g.setLineDash([2, 2]);
        for (let i = 1; i < n; i++) line(0, TITLE_H + rh * i, W, TITLE_H + rh * i);
        g.setLineDash([]); g.lineWidth = 2;
        [271, 481].forEach((x) => line(x, TITLE_H, x, H));
        g.lineWidth = 3; line(0, TITLE_H, W, TITLE_H);
        g.lineWidth = 4; g.strokeRect(2, 2, W - 4, H - 4);
        const url = cv.toDataURL('image/png');
        out = { url, png: url.slice(url.indexOf(',') + 1) };
      }
    } catch (_) { out = null; }
    memo.set(key, out);
    return out;
  }
  root.SouzaiPack = { draw, W, H };
})(typeof window !== 'undefined' ? window : globalThis);
