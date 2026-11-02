/* Internal Security Clearance Request (ISCR): fills the supervisor's own Excel form for one convoy trip,
   and draws that same filled sheet as a PDF. Works in the browser (ExcelJS, jsPDF) and in Node for tests. */
(function (root) {
  "use strict";
  var FIRST = 21, ROWS = 4, LAST = 34, COLS = 9; // participant rows 21-24 in the form; the form ends at row 34, column I

  function copyStyle(s) { return JSON.parse(JSON.stringify(s || {})); }

  // move rows from..LAST down by n, with their merges, heights and styles
  function shiftDown(ws, from, n) {
    if (n <= 0) return;
    var merges = (ws.model.merges || []).slice(), moved = [];
    merges.forEach(function (m) { var r = Number(m.match(/\d+/)[0]); if (r >= from) { ws.unMergeCells(m); moved.push(m); } });
    for (var r = LAST; r >= from; r--) {
      var src = ws.getRow(r), dst = ws.getRow(r + n);
      dst.height = src.height;
      for (var c = 1; c <= COLS + 1; c++) { var a = src.getCell(c), b = dst.getCell(c);
        b.value = a.value; b.style = copyStyle(a.style); a.value = null; }
    }
    moved.forEach(function (m) { ws.mergeCells(m.replace(/\d+/g, function (x) { return String(Number(x) + n); })); });
  }

  // data: {requestDate, missionDate, etd, duration, depart, place, when, from, to, remarks, objective, leader, vehicles, people:[{name, phone}]}
  function fill(wb, d) {
    var ws = wb.worksheets[0], people = d.people || [], extra = Math.max(0, people.length - ROWS);
    var put = function (addr, v) { ws.getCell(addr).value = v; };
    put("F4", d.requestDate); put("F6", d.missionDate);
    if (d.requester) put("B6", d.requester);
    put("B8", d.etd); put("C8", d.duration); put("F8", d.depart);
    put("B10", d.place); put("F10", d.place);
    put("B14", d.when); put("C14", d.from); put("F14", d.to); put("H14", d.remarks || "");
    put("B16", "Rotation of Critical Staff : " + (d.objective || ""));
    put("B18", d.leader || ""); put("F18", d.vehicles || "");
    // "Military Escort ...: IRISHBATT Escort" -> the escort of this convoy, keeping the form's own styling
    if (d.escort) { var ec = ws.getCell("B27"), ev = ec.value, swap = function (t) { return String(t).replace(/[A-Z]+BATT|IRISH\w*/, d.escort); };
      if (ev && ev.richText) ec.value = {richText: ev.richText.map(function (r) { return {font: r.font, text: swap(r.text)}; })};
      else if (typeof ev === "string") ec.value = swap(ev); }
    // more people than rows: open new rows that look like the first participant row
    shiftDown(ws, FIRST + ROWS, extra);
    var model = ws.getRow(FIRST);
    for (var i = 0; i < Math.max(ROWS, people.length); i++) {
      var r = FIRST + i, row = ws.getRow(r), p = people[i];
      // every participant row looks exactly like the first one (the form's own rows differ slightly)
      if (i > 0) { row.height = model.height; for (var c = 1; c <= COLS + 1; c++) row.getCell(c).style = copyStyle(model.getCell(c).style); }
      if (!ws.getCell("B" + r).isMerged) ws.mergeCells("B" + r + ":D" + r);
      if (!ws.getCell("G" + r).isMerged) ws.mergeCells("G" + r + ":H" + r);
      row.getCell(2).value = p ? p.name : null;
      row.getCell(5).value = p ? "UNIFIL" : null;
      row.getCell(6).value = p ? p.phone || "" : null;
      row.getCell(7).value = p ? p.callsign || "" : null; row.getCell(9).value = p ? p.plate || "" : null;
    }
    var last = LAST + extra;
    ws.pageSetup.printArea = "A1:I" + last;
    ws.pageSetup.fitToPage = true; ws.pageSetup.fitToWidth = 1; ws.pageSetup.fitToHeight = extra > 6 ? 0 : 1;
    if (ws.model && ws.model.definedNames) {}
    return last;
  }

  /* ---------- the same sheet as a PDF ---------- */
  var THEME = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"];
  function rgb(col, dflt) {
    if (!col) return dflt; var hex;
    if (col.argb) hex = col.argb.slice(-6); else if (col.theme !== undefined) hex = THEME[col.theme] || "000000"; else return dflt;
    var c = [0, 2, 4].map(function (i) { return parseInt(hex.substr(i, 2), 16); }), t = col.tint || 0;
    return c.map(function (x) { return Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t); });
  }
  function text(v) {
    if (v === null || v === undefined) return "";
    if (v.richText) return v.richText.map(function (r) { return r.text; }).join("");
    if (v instanceof Date) return v.toLocaleDateString("en-GB");
    if (typeof v === "object") return v.text || (v.result !== undefined ? String(v.result) : "");
    return String(v);
  }
  var colPt = function (w) { return ((w || 8.43) * 7 + 5) * 0.75; };

  function pdf(jsPDF, wb, logo) {
    var ws = wb.worksheets[0], last = Number(String(ws.pageSetup.printArea || "A1:I34").match(/(\d+)$/)[1]);
    var cw = [], cx = [0];
    for (var c = 1; c <= COLS; c++) { cw[c] = colPt(ws.getColumn(c).width); cx[c] = cx[c - 1] + cw[c - 1 > 0 ? c - 1 : 0] * (c > 1 ? 1 : 0); }
    cx = [0, 0]; for (c = 2; c <= COLS + 1; c++) cx[c] = cx[c - 1] + cw[c - 1];
    var sheetW = cx[COLS + 1], rh = function (r) { return ws.getRow(r).height || 12.75; };
    // merges: master cell -> range; other cells inside are skipped
    var master = {}, inside = {}, rangeOf = {};
    (ws.model.merges || []).forEach(function (m) { var p = m.split(":"), a = ws.getCell(p[0]), b = ws.getCell(p[1]);
      master[a.row + "," + a.col] = {r1: a.row, c1: a.col, r2: b.row, c2: b.col};
      for (var r = a.row; r <= b.row; r++) for (var q = a.col; q <= b.col; q++) { rangeOf[r + "," + q] = master[a.row + "," + a.col]; if (r !== a.row || q !== a.col) inside[r + "," + q] = true; } });
    var doc = new jsPDF({unit: "pt", format: "a4"}), M = 22, pageW = 595.28 - 2 * M, pageH = 841.89 - 2 * M;
    var s = pageW / sheetW, total = 0; for (var r = 1; r <= last; r++) total += rh(r);
    if (total * s > pageH && total * s < pageH * 1.12) s = pageH / total; // nearly fits: one page
    var y = M, rowTop = {}, pageOf = {}, page = 1;
    for (r = 1; r <= last; r++) { if (y + rh(r) * s > M + pageH + 0.5) { page++; y = M; } rowTop[r] = y; pageOf[r] = page; y += rh(r) * s; }
    for (var pg = 2; pg <= page; pg++) doc.addPage();
    var X = function (c) { return M + cx[c] * s; };
    var drawCell = function (r, c) {
      if (inside[r + "," + c]) return;
      var cell = ws.getCell(r, c), mg = master[r + "," + c] || {r1: r, c1: c, r2: r, c2: c};
      doc.setPage(pageOf[r]);
      var x = X(c), yy = rowTop[r], w = X(mg.c2 + 1) - x, h = 0;
      for (var q = mg.r1; q <= mg.r2; q++) h += rh(q) * s;
      var f = cell.fill; if (f && f.type === "pattern" && f.pattern === "solid") { var fc = rgb(f.fgColor, null); if (fc) { doc.setFillColor(fc[0], fc[1], fc[2]); doc.rect(x, yy, w, h, "F"); } }
      return {cell: cell, x: x, y: yy, w: w, h: h, mg: mg};
    };
    var boxes = [];
    for (r = 1; r <= last; r++) for (c = 1; c <= COLS; c++) { var b = drawCell(r, c); if (b) boxes.push(b); }
    // borders: every cell's own sides
    for (r = 1; r <= last; r++) for (c = 1; c <= COLS; c++) {
      var cell = ws.getCell(r, c), bd = cell.border || {}; doc.setPage(pageOf[r]);
      var x1 = X(c), x2 = X(c + 1), y1 = rowTop[r], y2 = y1 + rh(r) * s;
      [["top", x1, y1, x2, y1], ["bottom", x1, y2, x2, y2], ["left", x1, y1, x1, y2], ["right", x2, y1, x2, y2]].forEach(function (e) {
        var side = bd[e[0]]; if (!side || !side.style) return;
        var g = rangeOf[r + "," + c]; // no lines inside a merged box
        if (g && ((e[0] === "left" && c > g.c1) || (e[0] === "right" && c < g.c2) || (e[0] === "top" && r > g.r1) || (e[0] === "bottom" && r < g.r2))) return;
        var col = rgb(side.color, [0, 0, 0]); doc.setDrawColor(col[0], col[1], col[2]);
        doc.setLineWidth(side.style === "medium" ? 1.4 * s * 1.6 : side.style === "thick" ? 2.2 * s * 1.6 : 0.6 * s * 1.6);
        doc.line(e[1], e[2], e[3], e[4]); });
    }
    // text
    boxes.forEach(function (bx) {
      var cell = bx.cell, t = text(cell.value); if (!t.trim()) return;
      var fnt = (cell.value && cell.value.richText && cell.value.richText[0].font) || cell.font || {};
      var size = (fnt.size || 10) * s, style = fnt.bold && fnt.italic ? "bolditalic" : fnt.bold ? "bold" : fnt.italic ? "italic" : "normal";
      var col = rgb(fnt.color, [0, 0, 0]); doc.setPage(pageOf[bx.mg.r1]);
      doc.setFont("helvetica", style); doc.setFontSize(size); doc.setTextColor(col[0], col[1], col[2]);
      var al = cell.alignment || {}, pad = 1.6 * s * 1.6, lh = size * 1.15;
      var lines = [];
      t.replace(/\r/g, "").split("\n").forEach(function (part) { lines = lines.concat(doc.splitTextToSize(part.replace(/\t/g, " "), Math.max(4, bx.w - 2 * pad))); });
      while (lines.length > 1 && !String(lines[lines.length - 1]).trim()) lines.pop();
      var maxLines = Math.max(1, Math.floor((bx.h - pad) / lh)); if (lines.length > maxLines) lines = lines.slice(0, maxLines);
      var blockH = lines.length * lh, v = al.vertical || "bottom";
      var ty = v === "top" ? bx.y + pad + size * 0.85 : v === "middle" || v === "center" ? bx.y + (bx.h - blockH) / 2 + size * 0.85 : bx.y + bx.h - pad - blockH + size * 0.85;
      lines.forEach(function (ln, i) { var lw = doc.getTextWidth(ln), hz = al.horizontal || (typeof cell.value === "number" ? "right" : "left");
        var tx = hz === "center" || hz === "centerContinuous" ? bx.x + (bx.w - lw) / 2 : hz === "right" ? bx.x + bx.w - pad - lw : bx.x + pad;
        doc.text(ln, tx, ty + i * lh);
        if (fnt.underline) { doc.setDrawColor(col[0], col[1], col[2]); doc.setLineWidth(Math.max(0.4, size * 0.06)); doc.line(tx, ty + i * lh + size * 0.12, tx + lw, ty + i * lh + size * 0.12); } });
    });
    // the logo, where the form puts it
    if (logo && logo.data) { var tl = logo.tl, emu = 12700; doc.setPage(1);
      var lx = X(tl.col + 1) + (tl.colOff || 0) / emu * s, ly = rowTop[tl.row + 1] + (tl.rowOff || 0) / emu * s;
      try { doc.addImage(logo.data, logo.ext === "png" ? "PNG" : "JPEG", lx, ly, logo.w * 0.75 * s, logo.h * 0.75 * s); } catch (e) {} }
    return doc;
  }

  // the logo picture inside the form, for the PDF
  function logoOf(wb) {
    var ws = wb.worksheets[0], im = ws.getImages()[0]; if (!im) return null;
    var media = wb.model.media.find(function (m) { return m.index === im.imageId; }) || wb.getImage(Number(im.imageId)); if (!media) return null;
    var buf = media.buffer, b64 = typeof Buffer !== "undefined" && Buffer.from ? Buffer.from(buf).toString("base64") : btoa(Array.prototype.map.call(new Uint8Array(buf), function (x) { return String.fromCharCode(x); }).join(""));
    var tl = im.range.tl, ext = im.range.ext;
    var w = ext ? ext.width : 120, h = ext ? ext.height : 109;
    return {data: "data:image/" + (media.extension === "png" ? "png" : "jpeg") + ";base64," + b64, ext: media.extension,
      tl: {col: Math.floor(tl.nativeCol !== undefined ? tl.nativeCol : tl.col), row: Math.floor(tl.nativeRow !== undefined ? tl.nativeRow : tl.row), colOff: tl.nativeColOff || 0, rowOff: tl.nativeRowOff || 0}, w: w, h: h};
  }

  var api = {fill: fill, pdf: pdf, logoOf: logoOf};
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.ISCR = api;
})(typeof window !== "undefined" ? window : this);
