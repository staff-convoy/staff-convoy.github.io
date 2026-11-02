/* Builds the weekly duty list as .docx (docx library) and .pdf (jsPDF).
   Works in the browser (globals `docx`, `jspdf`) and in Node (for testing). */
(function (root) {
  "use strict";
  var MM_PX = 96 / 25.4;
  var BLUE = "1F4E79", BLUE_RGB = [31, 78, 121], ICE = "BDD7EE";
  var COL_W = 87, GAP = 6, DET_W = 29, HEAD_H = 7.5, BODY_H = 34.8, ROW_GAP = 3;
  var PIC_W = COL_W - DET_W;               // 58 mm
  var CARD_MAX_W = PIC_W - 3, CARD_MAX_H = BODY_H - 2;

  var ROUTES = {
    TO: { label: "TO NAQOURA", route: "Sector West  >  UNIFIL HQ Naqoura", short: "to Naqoura" },
    FROM: { label: "FROM NAQOURA", route: "UNIFIL HQ Naqoura  >  Sector West", short: "from Naqoura" }
  };

  function fit(w, h) {
    var s = Math.min(CARD_MAX_W / w, CARD_MAX_H / h);
    return { w: w * s, h: h * s };
  }

  function fileBase(info) {
    var d = info.date ? info.date.split("-").reverse().join("-") : "no-date";
    return "Critical staff " + d + " - " + ROUTES[info.direction].short;
  }
  // "UNIFIL 1980 / 4 Staff (Driver: Elias FARAH)", one entry per car used this week
  function vehicleItems(info) {
    return (info.vehicles || []).map(function (v) {
      return v.plate + " / " + v.count + " Staff" + (v.driver ? " (Driver: " + v.driver + ")" : "");
    });
  }
  function displayDate(iso) {
    return iso ? iso.split("-").reverse().join("/") : "";
  }

  /* ------------------------------------------------------------ DOCX */
  function buildDocx(D, info, staff) {
    // staff: [{no,name,idNo,section,phone,plate,img:{bytes:Uint8Array,w,h}|null}]
    var R = ROUTES[info.direction];
    var none = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    var noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
    var line = { style: D.BorderStyle.SINGLE, size: 6, color: BLUE };
    var boxBorders = { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line };
    var tw = function (mm) { return Math.round(mm / 25.4 * 1440); };
    var font = "Calibri";
    var run = function (text, o) {
      o = o || {};
      return new D.TextRun({ text: text, font: font, size: Math.round((o.size || 10) * 2), bold: !!o.bold, color: o.color || "000000" });
    };
    var para = function (children, o) {
      o = o || {};
      return new D.Paragraph({ children: children, alignment: o.align, spacing: { before: 0, after: Math.round((o.after || 0) * 20), line: o.line }, keepNext: !!o.keepNext });
    };

    function detail(label, value) {
      return [
        para([run(label.toUpperCase(), { size: 7, bold: true, color: BLUE })]),
        para([run(value || "-", { size: 9.5 })], { after: 3 })
      ];
    }

    function panel(s) {
      var head = new D.TableRow({
        cantSplit: true,
        height: { value: tw(HEAD_H), rule: D.HeightRule.EXACT },
        children: [new D.TableCell({
          columnSpan: 2,
          shading: { type: D.ShadingType.CLEAR, color: "auto", fill: BLUE },
          verticalAlign: D.VerticalAlign.CENTER,
          margins: { left: 100, right: 60, top: 0, bottom: 0 },
          children: [para([
            run(s.no + ".  ", { size: 11, bold: true, color: "FFFFFF" }),
            run(s.name || "", { size: 10, bold: true, color: "FFFFFF" }),
            run("  \u00b7  ID ", { size: 9, color: ICE }),
            run(s.idNo || "", { size: 10, bold: true, color: "FFFFFF" })
          ])]
        })]
      });
      var picChildren;
      if (s.img) {
        var f = fit(s.img.w, s.img.h);
        picChildren = [para([new D.ImageRun({ type: "jpg", data: s.img.bytes,
          transformation: { width: Math.round(f.w * MM_PX), height: Math.round(f.h * MM_PX) } })], { align: D.AlignmentType.CENTER })];
      } else {
        picChildren = [para([run("No ID card on file", { size: 9, color: "808080" })], { align: D.AlignmentType.CENTER })];
      }
      var body = new D.TableRow({
        cantSplit: true,
        height: { value: tw(BODY_H), rule: D.HeightRule.EXACT },
        children: [
          new D.TableCell({ width: { size: tw(DET_W), type: D.WidthType.DXA }, verticalAlign: D.VerticalAlign.CENTER,
            margins: { left: 100, right: 60, top: 20, bottom: 20 },
            children: [].concat(detail("Section / Unit", s.section), detail("Mobile", s.phone), detail("Vehicle plate", s.plate)) }),
          new D.TableCell({ width: { size: tw(PIC_W), type: D.WidthType.DXA }, verticalAlign: D.VerticalAlign.CENTER,
            margins: { left: 40, right: 40, top: 20, bottom: 20 }, children: picChildren })
        ]
      });
      return new D.Table({
        width: { size: tw(COL_W), type: D.WidthType.DXA },
        columnWidths: [tw(DET_W), tw(PIC_W)],
        layout: D.TableLayoutType.FIXED,
        borders: boxBorders,
        rows: [head, body]
      });
    }

    // grid: two staff boxes per row, inside a borderless outer table
    var rows = [];
    for (var i = 0; i < staff.length; i += 2) {
      var cells = [staff[i], staff[i + 1]].map(function (s, k) {
        return new D.TableCell({
          width: { size: tw(COL_W + GAP / 2), type: D.WidthType.DXA },
          margins: { left: k === 0 ? 0 : tw(GAP / 2), right: k === 0 ? tw(GAP / 2) : 0, top: 0, bottom: 0 },
          borders: { top: none, bottom: none, left: none, right: none },
          children: s ? [panel(s), para([run("", { size: 2 })])] : [para([run("")])]
        });
      });
      rows.push(new D.TableRow({ cantSplit: true, height: { value: tw(HEAD_H + BODY_H + ROW_GAP), rule: D.HeightRule.EXACT }, children: cells }));
    }
    var grid = new D.Table({
      width: { size: tw(2 * COL_W + GAP), type: D.WidthType.DXA },
      columnWidths: [tw(COL_W + GAP / 2), tw(COL_W + GAP / 2)],
      layout: D.TableLayoutType.FIXED,
      borders: noBorders,
      rows: rows
    });

    var header = new D.Header({ children: [new D.Paragraph({
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: BLUE, space: 2 } },
      tabStops: [{ type: D.TabStopType.RIGHT, position: tw(180) }],
      children: [run("UNITED NATIONS INTERIM FORCE IN LEBANON (UNIFIL)", { size: 8, bold: true, color: BLUE }),
                 run("\tCRITICAL STAFF \u2013 WEEKLY CONVOY", { size: 8, color: BLUE })]
    })] });
    var footer = new D.Footer({ children: [new D.Paragraph({
      tabStops: [{ type: D.TabStopType.RIGHT, position: tw(180) }],
      children: [run("FOR OFFICIAL USE ONLY  \u00b7  Handle staff personal data in confidence", { size: 8, color: "595959" }),
                 new D.TextRun({ font: font, size: 16, children: ["\tPage ", D.PageNumber.CURRENT, " of ", D.PageNumber.TOTAL_PAGES] })]
    })] });

    var doc = new D.Document({
      creator: "UNIFIL critical staff",
      title: fileBase(info),
      styles: { default: { document: { run: { font: font, size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: tw(210), height: tw(297) },
          margin: { top: tw(14), bottom: tw(14), left: tw(15), right: tw(15), header: tw(7), footer: tw(7) } } },
        headers: { default: header },
        footers: { default: footer },
        children: [
          para([run("CRITICAL STAFF \u2013 WEEKLY CONVOY MOVEMENT", { size: 17, bold: true, color: BLUE })], { align: D.AlignmentType.CENTER, after: 2 }),
          para([run(R.label + ":  ", { size: 12, bold: true, color: BLUE }), run(R.route.replace(">", "\u25b8"), { size: 12, color: BLUE })], { align: D.AlignmentType.CENTER, after: 1 }),
          para([run("Escort: " + (info.escort || "Irish Contingent (IRISHBATT)"), { size: 10, color: "404040" })], { align: D.AlignmentType.CENTER, after: 2 }),
          para([run("Date:  ", { size: 12, bold: true, color: BLUE }), run(displayDate(info.date), { size: 12 }),
                run("        Staff on duty: " + staff.length, { size: 11, color: "404040" })], { align: D.AlignmentType.CENTER, after: vehicleItems(info).length ? 1 : 6 }),
          vehicleItems(info).length ? para([run("Vehicles:  ", { size: 9, bold: true, color: BLUE }), run(vehicleItems(info).join("   \u00b7   "), { size: 9, color: "202020" })], { align: D.AlignmentType.CENTER, after: 4 }) : null,
          grid
        ].filter(Boolean)
      }]
    });
    return doc;
  }

  /* ------------------------------------------------------------ PDF */
  function buildPdf(JsPDF, info, staff) {
    var R = ROUTES[info.direction];
    var pdf = new JsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    var PITCH = HEAD_H + BODY_H + ROW_GAP, LEFT = 15, BOTTOM = 283;
    function chrome() {
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(7.5); pdf.setTextColor(31, 78, 121);
      pdf.text("UNITED NATIONS INTERIM FORCE IN LEBANON (UNIFIL)", LEFT, 12);
      pdf.setFont("helvetica", "normal");
      pdf.text("CRITICAL STAFF - WEEKLY CONVOY", 195, 12, { align: "right" });
      pdf.setDrawColor(31, 78, 121); pdf.setLineWidth(0.35); pdf.line(LEFT, 13.5, 195, 13.5);
    }
    chrome();
    pdf.setFont("helvetica", "bold"); pdf.setFontSize(16); pdf.setTextColor(31, 78, 121);
    pdf.text("CRITICAL STAFF - WEEKLY CONVOY MOVEMENT", 105, 23, { align: "center" });
    pdf.setFontSize(12);
    pdf.text(R.label + ":  " + R.route, 105, 30, { align: "center" });
    pdf.setFont("helvetica", "normal"); pdf.setFontSize(9); pdf.setTextColor(64, 64, 64);
    pdf.text("Escort: " + (info.escort || "Irish Contingent (IRISHBATT)"), 105, 35, { align: "center" });
    pdf.setFontSize(11);
    pdf.setFont("helvetica", "bold"); pdf.setTextColor(31, 78, 121);
    var dateTxt = "Date:  ", dateVal = displayDate(info.date), cnt = "        Staff on duty: " + staff.length;
    pdf.setFont("helvetica", "bold"); var w1 = pdf.getTextWidth(dateTxt);
    pdf.setFont("helvetica", "normal"); var w2 = pdf.getTextWidth(dateVal), w3 = pdf.getTextWidth(cnt);
    var x0 = 105 - (w1 + w2 + w3) / 2;
    pdf.setFont("helvetica", "bold"); pdf.text(dateTxt, x0, 41.5);
    pdf.setFont("helvetica", "normal"); pdf.setTextColor(0, 0, 0); pdf.text(dateVal, x0 + w1, 41.5);
    pdf.setTextColor(64, 64, 64); pdf.text(cnt, x0 + w1 + w2, 41.5);

    var y = 46, items = vehicleItems(info);
    if (items.length) {
      pdf.setFontSize(8.5);
      var lines = [], cur = "Vehicles:  ";
      items.forEach(function (it, i) { var add = (i ? "   \u00b7   " : "") + it;
        if (pdf.getTextWidth(cur + add) > 178 && i) { lines.push(cur); cur = it; } else cur += add; });
      lines.push(cur);
      lines.forEach(function (ln, i) {
        var yy = 46.5 + i * 4.2;
        if (i === 0) { var lab = "Vehicles:  "; pdf.setFont("helvetica", "normal"); var full = pdf.getTextWidth(ln);
          var x0 = 105 - full / 2; pdf.setFont("helvetica", "bold"); pdf.setTextColor(31, 78, 121); pdf.text(lab, x0, yy);
          pdf.setFont("helvetica", "normal"); pdf.setTextColor(32, 32, 32); pdf.text(ln.slice(lab.length), x0 + pdf.getTextWidth(lab) + 0.3, yy); }
        else { pdf.setFont("helvetica", "normal"); pdf.setTextColor(32, 32, 32); pdf.text(ln, 105, yy, { align: "center" }); }
      });
      y = 46.5 + lines.length * 4.2 - 1;
    }
    for (var i = 0; i < staff.length; i += 2) {
      if (y + HEAD_H + BODY_H > BOTTOM) { pdf.addPage(); chrome(); y = 18; }
      for (var k = 0; k < 2; k++) {
        var s = staff[i + k];
        if (!s) continue;
        var x = LEFT + k * (COL_W + GAP);
        pdf.setFillColor(31, 78, 121); pdf.rect(x, y, COL_W, HEAD_H, "F");
        pdf.setTextColor(255, 255, 255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(10.5);
        var t1 = s.no + ".  " + (s.name || "");
        pdf.text(t1, x + 2, y + 5);
        var tx = x + 2 + pdf.getTextWidth(t1);
        pdf.setFont("helvetica", "normal"); pdf.setFontSize(8.5); pdf.setTextColor(189, 215, 238);
        pdf.text("  \u00b7  ID ", tx, y + 5); tx += pdf.getTextWidth("  \u00b7  ID ");
        pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.setTextColor(255, 255, 255);
        pdf.text(s.idNo || "", tx, y + 5);
        pdf.setDrawColor(31, 78, 121); pdf.setLineWidth(0.25);
        pdf.rect(x, y, COL_W, HEAD_H + BODY_H);
        pdf.line(x + DET_W, y + HEAD_H, x + DET_W, y + HEAD_H + BODY_H);
        var dy = y + HEAD_H + 6;
        [["SECTION / UNIT", s.section], ["MOBILE", s.phone], ["VEHICLE PLATE", s.plate]].forEach(function (d) {
          pdf.setFont("helvetica", "bold"); pdf.setFontSize(6.5); pdf.setTextColor(31, 78, 121);
          pdf.text(d[0], x + 2, dy);
          pdf.setFont("helvetica", "normal"); pdf.setFontSize(9); pdf.setTextColor(0, 0, 0);
          pdf.text(pdf.splitTextToSize(d[1] || "-", DET_W - 3)[0], x + 2, dy + 4);
          dy += 10.5;
        });
        if (s.img) {
          var f = fit(s.img.w, s.img.h);
          pdf.addImage(s.img.dataUrl, "JPEG", x + DET_W + (PIC_W - f.w) / 2, y + HEAD_H + (BODY_H - f.h) / 2, f.w, f.h, undefined, "FAST");
        } else {
          pdf.setFontSize(8); pdf.setTextColor(128, 128, 128);
          pdf.text("No ID card on file", x + DET_W + PIC_W / 2, y + HEAD_H + BODY_H / 2, { align: "center" });
        }
      }
      y += PITCH;
    }
    var n = pdf.getNumberOfPages();
    for (var p = 1; p <= n; p++) {
      pdf.setPage(p);
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(7); pdf.setTextColor(89, 89, 89);
      pdf.text("FOR OFFICIAL USE ONLY  \u00b7  Handle staff personal data in confidence", LEFT, 290);
      pdf.text("Page " + p + " of " + n, 195, 290, { align: "right" });
    }
    return pdf;
  }

  var api = { buildDocx: buildDocx, buildPdf: buildPdf, fileBase: fileBase, displayDate: displayDate, ROUTES: ROUTES };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.ConvoyGen = api;
})(typeof window !== "undefined" ? window : this);
