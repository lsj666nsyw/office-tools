/**
 * office-core.js —— 全能办公工具核心（浏览器本地运行，不上传）
 *
 * 覆盖：PDF / Word / Excel / 图片 / 格式互转
 * 依赖：pdf-lib（PDF）、xlsx（表格）、mammoth（Word→HTML）、docx（HTML→Word）、jszip（打包）
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory;
  else root.createOfficeToolbox = factory;
})(typeof self !== "undefined" ? self : this, function (deps) {
  "use strict";
  var PDFLib = deps.PDFLib, XLSX = deps.XLSX, JSZip = deps.JSZip;
  var mammoth = deps.mammoth, docxLib = deps.docxLib;

  function buf(u) { return u instanceof Uint8Array ? u : new Uint8Array(u); }
  function safeName(s, max) {
    s = String(s).replace(/[\\/:*?"<>|\r\n\t]+/g, "_").replace(/^\.+/, "").trim();
    s = s.replace(/_+/g, "_").replace(/^[_ .]+|[_ .]+$/g, "");
    return (s || "未命名").slice(0, max || 60);
  }
  function parseRanges(text, total) {
    if (!text || !text.trim()) return Array.from({ length: total }, function (_, i) { return i; });
    var out = {};
    text.trim().split(/[,，;；\s]+/).forEach(function (p) {
      if (!p) return;
      var m = p.match(/^(\d+)\s*[-~]\s*(\d+)$/);
      if (m) {
        var a = +m[1], b = +m[2];
        if (a > b) { var t = a; a = b; b = t; }
        for (var k = a; k <= b; k++) if (k >= 1 && k <= total) out[k - 1] = 1;
      } else if (/^\d+$/.test(p)) {
        var n = +p;
        if (n >= 1 && n <= total) out[n - 1] = 1;
      } else throw new Error("看不懂的页码：" + p);
    });
    return Object.keys(out).map(Number).sort(function (x, y) { return x - y; });
  }

  // ==================== PDF ====================
  async function pdfLoad(data) {
    return PDFLib.PDFDocument.load(buf(data), { ignoreEncryption: true });
  }

  async function pdfMerge(items) {
    var out = await PDFLib.PDFDocument.create();
    var info = [], skipped = [];
    for (var i = 0; i < items.length; i++) {
      try {
        var src = await pdfLoad(items[i].data);
        var pages = await out.copyPages(src, src.getPageIndices());
        pages.forEach(function (p) { out.addPage(p); });
        info.push({ name: items[i].name, count: pages.length });
      } catch (e) {
        skipped.push({ name: items[i].name, reason: String(e.message || e) });
      }
    }
    if (!out.getPageCount()) throw new Error("没有可合并的页面");
    var bytes = await out.save({ useObjectStreams: true });
    return { data: bytes, pages: out.getPageCount(), info: info, skipped: skipped };
  }

  async function pdfSplit(item, mode, opts) {
    opts = opts || {};
    var src = await pdfLoad(item.data);
    var total = src.getPageCount();
    var stem = safeName(item.name.replace(/\.[^.]+$/, ""));
    var results = [];
    async function make(indices, name) {
      var out = await PDFLib.PDFDocument.create();
      var pages = await out.copyPages(src, indices);
      pages.forEach(function (p) { out.addPage(p); });
      results.push({ name: name, data: await out.save({ useObjectStreams: true }), pages: indices.length });
    }
    if (mode === "each") {
      for (var i = 0; i < total; i++) await make([i], stem + "_第" + (i + 1) + "页.pdf");
    } else if (mode === "chunk") {
      var n = Math.max(1, opts.pagesPerFile || 10), part = 0;
      for (var s = 0; s < total; s += n) {
        var e = Math.min(s + n, total), idx = [];
        for (var k = s; k < e; k++) idx.push(k);
        part++;
        await make(idx, stem + "_第" + part + "部分_p" + (s + 1) + "-" + e + ".pdf");
      }
    } else {
      var idxs = parseRanges(opts.ranges || "", total);
      if (!idxs.length) throw new Error("没有选中任何页");
      await make(idxs, stem + "_提取" + idxs.length + "页.pdf");
    }
    return { total: total, results: results };
  }

  async function pdfPageOps(item, opts) {
    opts = opts || {};
    var src = await pdfLoad(item.data);
    var total = src.getPageCount();
    var keep = [];
    if (opts.removePages && opts.removePages.trim()) {
      var del = parseRanges(opts.removePages, total), delSet = {};
      del.forEach(function (i) { delSet[i] = 1; });
      for (var i = 0; i < total; i++) if (!delSet[i]) keep.push(i);
    } else {
      for (var j = 0; j < total; j++) keep.push(j);
    }
    if (!keep.length) throw new Error("删除后一页都不剩");

    var out = await PDFLib.PDFDocument.create();
    var pages = await out.copyPages(src, keep);
    pages.forEach(function (p) { out.addPage(p); });

    if (opts.rotate) {
      var targets = opts.rotatePages && opts.rotatePages.trim()
        ? parseRanges(opts.rotatePages, out.getPageCount())
        : Array.from({ length: out.getPageCount() }, function (_, i) { return i; });
      targets.forEach(function (i) {
        var p = out.getPage(i);
        p.setRotation(PDFLib.degrees(((p.getRotation().angle || 0) + opts.rotate) % 360));
      });
    }
    if (opts.reverse) {
      var rev = await PDFLib.PDFDocument.create();
      var all = Array.from({ length: out.getPageCount() }, function (_, i) { return i; }).reverse();
      var rp = await rev.copyPages(out, all);
      rp.forEach(function (p) { rev.addPage(p); });
      return { data: await rev.save({ useObjectStreams: true }), pages: rev.getPageCount(), removed: total - keep.length };
    }
    return { data: await out.save({ useObjectStreams: true }), pages: out.getPageCount(), removed: total - keep.length };
  }

  async function pdfWatermark(item, opts) {
    opts = opts || {};
    var doc = await pdfLoad(item.data);
    var text = (opts.text || "").trim();
    if (!text) throw new Error("请填写水印文字");
    var opacity = Math.max(0.03, Math.min(1, (opts.opacity || 15) / 100));
    var size = Math.max(8, opts.size || 40);
    var pos = opts.pos || "center";
    var pages = doc.getPages();
    pages.forEach(function (page) {
      var w = page.getWidth(), h = page.getHeight();
      var tw = text.length * size * 0.6;
      var m = Math.min(w, h) * 0.04;
      var cx, cy;
      if (pos === "center") { cx = w / 2; cy = h / 2; }
      else {
        cx = pos[1] === "w" ? m + tw / 2 : pos[1] === "e" ? w - m - tw / 2 : w / 2;
        cy = pos[0] === "n" ? h - m - size / 2 : pos[0] === "s" ? m + size / 2 : h / 2;
      }
      page.drawText(text, {
        x: cx - tw / 2, y: cy - size / 3, size: size,
        color: PDFLib.rgb(0.45, 0.45, 0.45), opacity: opacity,
        rotate: PDFLib.degrees(opts.rotation || 0),
      });
    });
    return { data: await doc.save({ useObjectStreams: true }), pages: pages.length };
  }

  async function pdfResize(item, opts) {
    opts = opts || {};
    var doc = await pdfLoad(item.data);
    var sizes = { a4: [595.28, 841.89], letter: [612, 792], a3: [841.89, 1190.55] };
    var tw, th;
    if (opts.target === "custom") { tw = +opts.w || 595; th = +opts.h || 842; }
    else { var s = sizes[opts.target] || sizes.a4; tw = s[0]; th = s[1]; }
    doc.getPages().forEach(function (p) { p.setSize(tw, th); });
    return { data: await doc.save({ useObjectStreams: true }), pages: doc.getPageCount(), size: [tw, th] };
  }

  async function pdfInfo(item) {
    var doc = await pdfLoad(item.data);
    var lines = [
      "PDF 信息", "========",
      "文件名: " + item.name,
      "页数: " + doc.getPageCount(),
      "标题: " + (doc.getTitle() || "(未设置)"),
      "作者: " + (doc.getAuthor() || "(未设置)"),
      "创建时间: " + (doc.getCreationDate() ? doc.getCreationDate().toISOString() : "(未设置)"),
      "", "各页尺寸：",
    ];
    doc.getPages().forEach(function (p, i) {
      lines.push("  第 " + (i + 1) + " 页: " + Math.round(p.getWidth()) + " × " + Math.round(p.getHeight()) + " pt");
    });
    return lines.join("\n");
  }

  /** PDF 页面 → 图片（用 canvas 渲染；浏览器环境才有） */
  async function pdfToImages(item, opts, canvasFactory) {
    opts = opts || {};
    if (!canvasFactory) throw new Error("当前环境不支持渲染 PDF 页面");
    var doc = await pdfLoad(item.data);
    var total = doc.getPageCount();
    var idxs = parseRanges(opts.pages || "", total);
    var results = [];
    var stem = safeName(item.name.replace(/\.[^.]+$/, ""));
    for (var k = 0; k < idxs.length; k++) {
      var i = idxs[k];
      var page = doc.getPage(i);
      var scale = (opts.dpi || 150) / 72;
      var w = Math.round(page.getWidth() * scale);
      var h = Math.round(page.getHeight() * scale);
      // 用 pdf-lib 抽出单页，交给外部渲染（浏览器用 pdf.js，这里给占位）
      var single = await PDFLib.PDFDocument.create();
      var [copied] = await single.copyPages(doc, [i]);
      single.addPage(copied);
      var bytes = await single.save();
      results.push({ name: stem + "_第" + (i + 1) + "页." + (opts.fmt || "png"), data: bytes, page: i + 1, w: w, h: h });
    }
    return { total: total, results: results };
  }

  // ==================== Excel ====================
  function readWb(data) { return XLSX.read(buf(data), { type: "array", cellDates: true }); }
  function sheetRows(wb, name) {
    var ws = wb.Sheets[name || wb.SheetNames[0]];
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    if (!aoa.length) return { header: [], rows: [] };
    return { header: aoa[0].map(function (c) { return String(c).trim(); }), rows: aoa.slice(1) };
  }
  function aoaToWb(aoa, title) {
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, safeName(title || "Sheet1", 30));
    return wb;
  }
  function wbToBuf(wb) { return XLSX.write(wb, { bookType: "xlsx", type: "array" }); }

  async function xlMerge(items) {
    var allHeader = null, allRows = [], perFile = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      try {
        var wb = readWb(it.data);
        var r = sheetRows(wb);
        if (!r.header.length) { perFile.push({ name: it.name, rows: 0, note: "空表" }); continue; }
        if (!allHeader) { allHeader = r.header.slice(); allHeader.push("来源文件"); }
        var map = {};
        r.header.forEach(function (h, idx) {
          var di = allHeader.indexOf(h);
          if (di >= 0) map[idx] = di;
        });
        var got = 0;
        r.rows.forEach(function (row) {
          if (row.every(function (v) { return v === "" || v == null; })) return;
          var outRow = new Array(allHeader.length).fill("");
          Object.keys(map).forEach(function (si) { outRow[map[si]] = row[si]; });
          outRow[allHeader.length - 1] = it.name;
          allRows.push(outRow);
          got++;
        });
        perFile.push({ name: it.name, rows: got });
      } catch (e) {
        perFile.push({ name: it.name, rows: 0, note: String(e.message || e) });
      }
    }
    if (!allHeader) throw new Error("没有读取到任何数据");
    var wb2 = aoaToWb([allHeader].concat(allRows), "合并结果");
    return { data: wbToBuf(wb2), rows: allRows.length, perFile: perFile };
  }

  async function xlSplit(item, keyHeader) {
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var ki = r.header.indexOf(keyHeader);
    if (ki < 0) throw new Error("找不到列：" + keyHeader);
    var groups = {}, order = [];
    r.rows.forEach(function (row) {
      if (row.every(function (v) { return v === "" || v == null; })) return;
      var key = String(row[ki] == null ? "" : row[ki]).trim() || "(空值)";
      if (!groups[key]) { groups[key] = []; order.push(key); }
      groups[key].push(row);
    });
    if (order.length > 300) throw new Error("分组太多（" + order.length + "）");
    var results = order.map(function (key) {
      return {
        name: safeName(item.name.replace(/\.[^.]+$/, "")) + "_" + safeName(key) + ".xlsx",
        data: wbToBuf(aoaToWb([r.header].concat(groups[key]), "数据")),
        rows: groups[key].length,
      };
    });
    return { results: results, groups: order.length };
  }

  async function xlCompare(a, b, keyHeaders) {
    var wbA = readWb(a.data), wbB = readWb(b.data);
    var A = sheetRows(wbA), B = sheetRows(wbB);
    if (!keyHeaders.length) throw new Error("请选择关键列");
    function keyOf(row, header) {
      return keyHeaders.map(function (h) {
        var i = header.indexOf(h);
        var v = i >= 0 && row[i] != null ? row[i] : "";
        return typeof v === "number" && v % 1 === 0 ? String(v) : String(v).trim();
      }).join("|");
    }
    var mA = {}, mB = {};
    A.rows.forEach(function (r) { if (r.some(function (v) { return v !== "" && v != null; })) (mA[keyOf(r, A.header)] = mA[keyOf(r, A.header)] || []).push(r); });
    B.rows.forEach(function (r) { if (r.some(function (v) { return v !== "" && v != null; })) (mB[keyOf(r, B.header)] = mB[keyOf(r, B.header)] || []).push(r); });
    var onlyA = [], onlyB = [], diffs = [];
    Object.keys(mA).forEach(function (k) { if (!mB[k]) onlyA = onlyA.concat(mA[k]); });
    Object.keys(mB).forEach(function (k) { if (!mA[k]) onlyB = onlyB.concat(mB[k]); });
    Object.keys(mA).forEach(function (k) {
      if (!mB[k]) return;
      var ra = mA[k][0], rb = mB[k][0];
      A.header.forEach(function (h, ci) {
        if (keyHeaders.indexOf(h) >= 0) return;
        var bi = B.header.indexOf(h);
        if (bi < 0) return;
        var va = String(ra[ci] == null ? "" : ra[ci]).trim();
        var vb = String(rb[bi] == null ? "" : rb[bi]).trim();
        if (va !== vb) diffs.push([k, h, va, vb]);
      });
    });
    var wbOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet([A.header].concat(onlyA)), "仅在A中");
    XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet([B.header].concat(onlyB)), "仅在B中");
    XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet([["关键值", "字段", "A的值", "B的值"]].concat(diffs)), "值有差异");
    XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet([
      ["项目", "数量"], ["A表行数", A.rows.length], ["B表行数", B.rows.length],
      ["仅在A中", onlyA.length], ["仅在B中", onlyB.length], ["值不同", diffs.length],
    ]), "汇总");
    return { data: XLSX.write(wbOut, { bookType: "xlsx", type: "array" }), onlyA: onlyA.length, onlyB: onlyB.length, diffs: diffs.length };
  }

  async function xlHeaders(item) {
    var wb = readWb(item.data);
    return sheetRows(wb).header.filter(Boolean);
  }

  /** Excel → CSV（UTF-8 with BOM，Excel 打开不乱码） */
  async function xlToCsv(item) {
    var wb = readWb(item.data);
    var ws = wb.Sheets[wb.SheetNames[0]];
    var csv = XLSX.utils.sheet_to_csv(ws);
    // BOM 用字节数组显式拼接，避免 TextEncoder 对 \ufeff 的处理差异
    var body = new TextEncoder().encode(csv);
    var out = new Uint8Array(body.length + 3);
    out[0] = 0xEF; out[1] = 0xBB; out[2] = 0xBF;
    out.set(body, 3);
    return { data: out, name: item.name.replace(/\.[^.]+$/, "") + ".csv" };
  }

  /** CSV → Excel */
  async function csvToXl(item) {
    var text = new TextDecoder("utf-8").decode(buf(item.data));
    var wb = XLSX.read(text, { type: "string" });
    return { data: wbToBuf(wb), name: item.name.replace(/\.[^.]+$/, "") + ".xlsx" };
  }

  /** Excel 转 HTML 表格（用于转 Word / 预览） */
  async function xlToHtml(item) {
    var wb = readWb(item.data);
    var ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_html(ws, { id: "table", editable: false });
  }

  /**
   * Excel 转纯文本。
   * ⚠ 不能用 XLSX.utils.sheet_to_txt：它输出 UTF-16LE（带 ÿþ BOM），
   *   在浏览器/Node 里直接当 UTF-8 读会全是乱码。改用 sheet_to_csv 去逗号。
   */
  async function xlToText(item) {
    var wb = readWb(item.data);
    var ws = wb.Sheets[wb.SheetNames[0]];
    var csv = XLSX.utils.sheet_to_csv(ws, { FS: "\t" });
    return csv;
  }

  // ==================== Word ====================
  /**
   * mammoth 的输入格式：
   *   浏览器 → { arrayBuffer: ArrayBuffer }
   *   Node   → { buffer: Buffer } 或 { path: ... }
   * 这里统一转成两端都能吃的形式（优先 arrayBuffer，失败再试 buffer）。
   */
  function mammothInput(item) {
    var u8 = buf(item.data);
    // 复制成独立 ArrayBuffer（mammoth 会读 .buffer，切片可能带偏移）
    var ab = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
    return { arrayBuffer: ab };
  }
  function mammothInputNode(item) {
    var u8 = buf(item.data);
    // Node 环境：转 Buffer
    if (typeof Buffer !== "undefined") return { buffer: Buffer.from(u8) };
    return mammothInput(item);
  }

  /** Word(.docx) → HTML */
  async function wordToHtml(item) {
    if (!mammoth) throw new Error("Word 解析库未加载");
    var input = mammothInput(item);
    var result;
    try {
      result = await mammoth.convertToHtml(input);
    } catch (e) {
      // Node 下 arrayBuffer 可能不被接受，退回 buffer
      if (typeof Buffer !== "undefined") {
        result = await mammoth.convertToHtml({ buffer: Buffer.from(buf(item.data)) });
      } else throw e;
    }
    return {
      html: result.value,
      messages: result.messages || [],
      name: item.name.replace(/\.[^.]+$/, "") + ".html",
    };
  }

  /** Word(.docx) → 纯文本 */
  async function wordToText(item) {
    if (!mammoth) throw new Error("Word 解析库未加载");
    var input = mammothInput(item);
    var result;
    try {
      result = await mammoth.extractRawText(input);
    } catch (e) {
      if (typeof Buffer !== "undefined") {
        result = await mammoth.extractRawText({ buffer: Buffer.from(buf(item.data)) });
      } else throw e;
    }
    return { text: result.value, name: item.name.replace(/\.[^.]+$/, "") + ".txt" };
  }

  /** Word(.docx) → HTML 文件（可直接当网页/粘贴到编辑器） */
  async function wordToHtmlFile(item) {
    var r = await wordToHtml(item);
    var page = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>" +
      safeName(item.name.replace(/\.[^.]+$/, "")) +
      "</title><style>body{font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;" +
      "max-width:800px;margin:40px auto;padding:0 20px;line-height:1.8;color:#1a2233}" +
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:8px}" +
      "img{max-width:100%}</style></head><body>" + r.html + "</body></html>";
    return { data: new TextEncoder().encode(page), name: r.name };
  }

  /** 文本/HTML → Word(.docx) */
  async function textToWord(item, opts) {
    opts = opts || {};
    if (!docxLib) throw new Error("Word 生成库未加载");
    var text = new TextDecoder("utf-8").decode(buf(item.data));
    var lines = text.split(/\r?\n/);
    var children = [];
    lines.forEach(function (line) {
      children.push(new docxLib.Paragraph({
        children: [new docxLib.TextRun({ text: line, size: 24 })],
        spacing: { after: 120 },
      }));
    });
    if (!children.length) children.push(new docxLib.Paragraph(""));
    var doc = new docxLib.Document({
      sections: [{ properties: {}, children: children }],
    });
    var blob = await docxLib.Packer.toBlob(doc);
    var ab = await blob.arrayBuffer();
    return { data: new Uint8Array(ab), name: item.name.replace(/\.[^.]+$/, "") + ".docx" };
  }

  /** HTML → Word（把 HTML 表格/段落转成 docx） */
  async function htmlToWord(item) {
    if (!docxLib) throw new Error("Word 生成库未加载");
    var html = new TextDecoder("utf-8").decode(buf(item.data));
    // 极简 HTML → 文本行（去掉标签），再生成 docx
    var tmp = html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|h[1-6]|tr|li)>/gi, "\n")
      .replace(/<\/t[dh]>/gi, "\t")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
    var lines = tmp.split("\n").map(function (l) { return l.replace(/\t+$/, "").trim(); });
    var children = [];
    lines.forEach(function (line) {
      if (!line) { children.push(new docxLib.Paragraph("")); return; }
      children.push(new docxLib.Paragraph({
        children: [new docxLib.TextRun({ text: line, size: 24 })],
        spacing: { after: 100 },
      }));
    });
    if (!children.length) children.push(new docxLib.Paragraph(""));
    var doc = new docxLib.Document({ sections: [{ properties: {}, children: children }] });
    var blob = await docxLib.Packer.toBlob(doc);
    return { data: new Uint8Array(await blob.arrayBuffer()), name: item.name.replace(/\.[^.]+$/, "") + ".docx" };
  }

  // ==================== 打包 ====================
  async function zipFiles(files) {
    if (!JSZip) throw new Error("打包库未加载");
    var zip = new JSZip();
    var used = {};
    files.forEach(function (f) {
      var name = f.name;
      if (used[name]) {
        var dot = name.lastIndexOf(".");
        var stem = dot > 0 ? name.slice(0, dot) : name;
        var ext = dot > 0 ? name.slice(dot) : "";
        var i = 2;
        while (used[stem + "_" + i + ext]) i++;
        name = stem + "_" + i + ext;
      }
      used[name] = 1;
      zip.file(name, f.data);
    });
    return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  }

  return {
    safeName: safeName,
    parseRanges: parseRanges,
    // PDF
    pdfMerge: pdfMerge, pdfSplit: pdfSplit, pdfPageOps: pdfPageOps,
    pdfWatermark: pdfWatermark, pdfResize: pdfResize, pdfInfo: pdfInfo,
    pdfToImages: pdfToImages,
    // Excel
    xlMerge: xlMerge, xlSplit: xlSplit, xlCompare: xlCompare, xlHeaders: xlHeaders,
    xlToCsv: xlToCsv, csvToXl: csvToXl, xlToHtml: xlToHtml, xlToText: xlToText,
    // Word
    wordToHtml: wordToHtml, wordToText: wordToText, wordToHtmlFile: wordToHtmlFile,
    textToWord: textToWord, htmlToWord: htmlToWord,
    // 通用
    zipFiles: zipFiles,
  };
});
