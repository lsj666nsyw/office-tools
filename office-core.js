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

  /** PDF 页面 → 图片（用 pdf.js 渲染，浏览器环境） */
  async function pdfToImages(item, opts, pdfjsLib) {
    opts = opts || {};
    var lib = pdfjsLib || (typeof window !== "undefined" ? window.pdfjsLib : null);
    if (!lib) throw new Error("PDF 渲染库未加载");
    var doc = await pdfLibLoad(lib, buf(item.data));
    var total = doc.numPages;
    var idxs = parseRanges(opts.pages || "", total);
    var stem = safeName(item.name.replace(/\.[^.]+$/, ""));
    var fmt = opts.fmt || "png";
    var scale = Math.max(0.2, Math.min(8, (opts.dpi || 150) / 72));
    var results = [];
    for (var k = 0; k < idxs.length; k++) {
      var i = idxs[k];
      var page = await doc.getPage(i + 1);
      var vp = page.getViewport({ scale: scale });
      var canvas = document.createElement("canvas");
      canvas.width = Math.round(vp.width);
      canvas.height = Math.round(vp.height);
      var ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      var blob = await new Promise(function (res) {
        canvas.toBlob(res, fmt === "jpg" ? "image/jpeg" : "image/png", 0.92);
      });
      results.push({
        name: stem + "_第" + (i + 1) + "页." + fmt,
        data: new Uint8Array(await blob.arrayBuffer()),
        page: i + 1, w: canvas.width, h: canvas.height,
      });
    }
    return { total: total, results: results };
  }

  async function pdfLibLoad(lib, u8, ms) {
    // pdf.js 需要独立 ArrayBuffer（worker 会转移所有权）
    var ab = u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
    // 带超时：worker 起不来时（如某些沙箱/内置浏览器）promise 永不 resolve，
    // 没有超时会让用户点完按钮一直转圈且得不到任何提示。
    return withTimeout(lib.getDocument({ data: ab }).promise, ms || 45000, "PDF 打开超时，请重试或换用本地版");
  }

  /** PDF → 纯文本（用 pdf.js 提取文字层） */
  async function pdfToText(item, opts, pdfjsLib) {
    opts = opts || {};
    var lib = pdfjsLib || (typeof window !== "undefined" ? window.pdfjsLib : null);
    if (!lib) throw new Error("PDF 渲染库未加载");
    var doc = await pdfLibLoad(lib, buf(item.data));
    var total = doc.numPages;
    var idxs = parseRanges(opts.pages || "", total);
    var chunks = [];
    for (var k = 0; k < idxs.length; k++) {
      var i = idxs[k];
      var page = await doc.getPage(i + 1);
      var tc = await page.getTextContent();
      var line = tc.items.map(function (it) { return it.str; }).join("");
      chunks.push("===== 第 " + (i + 1) + " 页 =====\n" + line);
    }
    var text = chunks.join("\n\n");
    return {
      text: text,
      chars: text.length,
      pages: idxs.length,
      name: item.name.replace(/\.[^.]+$/, "") + ".txt",
    };
  }

  /** PDF → 单张长图（所有页纵向拼接） */
  async function pdfToLongImage(item, opts, pdfjsLib) {
    opts = opts || {};
    var lib = pdfjsLib || (typeof window !== "undefined" ? window.pdfjsLib : null);
    if (!lib) throw new Error("PDF 渲染库未加载");
    var doc = await pdfLibLoad(lib, buf(item.data));
    var total = doc.numPages;
    var idxs = parseRanges(opts.pages || "", total);
    var scale = Math.max(0.2, Math.min(4, (opts.dpi || 120) / 72));
    var canvases = [];
    var totalH = 0, maxW = 0;
    for (var k = 0; k < idxs.length; k++) {
      var page = await doc.getPage(idxs[k] + 1);
      var vp = page.getViewport({ scale: scale });
      var c = document.createElement("canvas");
      c.width = Math.round(vp.width);
      c.height = Math.round(vp.height);
      var ctx = c.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      canvases.push(c);
      totalH += c.height;
      maxW = Math.max(maxW, c.width);
    }
    var out = document.createElement("canvas");
    out.width = maxW;
    out.height = totalH;
    var octx = out.getContext("2d");
    octx.fillStyle = "#fff";
    octx.fillRect(0, 0, maxW, totalH);
    var y = 0;
    canvases.forEach(function (c) {
      octx.drawImage(c, Math.round((maxW - c.width) / 2), y);
      y += c.height;
    });
    var blob = await new Promise(function (res) { out.toBlob(res, "image/png"); });
    return {
      data: new Uint8Array(await blob.arrayBuffer()),
      name: safeName(item.name.replace(/\.[^.]+$/, "")) + "_长图.png",
      w: maxW, h: totalH, pages: idxs.length,
    };
  }

  /** 图片 → PDF */
  async function imagesToPdf(items, opts) {
    opts = opts || {};
    var doc = await PDFLib.PDFDocument.create();
    var perFile = [];
    for (var i = 0; i < items.length; i++) {
      try {
        var it = items[i];
        var u8 = buf(it.data);
        var type = it.type || "";
        var img = null;
        if (/png/i.test(type) || /\.png$/i.test(it.name)) {
          img = await doc.embedPng(u8);
        } else if (/jpe?g/i.test(type) || /\.jpe?g$/i.test(it.name)) {
          img = await doc.embedJpg(u8);
        } else {
          // 其它格式（webp/bmp/gif）：先用 canvas 转 PNG
          var blob = new Blob([u8], { type: type || "image/*" });
          var bmp = await createImageBitmap(blob).catch(function () { return null; });
          if (!bmp) throw new Error("不支持的图片格式");
          var c = document.createElement("canvas");
          c.width = bmp.width; c.height = bmp.height;
          c.getContext("2d").drawImage(bmp, 0, 0);
          var pngBlob = await new Promise(function (res) { c.toBlob(res, "image/png"); });
          img = await doc.embedPng(new Uint8Array(await pngBlob.arrayBuffer()));
        }
        var target = opts.target || "fit";
        var page;
        if (target === "fit") {
          page = doc.addPage([img.width + 24, img.height + 24]);
          page.drawImage(img, { x: 12, y: 12, width: img.width, height: img.height });
        } else {
          var sizes = { a4: [595.28, 841.89], letter: [612, 792], a3: [841.89, 1190.55] };
          var s = sizes[target] || sizes.a4;
          page = doc.addPage(s);
          var sc = Math.min((s[0] - 24) / img.width, (s[1] - 24) / img.height);
          var w = img.width * sc, h = img.height * sc;
          page.drawImage(img, { x: (s[0] - w) / 2, y: (s[1] - h) / 2, width: w, height: h });
        }
        perFile.push({ name: it.name, ok: true });
      } catch (e) {
        perFile.push({ name: items[i].name, ok: false, reason: String(e.message || e) });
      }
    }
    if (!doc.getPageCount()) throw new Error("没有成功插入任何图片");
    return { data: await doc.save({ useObjectStreams: true }), pages: doc.getPageCount(), perFile: perFile };
  }

  /** PDF 提取图片（把页面里的内嵌图片抽出来） */
  async function pdfExtractImages(item, pdfjsLib) {
    var lib = pdfjsLib || (typeof window !== "undefined" ? window.pdfjsLib : null);
    if (!lib) throw new Error("PDF 渲染库未加载");
    var doc = await pdfLibLoad(lib, buf(item.data));
    var stem = safeName(item.name.replace(/\.[^.]+$/, ""));
    var results = [];
    for (var p = 1; p <= doc.numPages; p++) {
      var page = await doc.getPage(p);
      var ops = await page.getOperatorList();
      var imgIndex = 0;
      for (var i = 0; i < ops.fnArray.length; i++) {
        var fn = ops.fnArray[i];
        // OPS.paintImageXObject = 85, paintJpegXObject = 82
        if (fn === 85 || fn === 82) {
          var name = ops.argsArray[i][0];
          try {
            var imgObj = await new Promise(function (res) {
              try { page.objs.get(name, res); } catch (e) { res(null); }
              setTimeout(function () { res(null); }, 800);
            });
            if (imgObj && imgObj.data) {
              var c = document.createElement("canvas");
              c.width = imgObj.width; c.height = imgObj.height;
              var ctx = c.getContext("2d");
              var id = ctx.createImageData(imgObj.width, imgObj.height);
              id.data.set(imgObj.data);
              ctx.putImageData(id, 0, 0);
              var blob = await new Promise(function (res) { c.toBlob(res, "image/png"); });
              results.push({
                name: stem + "_p" + p + "_图" + (++imgIndex) + ".png",
                data: new Uint8Array(await blob.arrayBuffer()),
                page: p,
              });
            }
          } catch (e) { /* 单张失败跳过 */ }
        }
      }
    }
    return { pages: doc.numPages, results: results };
  }

  // ==================== 图片高级功能 ====================

  /** 二维码识别：传入 canvas，内部取像素扫描（兼容 jsQR 全局名差异） */
  async function scanQR(canvas, jsQR) {
    var lib = jsQR || (typeof window !== "undefined" && (window.jsQR || window.default)) || null;
    if (typeof lib !== "function") throw new Error("二维码识别库未加载");
    var ctx = canvas.getContext("2d");
    var imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    var code = lib(imageData.data, imageData.width, imageData.height);
    if (!code || !code.data) return { found: false };
    return { found: true, text: code.data };
  }

  /** 多图压缩（browser-image-compression，逐张处理） */
  async function compressImages(items, opts) {
    var fn = (typeof window !== "undefined" && (window.imageCompression || window.default)) || null;
    if (typeof fn !== "function") throw new Error("压缩库未加载");
    opts = opts || {};
    var results = [];
    for (var i = 0; i < items.length; i++) {
      try {
        var it = items[i];
        var file = new File([it.data], it.name, { type: it.type || "image/jpeg" });
        var out = await fn(file, {
          maxSizeMB: opts.maxSizeMB || 0.2,
          maxWidthOrHeight: opts.maxSide || 1920,
          useWebWorker: true,
        });
        results.push({
          name: it.name.replace(/\.[^.]+$/, "") + ".jpg",
          data: new Uint8Array(await out.arrayBuffer()),
          ok: true, before: file.size, after: out.size,
        });
      } catch (e) {
        results.push({ name: items[i].name, ok: false, reason: String(e.message || e) });
      }
    }
    return results;
  }

  /** 图片拼接（横向/纵向/网格） */
  async function stitchImages(items, opts) {
    opts = opts || {};
    var mode = opts.mode || "vertical";  // vertical | horizontal | grid
    var gap = opts.gap || 0;
    var bg = opts.bg || "#ffffff";
    var cols = Math.max(1, opts.cols || 3);

    var images = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var u8 = buf(it.data);
      var blob = new Blob([u8], { type: it.type || "image/*" });
      var bmp = await createImageBitmap(blob).catch(function () { return null; });
      if (!bmp) throw new Error("第 " + (i + 1) + " 张图片无法读取");
      images.push(bmp);
    }
    if (!images.length) throw new Error("没有可拼接的图片");

    var W = 0, H = 0;
    if (mode === "vertical") {
      images.forEach(function (b) { W = Math.max(W, b.width); H += b.height; });
      H += gap * (images.length - 1);
    } else if (mode === "horizontal") {
      images.forEach(function (b) { H = Math.max(H, b.height); W += b.width; });
      W += gap * (images.length - 1);
    } else {
      // 网格：按最大宽度统一列宽
      var maxW = 0;
      images.forEach(function (b) { maxW = Math.max(maxW, b.width); });
      var rows = Math.ceil(images.length / cols);
      var maxH = 0;
      images.forEach(function (b) { maxH = Math.max(maxH, b.height); });
      W = maxW * cols + gap * (cols - 1);
      H = maxH * rows + gap * (rows - 1);
    }

    var c = document.createElement("canvas");
    c.width = W; c.height = H;
    var ctx = c.getContext("2d");
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, W, H);

    if (mode === "vertical") {
      var y = 0;
      images.forEach(function (b) {
        ctx.drawImage(b, Math.round((W - b.width) / 2), y);
        y += b.height + gap;
      });
    } else if (mode === "horizontal") {
      var x = 0;
      images.forEach(function (b) {
        ctx.drawImage(b, x, Math.round((H - b.height) / 2));
        x += b.width + gap;
      });
    } else {
      var maxW2 = 0, maxH2 = 0;
      images.forEach(function (b) { maxW2 = Math.max(maxW2, b.width); maxH2 = Math.max(maxH2, b.height); });
      images.forEach(function (b, idx) {
        var r = Math.floor(idx / cols), cc = idx % cols;
        var cx = cc * (maxW2 + gap) + Math.round((maxW2 - b.width) / 2);
        var cy = r * (maxH2 + gap) + Math.round((maxH2 - b.height) / 2);
        ctx.drawImage(b, cx, cy);
      });
    }

    var outBlob = await new Promise(function (res) { c.toBlob(res, "image/png"); });
    return {
      data: new Uint8Array(await outBlob.arrayBuffer()),
      name: "拼接图.png", w: W, h: H,
    };
  }

  /** GIF 拆分（逐帧导出 PNG） */
  async function splitGif(item) {
    // 用 ImageDecoder（现代浏览器支持）
    if (typeof ImageDecoder === "undefined") {
      throw new Error("当前浏览器不支持 GIF 拆分（需要较新版 Chrome/Edge）");
    }
    var u8 = buf(item.data);
    var decoder = new ImageDecoder({ data: u8, type: "image/gif" });
    await decoder.tracks.ready;
    var track = decoder.tracks.selectedTrack;
    var frameCount = track.frameCount;
    var stem = safeName(item.name.replace(/\.[^.]+$/, ""));
    var results = [];
    for (var i = 0; i < frameCount; i++) {
      var res = await decoder.decode({ frameIndex: i });
      var img = res.image;
      var c = document.createElement("canvas");
      c.width = img.displayWidth;
      c.height = img.displayHeight;
      c.getContext("2d").drawImage(img, 0, 0);
      var blob = await new Promise(function (r) { c.toBlob(r, "image/png"); });
      results.push({
        name: stem + "_第" + (i + 1) + "帧.png",
        data: new Uint8Array(await blob.arrayBuffer()),
        frame: i + 1,
      });
    }
    decoder.close();
    return { total: frameCount, results: results };
  }

  /** 图片加水印（批量，文字/图片水印都支持） */
  async function batchWatermark(items, opts, logoBitmap) {
    opts = opts || {};
    var results = [];
    var text = opts.text || "";
    var pos = opts.pos || "se";
    var opacity = Math.max(0.03, Math.min(1, (opts.opacity || 45) / 100));
    var scale = opts.scale || 18;

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var u8 = buf(it.data);
      var blob = new Blob([u8], { type: it.type || "image/*" });
      var bmp = await createImageBitmap(blob).catch(function () { return null; });
      if (!bmp) { results.push({ name: it.name, ok: false, reason: "无法读取" }); continue; }

      var c = document.createElement("canvas");
      c.width = bmp.width;
      c.height = bmp.height;
      var ctx = c.getContext("2d");
      ctx.drawImage(bmp, 0, 0);

      var m = Math.round(Math.min(c.width, c.height) * 0.03);
      if (logoBitmap) {
        var lw = Math.max(8, Math.round(c.width * scale / 100));
        var lh = Math.round(logoBitmap.height * (lw / logoBitmap.width));
        var lx = pos[1] === "w" ? m : pos[1] === "e" ? c.width - lw - m : Math.round((c.width - lw) / 2);
        var ly = pos[0] === "n" ? m : pos[0] === "s" ? c.height - lh - m : Math.round((c.height - lh) / 2);
        ctx.save();
        ctx.globalAlpha = opacity;
        ctx.drawImage(logoBitmap, lx, ly, lw, lh);
        ctx.restore();
      } else if (text.trim()) {
        var fs = Math.max(10, Math.round(Math.min(c.width, c.height) * 0.055));
        ctx.font = fs + 'px "Microsoft YaHei","PingFang SC",sans-serif';
        ctx.textBaseline = "middle";
        ctx.globalAlpha = opacity;
        var tw = ctx.measureText(text).width;
        var cx, cy, rot = 0;
        if (pos === "center") { cx = c.width / 2; cy = c.height / 2; rot = -Math.PI / 6; }
        else {
          cx = pos[1] === "w" ? m + tw / 2 : pos[1] === "e" ? c.width - m - tw / 2 : c.width / 2;
          cy = pos[0] === "n" ? m + fs / 2 : pos[0] === "s" ? c.height - m - fs / 2 : c.height / 2;
        }
        ctx.save();
        ctx.translate(cx, cy);
        ctx.rotate(rot);
        ctx.fillStyle = opts.color || "#ffffff";
        ctx.strokeStyle = "rgba(0,0,0,.35)";
        ctx.lineWidth = Math.max(2, fs * 0.07);
        ctx.strokeText(text, -tw / 2, 0);
        ctx.fillText(text, -tw / 2, 0);
        ctx.restore();
        ctx.globalAlpha = 1;
      }

      var type = /png/i.test(it.type || "") ? "image/png" : "image/jpeg";
      var outBlob = await new Promise(function (r) { c.toBlob(r, type, 0.92); });
      results.push({
        name: it.name.replace(/\.[^.]+$/, "") + (type === "image/png" ? ".png" : ".jpg"),
        data: new Uint8Array(await outBlob.arrayBuffer()),
        ok: true,
      });
    }
    return results;
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

  // ==================== Excel 数据处理 ====================
  /** 清理：去空行空列、去重、去首尾空格 */
  async function xlClean(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var before = r.rows.length;

    var rows = r.rows;
    if (opts.trim !== false) {
      rows = rows.map(function (row) {
        return row.map(function (v) { return typeof v === "string" ? v.trim() : v; });
      });
      r.header = r.header.map(function (h) { return String(h).trim(); });
    }
    if (opts.dropEmpty !== false) {
      rows = rows.filter(function (row) {
        return !row.every(function (v) { return v === "" || v == null; });
      });
    }
    if (opts.dropEmptyCols !== false && r.header.length) {
      var keepCols = [];
      for (var c = 0; c < r.header.length; c++) {
        var has = rows.some(function (row) {
          return String(row[c] == null ? "" : row[c]).trim() !== "";
        });
        if (has) keepCols.push(c);
      }
      r.header = keepCols.map(function (i) { return r.header[i]; });
      rows = rows.map(function (row) { return keepCols.map(function (i) { return row[i] == null ? "" : row[i]; }); });
    }
    var removed = 0;
    if (opts.dedup) {
      var seen = {}, kept = [];
      rows.forEach(function (row) {
        var k = row.join("\u0001");
        if (seen[k]) { removed++; return; }
        seen[k] = 1;
        kept.push(row);
      });
      rows = kept;
    }
    var wbOut = aoaToWb([r.header].concat(rows), "清理后");
    return { data: wbToBuf(wbOut), before: before, after: rows.length, removed: removed };
  }

  /** 排序：按指定列升/降序 */
  async function xlSort(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var ki = r.header.indexOf(opts.key);
    if (ki < 0) throw new Error("找不到列：" + opts.key);
    var dir = opts.desc ? -1 : 1;
    var rows = r.rows.slice().sort(function (a, b) {
      var x = a[ki], y = b[ki];
      var nx = parseFloat(x), ny = parseFloat(y);
      if (!isNaN(nx) && !isNaN(ny)) return (nx - ny) * dir;
      x = String(x == null ? "" : x);
      y = String(y == null ? "" : y);
      return (x < y ? -1 : x > y ? 1 : 0) * dir;
    });
    var wbOut = aoaToWb([r.header].concat(rows), "排序后");
    return { data: wbToBuf(wbOut), count: rows.length };
  }

  /** 筛选：等于 / 包含 / 大于 / 小于 / 非空 / 空 */
  async function xlFilter(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var ki = r.header.indexOf(opts.key);
    if (ki < 0) throw new Error("找不到列：" + opts.key);
    var cond = opts.cond || "contains";
    var val = (opts.val == null ? "" : String(opts.val)).trim();
    var nval = parseFloat(val);

    var rows = r.rows.filter(function (row) {
      var v = row[ki];
      var s = String(v == null ? "" : v).trim();
      var n = parseFloat(v);
      switch (cond) {
        case "eq": return s === val;
        case "gt": return !isNaN(n) && !isNaN(nval) && n > nval;
        case "lt": return !isNaN(n) && !isNaN(nval) && n < nval;
        case "notempty": return s !== "";
        case "empty": return s === "";
        default: return val === "" || s.indexOf(val) >= 0;
      }
    });
    var wbOut = aoaToWb([r.header].concat(rows), "筛选结果");
    return { data: wbToBuf(wbOut), count: rows.length, total: r.rows.length };
  }

  /** 统计汇总：按某列分组，对其它列求 和/均值/最大/最小 */
  async function xlStats(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    if (!r.header.length) throw new Error("空表");
    var keyIdx = opts.groupBy >= 0 ? opts.groupBy : -1;
    var groups = {};

    r.rows.forEach(function (row) {
      if (row.every(function (v) { return v === "" || v == null; })) return;
      var gk = keyIdx >= 0
        ? (String(row[keyIdx] == null ? "" : row[keyIdx]).trim() || "(空)")
        : "全部数据";
      if (!groups[gk]) {
        groups[gk] = { count: 0, cols: r.header.map(function () { return { nums: [] } }) };
      }
      var g = groups[gk];
      g.count++;
      r.header.forEach(function (h, i) {
        if (i === keyIdx) return;
        var n = parseFloat(row[i]);
        if (!isNaN(n)) g.cols[i].nums.push(n);
      });
    });

    var header = ["分组", "行数"];
    r.header.forEach(function (h, i) {
      if (i === keyIdx) return;
      header.push(h + " 求和", h + " 平均", h + " 最大", h + " 最小");
    });

    var out = [];
    Object.keys(groups).forEach(function (gk) {
      var g = groups[gk];
      var row = [gk, g.count];
      r.header.forEach(function (h, i) {
        if (i === keyIdx) return;
        var ns = g.cols[i].nums;
        if (!ns.length) { row.push("", "", "", ""); return; }
        var sum = ns.reduce(function (a, b) { return a + b; }, 0);
        row.push(+sum.toFixed(4), +(sum / ns.length).toFixed(4),
          Math.max.apply(null, ns), Math.min.apply(null, ns));
      });
      out.push(row);
    });
    var wbOut = aoaToWb([header].concat(out), "统计汇总");
    return { data: wbToBuf(wbOut), groups: Object.keys(groups).length };
  }

  /** 拆分列：按分隔符把一列拆成多列 */
  async function xlSplitColumn(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var ki = r.header.indexOf(opts.key);
    if (ki < 0) throw new Error("找不到列：" + opts.key);
    var sep = opts.sep || ",";
    var maxParts = 0;
    var splits = r.rows.map(function (row) {
      var s = String(row[ki] == null ? "" : row[ki]);
      var parts = sep === "" ? s.split("") : s.split(sep);
      maxParts = Math.max(maxParts, parts.length);
      return parts;
    });
    if (maxParts <= 1) throw new Error("用「" + sep + "」拆不出多列，换个分隔符试试");

    var header = r.header.slice(0, ki);
    for (var i = 0; i < maxParts; i++) header.push(r.header[ki] + "-" + (i + 1));
    header = header.concat(r.header.slice(ki + 1));

    var rows = r.rows.map(function (row, ri) {
      var parts = splits[ri].slice();
      for (var j = parts.length; j < maxParts; j++) parts.push("");
      return row.slice(0, ki).concat(parts, row.slice(ki + 1));
    });
    var wbOut = aoaToWb([header].concat(rows), "拆分列");
    return { data: wbToBuf(wbOut), cols: maxParts };
  }

  /** 合并列：用连接符把多列拼成一列 */
  async function xlMergeColumns(item, opts) {
    opts = opts || {};
    var wb = readWb(item.data);
    var r = sheetRows(wb);
    var keys = (opts.keys || []).map(function (h) { return r.header.indexOf(h); });
    if (keys.length < 2) throw new Error("请至少选 2 列");
    keys.forEach(function (i) { if (i < 0) throw new Error("有列不存在"); });
    var sep = opts.sep == null ? "-" : opts.sep;
    var newName = opts.name || "合并列";

    var header = r.header.slice();
    header.push(newName);
    var rows = r.rows.map(function (row) {
      var v = keys.map(function (i) {
        return String(row[i] == null ? "" : row[i]).trim();
      }).join(sep);
      return row.concat([v]);
    });
    var wbOut = aoaToWb([header].concat(rows), "合并列");
    return { data: wbToBuf(wbOut), cols: keys.length };
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
  /**
   * 把 docx 文档打包成 ArrayBuffer。
   * 优先 toBuffer（Node 与新版 docx 都有，且不依赖 Blob 实现）；
   * 退化到 toBlob + arrayBuffer。
   * 带超时：某些环境（如 jsdom）Blob 实现不完整会静默挂起，
   * 没有超时时用户点完按钮没有任何反馈，按钮也不恢复。
   */
  async function packDocx(doc, ms) {
    var P = docxLib.Packer;
    var deadline = ms || 60000;
    var out;
    if (typeof P.toBuffer === "function") {
      out = await withTimeout(P.toBuffer(doc), deadline, "Word 打包超时");
      return new Uint8Array(out.buffer ? out.buffer : out);
    }
    if (typeof P.toBlob === "function") {
      var blob = await withTimeout(P.toBlob(doc), deadline, "Word 打包超时");
      var ab = await withTimeout(blob.arrayBuffer(), deadline, "Word 打包超时");
      return new Uint8Array(ab);
    }
    throw new Error("Word 生成库缺少打包方法");
  }
  function withTimeout(p, ms, msg) {
    return Promise.race([
      Promise.resolve(p),
      new Promise(function (_, rej) { setTimeout(function () { rej(new Error(msg || "操作超时")); }, ms); }),
    ]);
  }

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
    var ab = await packDocx(doc);
    return { data: ab, name: item.name.replace(/\.[^.]+$/, "") + ".docx" };
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
    var ab = await packDocx(doc);
    return { data: ab, name: item.name.replace(/\.[^.]+$/, "") + ".docx" };
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
    // 带超时：JSZip 的 DEFLATE 依赖异步调度（setImmediate 等），
    // 在缺少这些 API 的环境里 promise 永不 settle。没超时会让按钮永久禁用。
    return await withTimeout(
      zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" }),
      45000,
      "打包超时，请重试或改用单个下载"
    );
  }

  return {
    safeName: safeName,
    parseRanges: parseRanges,
    // PDF
    pdfMerge: pdfMerge, pdfSplit: pdfSplit, pdfPageOps: pdfPageOps,
    pdfWatermark: pdfWatermark, pdfResize: pdfResize, pdfInfo: pdfInfo,
    pdfToImages: pdfToImages, pdfToText: pdfToText, pdfToLongImage: pdfToLongImage,
    pdfExtractImages: pdfExtractImages,
    imagesToPdf: imagesToPdf,
    // 图片高级
    scanQR: scanQR, compressImages: compressImages, stitchImages: stitchImages,
    splitGif: splitGif, batchWatermark: batchWatermark,
    // Excel
    xlMerge: xlMerge, xlSplit: xlSplit, xlCompare: xlCompare, xlHeaders: xlHeaders,
    xlToCsv: xlToCsv, csvToXl: csvToXl, xlToHtml: xlToHtml, xlToText: xlToText,
    xlClean: xlClean, xlSort: xlSort, xlFilter: xlFilter, xlStats: xlStats,
    xlSplitColumn: xlSplitColumn, xlMergeColumns: xlMergeColumns,
    // Word
    wordToHtml: wordToHtml, wordToText: wordToText, wordToHtmlFile: wordToHtmlFile,
    textToWord: textToWord, htmlToWord: htmlToWord,
    // 通用
    zipFiles: zipFiles,
  };
});
