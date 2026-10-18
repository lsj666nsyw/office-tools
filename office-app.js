/* office-app.js —— 全能办公工具箱前端 */
(function () {
  "use strict";
  var TB = createOfficeToolbox({
    PDFLib: window.PDFLib, XLSX: window.XLSX, JSZip: window.JSZip,
    mammoth: window.mammoth, docxLib: window.docx,
  });
  var $ = function (id) { return document.getElementById(id); };
  var docxLib = window.docx;

  function fmt(n) {
    var u = ["B", "KB", "MB", "GB"], i = 0;
    while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
    return (i === 0 ? n : n.toFixed(1)) + " " + u[i];
  }
  function log(box, msg, cls) {
    var d = document.createElement("div");
    if (cls) d.className = cls;
    d.textContent = msg;
    box.appendChild(d);
    box.scrollTop = box.scrollHeight;
  }
  function clear(el) { el.innerHTML = ""; }
  function prog(box, r) {
    box.style.display = r == null ? "none" : "block";
    if (r != null) box.firstElementChild.style.width = Math.round(r * 100) + "%";
  }
  function download(name, data, type) {
    var blob = data instanceof Blob ? data : new Blob([data], { type: type || "application/octet-stream" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function resultRow(box, name, size, data, type) {
    var row = document.createElement("div");
    row.className = "row";
    var nm = document.createElement("span"); nm.className = "nm"; nm.textContent = name;
    var sz = document.createElement("span"); sz.className = "sz"; sz.textContent = fmt(size);
    var a = document.createElement("a"); a.textContent = "下载";
    a.onclick = function () { download(name, data, type); };
    row.appendChild(nm); row.appendChild(sz); row.appendChild(a);
    box.appendChild(row);
  }
  async function bigZip(box, files, zipName) {
    var buf = await TB.zipFiles(files);
    var a = document.createElement("a");
    a.className = "bigdl";
    a.textContent = "打包下载 " + files.length + " 个文件（" + fmt(buf.byteLength) + "）";
    a.onclick = function () { download(zipName, buf, "application/zip"); };
    box.appendChild(a);
  }
  function readBuf(f) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(r.result); };
      r.onerror = function () { rej(new Error("读取失败")); };
      r.readAsArrayBuffer(f);
    });
  }
  function readText(f) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(r.result); };
      r.onerror = function () { rej(new Error("读取失败")); };
      r.readAsText(f, "utf-8");
    });
  }

  // ===================== 文件选择器 =====================
  function picker(dropId, inputId, listId, accept, multi) {
    var drop = $(dropId), input = $(inputId), list = $(listId);
    var files = [];
    function render() {
      clear(list);
      files.forEach(function (f, i) {
        var d = document.createElement("div");
        d.className = "fitem";
        var ord = document.createElement("span"); ord.className = "ord"; ord.textContent = i + 1;
        var nm = document.createElement("span"); nm.className = "nm"; nm.textContent = f.name;
        var sz = document.createElement("span"); sz.className = "sz"; sz.textContent = fmt(f.size);
        d.appendChild(ord); d.appendChild(nm); d.appendChild(sz);
        if (multi) {
          var up = document.createElement("span"); up.className = "mv"; up.textContent = i ? "↑" : "";
          up.onclick = function (e) {
            e.stopPropagation();
            if (i > 0) { var t = files[i - 1]; files[i - 1] = files[i]; files[i] = t; render(); }
          };
          d.appendChild(up);
        }
        var del = document.createElement("button"); del.textContent = "×";
        del.onclick = function (e) { e.stopPropagation(); files.splice(i, 1); render(); };
        d.appendChild(del);
        list.appendChild(d);
      });
    }
    function add(fl) {
      Array.prototype.forEach.call(fl, function (f) {
        if (accept && !accept.test(f.name)) return;
        if (files.some(function (x) { return x.name === f.name && x.size === f.size; })) return;
        if (!multi) files = [];
        files.push(f);
      });
      render();
      if (typeof input.onchange === "function" && input.dataset.hook) {
        // 交给外部钩子
      }
    }
    drop.onclick = function () { input.click(); };
    input.addEventListener("change", function () {
      add(input.files); input.value = "";
      if (input._after) input._after();
    });
    ["dragover", "dragenter"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("over"); });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("over"); });
    });
    drop.addEventListener("drop", function (e) {
      if (e.dataTransfer) { add(e.dataTransfer.files); if (input._after) input._after(); }
    });
    return {
      get files() { return files; },
      get first() { return files[0]; },
      clear: function () { files = []; render(); },
      onAfter: function (fn) { input._after = fn; },
    };
  }

  // ===================== 选项卡 =====================
  document.querySelector("nav").addEventListener("click", function (e) {
    var b = e.target.closest("button");
    if (!b) return;
    document.querySelectorAll("nav button").forEach(function (x) { x.classList.remove("on"); });
    document.querySelectorAll("section").forEach(function (x) { x.classList.remove("on"); });
    b.classList.add("on");
    $("s-" + b.dataset.t).classList.add("on");
  });

  function seg(id, cb) {
    var box = $(id);
    box.addEventListener("click", function (e) {
      var b = e.target.closest("button");
      if (!b) return;
      box.querySelectorAll("button").forEach(function (x) { x.classList.remove("on"); });
      b.classList.add("on");
      if (cb) cb(b.dataset.v);
    });
    return { get value() { return box.querySelector(".on").dataset.v; } };
  }

  // ==================================================================
  // 图片
  // ==================================================================
  var iPick = picker("i-drop", "i-files", "i-list", /^image\//, true);
  var iMode = { v: "none" };
  seg("i-mode", function (v) {
    iMode.v = v;
    var box = $("i-args"); clear(box);
    function inp(id, val, w) {
      var i = document.createElement("input"); i.type = "number"; i.id = id;
      i.value = val; i.style.width = w || "70px"; return i;
    }
    function lab(t) {
      var s = document.createElement("span");
      s.style.cssText = "font-size:12.5px;color:var(--ink3)"; s.textContent = t; return s;
    }
    if (v === "long") { box.appendChild(inp("i-long", 1600)); box.appendChild(lab("px 长边")); }
    if (v === "fix") {
      box.appendChild(inp("i-w", 800, "62px")); box.appendChild(lab("×"));
      box.appendChild(inp("i-h", 800, "62px")); box.appendChild(lab("等比补边"));
    }
    if (v === "pct") { box.appendChild(inp("i-pct", 50, "60px")); box.appendChild(lab("%")); }
  });
  $("i-q").oninput = function () { $("i-qv").textContent = this.value; };

  var logoImg = null;
  $("i-logobtn").onclick = function () { $("i-logofile").click(); };
  $("i-logofile").onchange = function () {
    var f = this.files[0];
    if (!f) return;
    var url = URL.createObjectURL(f);
    var im = new Image();
    im.onload = function () {
      logoImg = im;
      $("i-logolabel").textContent = f.name;
      $("i-logolabel").style.color = "#2fa360";
      $("i-logo").checked = true;
    };
    im.onerror = function () { $("i-logolabel").textContent = "打不开"; };
    im.src = url;
  };

  function loadImage(file) {
    return new Promise(function (res, rej) {
      var url = URL.createObjectURL(file);
      var im = new Image();
      im.onload = function () { res({ img: im, url: url }); };
      im.onerror = function () { URL.revokeObjectURL(url); rej(new Error("图片打不开")); };
      im.src = url;
    });
  }
  function trimEdges(c, tol) {
    var w = c.width, h = c.height;
    if (w < 8 || h < 8) return c;
    var ctx = c.getContext("2d");
    var d = ctx.getImageData(0, 0, w, h).data;
    var t = Math.max(1, Math.round(tol * 2.55));
    function px(x, y) { var i = (y * w + x) * 4; return [d[i], d[i + 1], d[i + 2]]; }
    var cs = [px(0, 0), px(w - 1, 0), px(0, h - 1), px(w - 1, h - 1)];
    var bg = [0, 1, 2].map(function (k) {
      return Math.round(cs.reduce(function (s, x) { return s + x[k]; }, 0) / 4);
    });
    function isBg(x, y) {
      var i = (y * w + x) * 4;
      return Math.abs(d[i] - bg[0]) <= t && Math.abs(d[i + 1] - bg[1]) <= t && Math.abs(d[i + 2] - bg[2]) <= t;
    }
    function rowBg(y) { var n = 0; for (var x = 0; x < w; x++) if (isBg(x, y)) n++; return n >= w * 0.98; }
    function colBg(x) { var n = 0; for (var y = 0; y < h; y++) if (isBg(x, y)) n++; return n >= h * 0.98; }
    var top = 0; while (top < h && rowBg(top)) top++;
    var bot = h; while (bot > top && rowBg(bot - 1)) bot--;
    var left = 0; while (left < w && colBg(left)) left++;
    var right = w; while (right > left && colBg(right - 1)) right--;
    if (top === 0 && bot === h && left === 0 && right === w) return c;
    var nw = right - left, nh = bot - top;
    if (nw < 8 || nh < 8 || nw < w * 0.2 || nh < h * 0.2) return c;
    var out = document.createElement("canvas");
    out.width = nw; out.height = nh;
    out.getContext("2d").drawImage(c, left, top, nw, nh, 0, 0, nw, nh);
    return out;
  }
  function removeBg(c, tol) {
    var w = c.width, h = c.height, ctx = c.getContext("2d");
    var img = ctx.getImageData(0, 0, w, h), d = img.data;
    var t = Math.max(1, Math.round(tol * 2.55));
    function px(x, y) { var i = (y * w + x) * 4; return [d[i], d[i + 1], d[i + 2]]; }
    var cs = [px(0, 0), px(w - 1, 0), px(0, h - 1), px(w - 1, h - 1)];
    var bg = [0, 1, 2].map(function (k) {
      return Math.round(cs.reduce(function (s, x) { return s + x[k]; }, 0) / 4);
    });
    for (var i = 0; i < d.length; i += 4) {
      if (Math.abs(d[i] - bg[0]) <= t && Math.abs(d[i + 1] - bg[1]) <= t && Math.abs(d[i + 2] - bg[2]) <= t) d[i + 3] = 0;
    }
    ctx.putImageData(img, 0, 0);
    return c;
  }
  function addBorder(c, permille, color) {
    var w = c.width, h = c.height;
    var p = Math.max(1, Math.round(Math.min(w, h) * permille / 1000));
    if (p * 2 >= Math.min(w, h)) return c;
    var out = document.createElement("canvas");
    out.width = w + p * 2; out.height = h + p * 2;
    var x = out.getContext("2d");
    x.fillStyle = color; x.fillRect(0, 0, out.width, out.height);
    x.drawImage(c, p, p);
    return out;
  }
  function drawLogo(c, img, pos, scalePct, opacity) {
    var w = c.width, h = c.height;
    var lw = Math.max(8, Math.round(w * scalePct / 100));
    var lh = Math.round(img.naturalHeight * (lw / img.naturalWidth));
    var m = Math.round(Math.min(w, h) * 0.03);
    var x0 = pos[1] === "w" ? m : pos[1] === "e" ? w - lw - m : Math.round((w - lw) / 2);
    var y0 = pos[0] === "n" ? m : pos[0] === "s" ? h - lh - m : Math.round((h - lh) / 2);
    var ctx = c.getContext("2d");
    ctx.save();
    ctx.globalAlpha = Math.max(0.05, Math.min(1, opacity / 100));
    ctx.drawImage(img, x0, y0, lw, lh);
    ctx.restore();
    return c;
  }
  function drawText(c, text, pos) {
    var ctx = c.getContext("2d");
    var fs = Math.max(10, Math.round(Math.min(c.width, c.height) * 0.055));
    ctx.font = fs + 'px "Microsoft YaHei", "PingFang SC", sans-serif';
    ctx.textBaseline = "middle";
    ctx.globalAlpha = 0.45;
    var tw = ctx.measureText(text).width;
    var m = Math.round(Math.min(c.width, c.height) * 0.04);
    var cx, cy, rot = 0;
    if (pos === "center") { cx = c.width / 2; cy = c.height / 2; rot = -Math.PI / 6; }
    else {
      cx = pos[1] === "w" ? m + tw / 2 : pos[1] === "e" ? c.width - m - tw / 2 : c.width / 2;
      cy = pos[0] === "n" ? m + fs / 2 : pos[0] === "s" ? c.height - m - fs / 2 : c.height / 2;
    }
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(rot);
    ctx.fillStyle = "#fff";
    ctx.strokeStyle = "rgba(0,0,0,.35)";
    ctx.lineWidth = Math.max(2, fs * 0.07);
    ctx.strokeText(text, -tw / 2, 0);
    ctx.fillText(text, -tw / 2, 0);
    ctx.restore();
    ctx.globalAlpha = 1;
    return c;
  }
  function canvasToBlob(c, type, q) {
    return new Promise(function (res, rej) {
      c.toBlob(function (b) { b ? res(b) : rej(new Error("导出失败")); }, type, q);
    });
  }
  async function compressTo(c, type, startQ, target) {
    var lo = 5, hi = Math.round(startQ * 100), best = null;
    for (var i = 0; i < 8; i++) {
      var mid = Math.round((lo + hi) / 2);
      var b = await canvasToBlob(c, type, mid / 100);
      if (b.size <= target) { best = b; lo = mid + 1; if (b.size >= target * 0.85) break; }
      else hi = mid - 1;
      if (lo > hi) break;
    }
    return best || await canvasToBlob(c, type, 0.05);
  }

  $("i-run").onclick = async function () {
    var box = $("i-log"), res = $("i-res"), pv = $("i-pv");
    clear(box); clear(res); clear(pv);
    if (!iPick.files.length) { log(box, "请先选择图片", "err"); return; }
    var opts = {
      mode: iMode.v,
      long: +(($("i-long") || {}).value || 1600),
      fw: +(($("i-w") || {}).value || 800),
      fh: +(($("i-h") || {}).value || 800),
      pct: +(($("i-pct") || {}).value || 50),
      sq: $("i-sq").checked, trim: $("i-trim").checked, cut: $("i-cut").checked,
      wm: $("i-wm").checked, wmt: $("i-wmt").value, wmp: $("i-wmp").value,
      logo: $("i-logo").checked,
      bd: $("i-bd").checked, bdw: +$("i-bdw").value || 4, bdc: $("i-bdc").value,
      fmt: $("i-fmt").value, q: +$("i-q").value,
      limit: $("i-lim").checked ? (+$("i-kb").value || 200) * 1024 : 0,
    };
    if (opts.wm && !opts.wmt.trim()) { log(box, "勾了水印但没填文字", "err"); return; }
    if (opts.logo && !logoImg) { log(box, "勾了 Logo 水印但没选文件", "err"); return; }
    if (opts.cut && opts.fmt === "image/jpeg") { log(box, "抠图需透明，输出不能是 JPG", "err"); return; }
    this.disabled = true; prog($("i-prog"), 0);
    var outs = [], tin = 0, tout = 0;
    for (var i = 0; i < iPick.files.length; i++) {
      var f = iPick.files[i];
      try {
        var r = await loadImage(f);
        var iw = r.img.naturalWidth, ih = r.img.naturalHeight, w = iw, h = ih;
        if (opts.mode === "long") {
          var sc = opts.long / Math.max(iw, ih);
          if (sc < 1) { w = Math.round(iw * sc); h = Math.round(ih * sc); }
        } else if (opts.mode === "pct") {
          w = Math.max(1, Math.round(iw * opts.pct / 100));
          h = Math.max(1, Math.round(ih * opts.pct / 100));
        }
        var c;
        if (opts.mode === "fix") {
          c = document.createElement("canvas");
          c.width = opts.fw; c.height = opts.fh;
          var cx = c.getContext("2d");
          cx.fillStyle = "#fff"; cx.fillRect(0, 0, opts.fw, opts.fh);
          var s2 = Math.min(opts.fw / iw, opts.fh / ih);
          var nw = Math.round(iw * s2), nh = Math.round(ih * s2);
          cx.drawImage(r.img, (opts.fw - nw) / 2, (opts.fh - nh) / 2, nw, nh);
        } else {
          c = document.createElement("canvas");
          c.width = w; c.height = h;
          var cx2 = c.getContext("2d");
          cx2.fillStyle = "#fff"; cx2.fillRect(0, 0, w, h);
          cx2.drawImage(r.img, 0, 0, w, h);
        }
        URL.revokeObjectURL(r.url);
        if (opts.sq) {
          var side = Math.max(c.width, c.height);
          var sq = document.createElement("canvas");
          sq.width = side; sq.height = side;
          var sx = sq.getContext("2d");
          sx.fillStyle = "#fff"; sx.fillRect(0, 0, side, side);
          sx.drawImage(c, (side - c.width) / 2, (side - c.height) / 2);
          c = sq;
        }
        if (opts.trim) c = trimEdges(c, 12);
        if (opts.cut) c = removeBg(c, 30);
        if (opts.bd) c = addBorder(c, opts.bdw, opts.bdc);
        if (opts.logo && logoImg) c = drawLogo(c, logoImg, "se", 18, 60);
        if (opts.wm && opts.wmt.trim()) c = drawText(c, opts.wmt, opts.wmp);

        var type = opts.fmt === "keep" ? (f.type === "image/png" ? "image/png" : "image/jpeg") : opts.fmt;
        if (opts.cut && type === "image/jpeg") type = "image/png";
        var blob = (opts.limit && type !== "image/png")
          ? await compressTo(c, type, opts.q / 100, opts.limit)
          : await canvasToBlob(c, type, opts.q / 100);
        tin += f.size; tout += blob.size;
        var ext = type === "image/png" ? ".png" : type === "image/webp" ? ".webp" : ".jpg";
        var nm = f.name.replace(/\.[^.]+$/, "") + ext;
        outs.push({ name: nm, data: blob });
        var fig = document.createElement("figure");
        var im = document.createElement("img");
        im.src = URL.createObjectURL(blob);
        var cap = document.createElement("figcaption");
        cap.textContent = fmt(blob.size);
        fig.appendChild(im); fig.appendChild(cap);
        pv.appendChild(fig);
        log(box, "✓ " + f.name + "  " + fmt(f.size) + " → " + fmt(blob.size), "ok");
      } catch (e) { log(box, "✗ " + f.name + "：" + (e.message || e), "err"); }
      prog($("i-prog"), (i + 1) / iPick.files.length);
      if (i % 3 === 2) await new Promise(function (r2) { setTimeout(r2, 0); });
    }
    if (outs.length === 1) resultRow(res, outs[0].name, outs[0].data.size, outs[0].data, outs[0].data.type);
    else if (outs.length > 1) await bigZip(res, outs, "图片处理结果.zip");
    if (outs.length) {
      var saved = tin - tout;
      log(box, "", null);
      log(box, "完成 " + outs.length + " 张" + (saved > 0 ? "，共省 " + fmt(saved) : ""), "ok");
    }
    this.disabled = false; prog($("i-prog"), null);
  };

  // ==================================================================
  // PDF
  // ==================================================================
  var pPick = picker("p-drop", "p-files", "p-list", /\.pdf$/i, true);
  pPick.onAfter(async function () {
    var f = pPick.first, info = $("p-info");
    if (!f) { info.textContent = ""; return; }
    try {
      var doc = await PDFLib.PDFDocument.load(await readBuf(f), { ignoreEncryption: true });
      info.textContent = "第一个文件：" + f.name + "，共 " + doc.getPageCount() + " 页";
    } catch (e) { info.textContent = "读取失败：" + e.message; }
  });
  var pMode = seg("p-mode", renderPdfArgs);
  function renderPdfArgs(v) {
    var box = $("p-args"); clear(box);
    function lab(t) {
      var s = document.createElement("label"); s.textContent = t; return s;
    }
    function inp(id, type, val, w, ph) {
      var i = document.createElement("input"); i.type = type; i.id = id;
      i.value = val; i.style.width = w || "80px"; if (ph) i.placeholder = ph;
      return i;
    }
    if (v === "chunk") {
      box.appendChild(lab("每")); box.appendChild(inp("p-n", "number", 10, "70px"));
      box.appendChild(lab("页一个文件"));
    } else if (v === "range") {
      box.appendChild(lab("页码")); box.appendChild(inp("p-r", "text", "", "200px", "如 1-3,5,8-10"));
    } else if (v === "pages") {
      box.appendChild(lab("删除页")); box.appendChild(inp("p-del", "text", "", "150px", "如 2,5,8-10"));
      box.appendChild(lab("旋转"));
      var sel = document.createElement("select"); sel.id = "p-rot";
      [["0", "不旋转"], ["90", "90°"], ["180", "180°"], ["270", "270°"]].forEach(function (o) {
        var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1]; sel.appendChild(op);
      });
      box.appendChild(sel);
      var chk = document.createElement("label"); chk.className = "check";
      chk.innerHTML = '<input type="checkbox" id="p-rev">倒序';
      box.appendChild(chk);
    } else if (v === "wm") {
      box.appendChild(lab("文字")); box.appendChild(inp("p-wmt", "text", "", "160px", "英文/数字"));
      box.appendChild(lab("字号")); box.appendChild(inp("p-wms", "number", 40, "70px"));
      box.appendChild(lab("透明度")); box.appendChild(inp("p-wmo", "number", 15, "70px"));
      var tip = document.createElement("span");
      tip.style.cssText = "font-size:12px;color:#8a6420";
      tip.textContent = "⚠ 网页版只支持英文水印";
      box.appendChild(tip);
    } else if (v === "size") {
      box.appendChild(lab("目标"));
      var sel2 = document.createElement("select"); sel2.id = "p-target";
      [["a4", "A4"], ["letter", "Letter"], ["a3", "A3"]].forEach(function (o) {
        var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1]; sel2.appendChild(op);
      });
      box.appendChild(sel2);
    } else if (v === "2img") {
      box.appendChild(lab("清晰度"));
      var dpiSel = document.createElement("select"); dpiSel.id = "p-dpi";
      [["72", "72 屏幕"], ["150", "150 普通打印"], ["300", "300 高清"]].forEach(function (o) {
        var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1];
        if (o[0] === "150") op.selected = true;
        dpiSel.appendChild(op);
      });
      box.appendChild(dpiSel);
      box.appendChild(lab("格式"));
      var fmtSel = document.createElement("select"); fmtSel.id = "p-imgfmt";
      [["png", "PNG"], ["jpg", "JPG"]].forEach(function (o) {
        var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1]; fmtSel.appendChild(op);
      });
      box.appendChild(fmtSel);
      box.appendChild(lab("页码"));
      box.appendChild(inp("p-imgpages", "text", "", "140px", "留空=全部"));
    } else if (v === "2txt") {
      box.appendChild(lab("页码"));
      box.appendChild(inp("p-txtpages", "text", "", "160px", "留空=全部"));
      var tip2 = document.createElement("span");
      tip2.style.cssText = "font-size:12px;color:var(--ink3)";
      tip2.textContent = "（扫描件没有文字层，提取会为空）";
      box.appendChild(tip2);
    } else if (v === "longimg") {
      box.appendChild(lab("清晰度"));
      var dpi2 = document.createElement("select"); dpi2.id = "p-longdpi";
      [["80", "80 省流量"], ["120", "120 标准"], ["200", "200 高清"]].forEach(function (o) {
        var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1];
        if (o[0] === "120") op.selected = true;
        dpi2.appendChild(op);
      });
      box.appendChild(dpi2);
      box.appendChild(lab("页码"));
      box.appendChild(inp("p-longpages", "text", "", "140px", "留空=全部"));
    } else if (v === "merge") {
      var s = document.createElement("span");
      s.style.cssText = "font-size:12.5px;color:var(--ink3)";
      s.textContent = "按列表顺序合并，可点 ↑ 调整";
      box.appendChild(s);
    } else {
      var s2 = document.createElement("span");
      s2.style.cssText = "font-size:12.5px;color:var(--ink3)";
      s2.textContent = v === "info" ? "读取页数、尺寸、元数据" : "每页拆成一个文件";
      box.appendChild(s2);
    }
  }
  renderPdfArgs("merge");

  $("p-run").onclick = async function () {
    var box = $("p-log"), res = $("p-res");
    clear(box); clear(res);
    var v = pMode.value;
    if (!pPick.files.length) { log(box, "请先选择 PDF", "err"); return; }
    if (v === "merge" && pPick.files.length < 2) { log(box, "合并至少需要 2 个 PDF", "err"); return; }
    this.disabled = true; prog($("p-prog"), 0);
    try {
      if (v === "merge") {
        var items = [];
        for (var i = 0; i < pPick.files.length; i++) {
          items.push({ name: pPick.files[i].name, data: await readBuf(pPick.files[i]) });
          prog($("p-prog"), (i + 1) / (pPick.files.length + 1));
        }
        var r = await TB.pdfMerge(items);
        r.skipped.forEach(function (s) { log(box, "✗ " + s.name + "：" + s.reason, "err"); });
        r.info.forEach(function (o) { log(box, "  " + o.name + "  " + o.count + " 页"); });
        log(box, "", null);
        log(box, "合并完成：共 " + r.pages + " 页", "ok");
        resultRow(res, "合并结果.pdf", r.data.byteLength, r.data, "application/pdf");
      } else if (v === "each" || v === "chunk" || v === "range") {
        var f = pPick.first;
        var rr = await TB.pdfSplit({ name: f.name, data: await readBuf(f) }, v, {
          pagesPerFile: +(($("p-n") || {}).value || 10),
          ranges: ($("p-r") || {}).value || "",
        });
        rr.results.forEach(function (x) { log(box, "  " + x.name + "  " + x.pages + " 页"); });
        log(box, "", null);
        log(box, "拆分完成：" + rr.results.length + " 个文件", "ok");
        if (rr.results.length === 1) resultRow(res, rr.results[0].name, rr.results[0].data.byteLength, rr.results[0].data, "application/pdf");
        else await bigZip(res, rr.results, "PDF拆分.zip");
      } else if (v === "pages") {
        var f2 = pPick.first;
        var r2 = await TB.pdfPageOps({ name: f2.name, data: await readBuf(f2) }, {
          removePages: ($("p-del") || {}).value || "",
          rotate: +(($("p-rot") || {}).value || 0),
          reverse: !!($("p-rev") && $("p-rev").checked),
        });
        if (r2.removed) log(box, "  删除 " + r2.removed + " 页");
        log(box, "处理完成：共 " + r2.pages + " 页", "ok");
        resultRow(res, f2.name.replace(/\.[^.]+$/, "") + "_处理后.pdf", r2.data.byteLength, r2.data, "application/pdf");
      } else if (v === "wm") {
        var text = ($("p-wmt") || {}).value || "";
        if (/[\u4e00-\u9fa5]/.test(text)) { log(box, "✗ 网页版不支持中文水印，请用英文", "err"); this.disabled = false; prog($("p-prog"), null); return; }
        var outs = [];
        for (var k = 0; k < pPick.files.length; k++) {
          var fk = pPick.files[k];
          var rk = await TB.pdfWatermark({ name: fk.name, data: await readBuf(fk) }, {
            text: text, pos: "center", size: +(($("p-wms") || {}).value || 40),
            opacity: +(($("p-wmo") || {}).value || 15), rotation: 45,
          });
          outs.push({ name: fk.name.replace(/\.[^.]+$/, "") + "_水印.pdf", data: rk.data });
          log(box, "✓ " + fk.name + "  " + rk.pages + " 页", "ok");
          prog($("p-prog"), (k + 1) / pPick.files.length);
        }
        if (outs.length === 1) resultRow(res, outs[0].name, outs[0].data.byteLength, outs[0].data, "application/pdf");
        else await bigZip(res, outs, "PDF水印.zip");
      } else if (v === "size") {
        var outs2 = [];
        for (var m = 0; m < pPick.files.length; m++) {
          var fm = pPick.files[m];
          var rm = await TB.pdfResize({ name: fm.name, data: await readBuf(fm) }, {
            target: ($("p-target") || {}).value || "a4",
          });
          outs2.push({ name: fm.name.replace(/\.[^.]+$/, "") + "_" + ($("p-target") || {}).value + ".pdf", data: rm.data });
          log(box, "✓ " + fm.name + " → " + Math.round(rm.size[0]) + "×" + Math.round(rm.size[1]) + "pt", "ok");
        }
        if (outs2.length === 1) resultRow(res, outs2[0].name, outs2[0].data.byteLength, outs2[0].data, "application/pdf");
        else await bigZip(res, outs2, "PDF尺寸统一.zip");
      } else if (v === "2img") {
        var outs3 = [];
        for (var q = 0; q < pPick.files.length; q++) {
          var fq = pPick.files[q];
          var rq = await TB.pdfToImages({ name: fq.name, data: await readBuf(fq) }, {
            dpi: +(($("p-dpi") || {}).value || 150),
            fmt: ($("p-imgfmt") || {}).value || "png",
            pages: ($("p-imgpages") || {}).value || "",
          }, window.pdfjsLib);
          rq.results.forEach(function (x) {
            outs3.push({ name: x.name, data: x.data });
          });
          log(box, "✓ " + fq.name + " → " + rq.results.length + " 张图", "ok");
          prog($("p-prog"), (q + 1) / pPick.files.length);
        }
        if (outs3.length === 1) resultRow(res, outs3[0].name, outs3[0].data.byteLength, outs3[0].data, "image/png");
        else if (outs3.length > 1) await bigZip(res, outs3, "PDF转图片.zip");
      } else if (v === "2txt") {
        var outsT = [];
        for (var t = 0; t < pPick.files.length; t++) {
          var ft = pPick.files[t];
          var rt2 = await TB.pdfToText({ name: ft.name, data: await readBuf(ft) }, {
            pages: ($("p-txtpages") || {}).value || "",
          }, window.pdfjsLib);
          outsT.push({ name: rt2.name, data: new TextEncoder().encode(rt2.text) });
          log(box, "✓ " + ft.name + "  " + rt2.pages + " 页，" + rt2.chars + " 字符", "ok");
          if (rt2.chars < 10) log(box, "  ⚠ 文字很少，可能是扫描件（没有文字层）", null);
          prog($("p-prog"), (t + 1) / pPick.files.length);
        }
        if (outsT.length === 1) resultRow(res, outsT[0].name, outsT[0].data.byteLength, outsT[0].data, "text/plain");
        else if (outsT.length > 1) await bigZip(res, outsT, "PDF文字.zip");
      } else if (v === "longimg") {
        var fL = pPick.first;
        var rL = await TB.pdfToLongImage({ name: fL.name, data: await readBuf(fL) }, {
          dpi: +(($("p-longdpi") || {}).value || 120),
          pages: ($("p-longpages") || {}).value || "",
        }, window.pdfjsLib);
        log(box, "✓ " + rL.pages + " 页拼成 " + rL.w + "×" + rL.h + " 长图", "ok");
        resultRow(res, rL.name, rL.data.byteLength, rL.data, "image/png");
      } else if (v === "info") {
        var fi = pPick.first;
        var txt = await TB.pdfInfo({ name: fi.name, data: await readBuf(fi) });
        txt.split("\n").forEach(function (line) { log(box, line); });
      }
    } catch (e) { log(box, "✗ " + (e.message || e), "err"); }
    this.disabled = false; prog($("p-prog"), null);
  };

  // ==================================================================
  // Word
  // ==================================================================
  var wPick = picker("w-drop", "w-files", "w-list", /\.docx$/i, true);
  var wMode = seg("w-mode", function (v) {
    var hint = $("w-hint");
    if (v === "html") hint.textContent = "Word 转 HTML 保留标题、段落、列表、表格、图片，可用于发公众号/网页";
    else if (v === "text") hint.textContent = "只提取纯文字，丢掉所有格式，适合复制到别处";
    else hint.textContent = "直接在网页里预览 Word 内容（不下载）";
  });
  $("w-run").onclick = async function () {
    var box = $("w-log"), res = $("w-res"), pv = $("w-preview");
    clear(box); clear(res); pv.style.display = "none"; pv.innerHTML = "";
    if (!wPick.files.length) { log(box, "请先选择 .docx 文件", "err"); return; }
    var v = wMode.value;
    this.disabled = true; prog($("w-prog"), 0);
    var outs = [];
    for (var i = 0; i < wPick.files.length; i++) {
      var f = wPick.files[i];
      try {
        if (v === "html") {
          var r = await TB.wordToHtmlFile({ name: f.name, data: await readBuf(f) });
          outs.push({ name: r.name, data: r.data });
          log(box, "✓ " + f.name + " → " + r.name, "ok");
        } else if (v === "text") {
          var rt = await TB.wordToText({ name: f.name, data: await readBuf(f) });
          outs.push({ name: rt.name, data: new TextEncoder().encode(rt.text) });
          log(box, "✓ " + f.name + " → " + rt.name + "（" + rt.text.length + " 字符）", "ok");
        } else {
          var rp = await TB.wordToHtml({ name: f.name, data: await readBuf(f) });
          pv.style.display = "block";
          pv.innerHTML += "<h4 style='margin:8px 0;font-size:13px;color:#55617a'>" + f.name + "</h4>" + rp.html + "<hr style='margin:16px 0;border:none;border-top:1px solid #e3e8f0'>";
          log(box, "✓ " + f.name + " 已预览", "ok");
        }
      } catch (e) { log(box, "✗ " + f.name + "：" + (e.message || e), "err"); }
      prog($("w-prog"), (i + 1) / wPick.files.length);
    }
    if (outs.length === 1) resultRow(res, outs[0].name, outs[0].data.byteLength, outs[0].data);
    else if (outs.length > 1) await bigZip(res, outs, "Word转换结果.zip");
    this.disabled = false; prog($("w-prog"), null);
  };

  // ==================================================================
  // Excel
  // ==================================================================
  var xPick = picker("x-drop", "x-files", "x-list", /\.(xlsx|xls|csv)$/i, true);
  var xHeaders = [];
  xPick.onAfter(async function () {
    var f = xPick.first, info = $("x-info");
    if (!f) { xHeaders = []; info.textContent = ""; return; }
    try {
      xHeaders = await TB.xlHeaders({ name: f.name, data: await readBuf(f) });
      info.textContent = "列名：" + xHeaders.slice(0, 12).join("、") + (xHeaders.length > 12 ? "…" : "");
      if (xMode.value === "split" || xMode.value === "cmp") renderXlArgs(xMode.value);
    } catch (e) { info.textContent = "读取失败：" + e.message; }
  });
  var xMode = seg("x-mode", renderXlArgs);
  function renderXlArgs(v) {
    var box = $("x-args"); clear(box);
    if (v === "split") {
      var l = document.createElement("label"); l.textContent = "按列";
      var sel = document.createElement("select"); sel.id = "x-key";
      if (!xHeaders.length) { var o = document.createElement("option"); o.textContent = "（先选文件）"; sel.appendChild(o); }
      xHeaders.forEach(function (h) { var op = document.createElement("option"); op.value = h; op.textContent = h; sel.appendChild(op); });
      var l2 = document.createElement("label"); l2.textContent = "拆分";
      box.appendChild(l); box.appendChild(sel); box.appendChild(l2);
    } else if (v === "cmp") {
      var l3 = document.createElement("label"); l3.textContent = "关键列";
      var sel2 = document.createElement("select"); sel2.id = "x-cmpkey";
      if (!xHeaders.length) { var o2 = document.createElement("option"); o2.textContent = "（先选文件）"; sel2.appendChild(o2); }
      xHeaders.forEach(function (h) { var op = document.createElement("option"); op.value = h; op.textContent = h; sel2.appendChild(op); });
      var tip = document.createElement("span");
      tip.style.cssText = "font-size:12px;color:var(--ink3)"; tip.textContent = "（需要恰好 2 个文件）";
      box.appendChild(l3); box.appendChild(sel2); box.appendChild(tip);
    } else {
      var s = document.createElement("span");
      s.style.cssText = "font-size:12.5px;color:var(--ink3)";
      var map = {
        merge: "按列名自动对齐（列顺序不同也行）",
        csv: "转成 CSV（带 BOM，Excel 打开不乱码）",
        text: "转成制表符分隔的纯文本",
        html: "转成 HTML 表格（可贴到网页）",
      };
      s.textContent = map[v] || "";
      box.appendChild(s);
    }
  }
  renderXlArgs("merge");

  $("x-run").onclick = async function () {
    var box = $("x-log"), res = $("x-res");
    clear(box); clear(res);
    if (!xPick.files.length) { log(box, "请先选择表格", "err"); return; }
    var v = xMode.value;
    if (v === "merge" && xPick.files.length < 2) { log(box, "合并至少需要 2 个表格", "err"); return; }
    if (v === "cmp" && xPick.files.length !== 2) { log(box, "比对需要恰好 2 个表格", "err"); return; }
    this.disabled = true; prog($("x-prog"), 0);
    try {
      var items = [];
      for (var i = 0; i < xPick.files.length; i++) {
        items.push({ name: xPick.files[i].name, data: await readBuf(xPick.files[i]) });
        prog($("x-prog"), (i + 1) / (xPick.files.length + 1));
      }
      if (v === "merge") {
        var r = await TB.xlMerge(items);
        r.perFile.forEach(function (p) { log(box, "  " + p.name + "  " + p.rows + " 行" + (p.note ? "（" + p.note + "）" : "")); });
        log(box, "", null);
        log(box, "合并完成：共 " + r.rows + " 行", "ok");
        resultRow(res, "合并结果.xlsx", r.data.byteLength, r.data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      } else if (v === "split") {
        var key = ($("x-key") || {}).value;
        if (!key) { log(box, "请选择拆分列", "err"); this.disabled = false; prog($("x-prog"), null); return; }
        var rr = await TB.xlSplit(items[0], key);
        log(box, "  共 " + rr.groups + " 组", null);
        log(box, "拆分完成", "ok");
        await bigZip(res, rr.results, "表格拆分.zip");
      } else if (v === "cmp") {
        var ck = ($("x-cmpkey") || {}).value;
        if (!ck) { log(box, "请选择关键列", "err"); this.disabled = false; prog($("x-prog"), null); return; }
        var rc = await TB.xlCompare(items[0], items[1], [ck]);
        log(box, "  仅在A：" + rc.onlyA + " 行｜仅在B：" + rc.onlyB + " 行｜值不同：" + rc.diffs + " 处");
        log(box, "比对完成", "ok");
        resultRow(res, "比对结果.xlsx", rc.data.byteLength, rc.data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      } else if (v === "csv") {
        var outs = [];
        for (var k = 0; k < items.length; k++) {
          var rk = await TB.xlToCsv(items[k]);
          outs.push({ name: rk.name, data: rk.data });
          log(box, "✓ " + items[k].name + " → " + rk.name, "ok");
        }
        if (outs.length === 1) resultRow(res, outs[0].name, outs[0].data.byteLength, outs[0].data, "text/csv");
        else await bigZip(res, outs, "CSV结果.zip");
      } else if (v === "text") {
        var outs2 = [];
        for (var m = 0; m < items.length; m++) {
          var txt = await TB.xlToText(items[m]);
          var nm = items[m].name.replace(/\.[^.]+$/, "") + ".txt";
          outs2.push({ name: nm, data: new TextEncoder().encode(txt) });
          log(box, "✓ " + items[m].name + " → " + nm, "ok");
        }
        if (outs2.length === 1) resultRow(res, outs2[0].name, outs2[0].data.byteLength, outs2[0].data, "text/plain");
        else await bigZip(res, outs2, "文本结果.zip");
      } else if (v === "html") {
        var outs3 = [];
        for (var n = 0; n < items.length; n++) {
          var html = await TB.xlToHtml(items[n]);
          var page = "<!DOCTYPE html><html><head><meta charset='utf-8'><style>table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:6px}</style></head><body>" + html + "</body></html>";
          var nm2 = items[n].name.replace(/\.[^.]+$/, "") + ".html";
          outs3.push({ name: nm2, data: new TextEncoder().encode(page) });
          log(box, "✓ " + items[n].name + " → " + nm2, "ok");
        }
        if (outs3.length === 1) resultRow(res, outs3[0].name, outs3[0].data.byteLength, outs3[0].data, "text/html");
        else await bigZip(res, outs3, "HTML表格.zip");
      }
    } catch (e) { log(box, "✗ " + (e.message || e), "err"); }
    this.disabled = false; prog($("x-prog"), null);
  };

  // ==================================================================
  // 图片 → PDF
  // ==================================================================
  var tiPick = picker("ti-drop", "ti-files", "ti-list", /^image\//, true);
  $("ti-run").onclick = async function () {
    var box = $("ti-log"), res = $("ti-res");
    clear(box); clear(res);
    if (!tiPick.files.length) { log(box, "请先选择图片", "err"); return; }
    this.disabled = true; prog($("ti-prog"), 0);
    try {
      var items = [];
      for (var i = 0; i < tiPick.files.length; i++) {
        var f = tiPick.files[i];
        items.push({ name: f.name, data: await readBuf(f), type: f.type });
        prog($("ti-prog"), (i + 1) / (tiPick.files.length + 1));
      }
      var r = await TB.imagesToPdf(items, { target: $("ti-target").value });
      r.perFile.forEach(function (p) {
        if (p.ok) log(box, "  ✓ " + p.name);
        else log(box, "  ✗ " + p.name + "：" + p.reason, "err");
      });
      log(box, "", null);
      log(box, "合成完成：共 " + r.pages + " 页", "ok");
      resultRow(res, "图片合集.pdf", r.data.byteLength, r.data, "application/pdf");
    } catch (e) { log(box, "✗ " + (e.message || e), "err"); }
    this.disabled = false; prog($("ti-prog"), null);
  };

  // ==================================================================
  // 格式转换
  // ==================================================================
  var cPick = picker("c-drop", "c-files", "c-list", /\.(txt|md|html?|csv|json)$/i, true);
  $("c-run").onclick = async function () {
    var box = $("c-log"), res = $("c-res");
    clear(box); clear(res);
    if (!cPick.files.length) { log(box, "请先选择文件", "err"); return; }
    var target = $("c-target").value;
    this.disabled = true; prog($("c-prog"), 0);
    var outs = [];
    for (var i = 0; i < cPick.files.length; i++) {
      var f = cPick.files[i];
      try {
        var text = await readText(f);
        var stem = f.name.replace(/\.[^.]+$/, "");
        var out = null, outName = "";
        if (target === "docx") {
          var tmp = { name: f.name, data: new TextEncoder().encode(text) };
          var r = await TB.textToWord(tmp, {});
          out = r.data; outName = stem + ".docx";
        } else if (target === "html") {
          var isHtml = /\.html?$/i.test(f.name);
          var body = isHtml ? text : "<pre style='white-space:pre-wrap;font-family:inherit'>" +
            text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;") + "</pre>";
          out = new TextEncoder().encode(
            "<!DOCTYPE html><html><head><meta charset='utf-8'><title>" + stem +
            "</title><style>body{font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;" +
            "max-width:800px;margin:40px auto;padding:0 20px;line-height:1.8;color:#1a2233}</style></head><body>" +
            body + "</body></html>");
          outName = stem + ".html";
        } else if (target === "txt") {
          out = new TextEncoder().encode(text.replace(/<[^>]+>/g, ""));
          outName = stem + ".txt";
        } else if (target === "csv") {
          // 按行/逗号/制表符切分
          var rows = text.split(/\r?\n/).filter(function (l) { return l.trim(); })
            .map(function (l) { return l.split(/\t|,|，/).map(function (c) { return '"' + c.replace(/"/g, '""') + '"'; }).join(","); });
          out = new TextEncoder().encode("\ufeff" + rows.join("\n"));
          outName = stem + ".csv";
        } else if (target === "json") {
          var data;
          try { data = JSON.parse(text); } catch (e) {
            data = { lines: text.split(/\r?\n/) };
          }
          out = new TextEncoder().encode(JSON.stringify(data, null, 2));
          outName = stem + ".json";
        }
        if (out) {
          outs.push({ name: outName, data: out });
          log(box, "✓ " + f.name + " → " + outName, "ok");
        }
      } catch (e) { log(box, "✗ " + f.name + "：" + (e.message || e), "err"); }
      prog($("c-prog"), (i + 1) / cPick.files.length);
    }
    if (outs.length === 1) resultRow(res, outs[0].name, outs[0].data.byteLength, outs[0].data);
    else if (outs.length > 1) await bigZip(res, outs, "格式转换结果.zip");
    this.disabled = false; prog($("c-prog"), null);
  };
})();
