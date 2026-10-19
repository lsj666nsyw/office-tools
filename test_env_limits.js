/**
 * 环境限制行为测试：验证超时保护真的生效
 *
 * 背景（实测结论，2026-10-19）：
 *   - jsdom 的 pdf.js worker 是 fake worker，getDocument().promise 永不 resolve
 *   - jsdom 的 Blob 实现不完整，docx.Packer.toBlob 静默挂起
 *   - 纯 Node 下同一个 docx.iife.js 打包正常（7652 字节），说明是环境差异
 *
 * 真实浏览器两者都正常。本测试验证的是：即使在这些受限环境里，
 * 也不能让用户面对一个永久转圈、没有任何提示的按钮。
 */
const path = require("path");
const Module = require("module");
const fs = require("fs");
const NM = path.join(__dirname, "..", "webapp", "node_modules");
const req = Module.createRequire(path.join(NM, "jsdom", "package.json"));
const { JSDOM } = req("jsdom");
const XLSX = req("xlsx");
const PDFLib = req("pdf-lib");
const SRC = __dirname;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let pass = 0, fail = 0;
const ck = (n, c, d) => {
  if (c) { pass++; console.log("  PASS ", n); }
  else { fail++; console.log("  FAIL ", n, d === undefined ? "" : String(d).slice(0, 160)); }
};

(async () => {
  let html = fs.readFileSync(path.join(SRC, "index.html"), "utf8");
  html = html.replace(/<script src="([^"]+)"><\/script>/g, (m, src) => {
    const p = path.join(SRC, src);
    return fs.existsSync(p) ? "<script>\n" + fs.readFileSync(p, "utf8") + "\n</script>" : m;
  });
  const dom = new JSDOM(html, { runScripts: "dangerously", pretendToBeVisual: true, url: "http://localhost/" });
  const w = dom.window, d = w.document;
  const $ = (id) => d.getElementById(id);
  const rdy = Date.now() + 20000;
  while (Date.now() < rdy) { if (d.documentElement.getAttribute("data-office-ready") === "1") break; await sleep(100); }

  // ---------- 1. docx 打包超时必须报错，不能静默 ----------
  console.log("[docx 打包：挂起时必须有明确错误]");
  const docxRes = await Promise.race([
    w.eval(`(async () => {
      const TB = createOfficeToolbox({ PDFLib: window.PDFLib, XLSX: window.XLSX, JSZip: window.JSZip, mammoth: window.mammoth, docxLib: window.docx });
      const enc = new TextEncoder();
      try {
        const r = await TB.textToWord({ name: "n.txt", data: enc.encode("第一行\\n第二行") }, {});
        return "OK " + r.data.byteLength;
      } catch (e) { return "ERR " + (e && e.message); }
    })()`),
    new Promise((r) => setTimeout(() => r("HUNG 70s"), 70000)),
  ]);
  // 结果可能是 OK（环境支持）或 ERR（走到超时）；绝不能是 HUNG
  ck("不静默挂死（要么成功要么报错）", /^(OK|ERR)/.test(String(docxRes)), docxRes);

  // ---------- 2. pdf.js 加载超时必须报错 ----------
  console.log("\n[pdf.js 加载：挂起时必须有明确错误]");
  const doc = await PDFLib.PDFDocument.create();
  const font = await doc.embedFont(PDFLib.StandardFonts.Helvetica);
  doc.addPage([300, 200]).drawText("Hello", { x: 20, y: 150, size: 14, font });
  const pdfBuf = new Uint8Array(await doc.save());

  const pdfRes = await Promise.race([
    w.eval(`(async () => {
      const TB = createOfficeToolbox({ PDFLib: window.PDFLib, XLSX: window.XLSX, JSZip: window.JSZip, mammoth: window.mammoth, docxLib: window.docx });
      try {
        const r = await TB.pdfToText({ name: "a.pdf", data: new Uint8Array(${JSON.stringify([...pdfBuf])}) }, {}, window.pdfjsLib);
        return "OK pages=" + r.pages;
      } catch (e) { return "ERR " + (e && e.message); }
    })()`),
    new Promise((r) => setTimeout(() => r("HUNG 60s"), 60000)),
  ]);
  ck("不静默挂死（要么成功要么报错）", /^(OK|ERR)/.test(String(pdfRes)), pdfRes);

  // ---------- 3. UI 层：按钮必须在结束后恢复可点 ----------
  console.log("\n[UI：处理结束后按钮恢复可用]");
  const file = new w.File([pdfBuf], "a.pdf", { type: "application/pdf" });
  const pi = $("p-files");
  Object.defineProperty(pi, "files", { value: [file], configurable: true });
  pi.dispatchEvent(new w.Event("change", { bubbles: true }));
  await sleep(700);
  [...d.querySelectorAll("#p-mode button")].find((b) => b.dataset.v === "2txt").click();
  await sleep(250);
  $("p-run").click();
  await sleep(3000);
  ck("点击后按钮暂时禁用（防重复提交）", $("p-run").disabled === true);
  // 等超时走完（45s），按钮应恢复
  const btnDeadline = Date.now() + 55000;
  while (Date.now() < btnDeadline) {
    if ($("p-run").disabled === false) break;
    await sleep(1000);
  }
  ck("超时后按钮恢复可点", $("p-run").disabled === false, "still disabled");
  ck("超时后有错误日志", /超时|失败|✗/.test($("p-log").textContent), JSON.stringify($("p-log").textContent.slice(0, 120)));

  console.log("\n" + "=".repeat(48));
  console.log("通过 " + pass + "，失败 " + fail);
  w.close();
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("FATAL", e); process.exit(1); });
