/**
 * 浏览器端冒烟测试（内联版，无网络竞态）
 *
 * 为什么内联：jsdom 通过 HTTP 加载多个大脚本（jsQR 256KB、docx 807KB、mammoth 642KB）
 * 时存在资源加载竞态，同一条命令反复跑会得到不同结果（曾一次 1 FAIL、一次 9 FAIL）。
 * 内联后脚本按文档顺序同步执行，结果确定。这不改变被测代码。
 *
 * 依赖解析：本目录没有独立 node_modules（pnpm 布局无法跨目录复制），
 * 统一从 ../webapp/node_modules 解析。用 module.createRequire 显式指定。
 */
const path = require("path");
const Module = require("module");
const NM = path.join(__dirname, "..", "webapp", "node_modules");
const req = Module.createRequire(path.join(NM, "jsdom", "package.json"));
const { JSDOM } = req("jsdom");
const XLSX = req("xlsx");
const fs = require("fs");

const SRC = __dirname;
const GLOBALS = ["PDFLib", "jsQR", "imageCompression", "XLSX", "JSZip", "mammoth", "docx", "createOfficeToolbox"];

let pass = 0, fail = 0;
const ck = (n, c, d) => {
  if (c) { pass++; console.log("  PASS ", n); }
  else { fail++; console.log("  FAIL ", n, d === undefined ? "" : d); }
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  let html = fs.readFileSync(path.join(SRC, "index.html"), "utf8");
  html = html.replace(/<script src="([^"]+)"><\/script>/g, (m, src) => {
    const p = path.join(SRC, src);
    if (!fs.existsSync(p)) { console.error("缺脚本文件: " + src); return m; }
    return "<script>\n" + fs.readFileSync(p, "utf8") + "\n</script>";
  });

  const dom = new JSDOM(html, {
    runScripts: "dangerously", pretendToBeVisual: true, url: "http://localhost/",
  });
  const w = dom.window;
  const d = w.document;
  const $ = (id) => d.getElementById(id);

  // 等初始化完成（内联脚本同步执行，这里只是兜底）
  const rdy = Date.now() + 20000;
  while (Date.now() < rdy) {
    if (d.documentElement.getAttribute("data-office-ready") === "1") break;
    await sleep(100);
  }
  ck("app 已初始化", d.documentElement.getAttribute("data-office-ready") === "1");

  console.log("[库全局名]");
  ck("window.jsQR 是函数", typeof w.jsQR === "function", typeof w.jsQR);
  ck("window.imageCompression 是函数", typeof w.imageCompression === "function", typeof w.imageCompression);
  ck("window.docx 存在", !!w.docx);
  ck("window.PDFLib 存在", !!w.PDFLib);
  ck("window.XLSX 存在", !!w.XLSX);
  ck("createOfficeToolbox 存在", typeof w.createOfficeToolbox === "function");

  console.log("[界面结构]");
  ck("8 个导航项", d.querySelectorAll("nav button").length === 8, d.querySelectorAll("nav button").length);
  ck("8 个功能区", d.querySelectorAll("section").length === 8, d.querySelectorAll("section").length);
  ck("图片批处理+区", !!$("s-ib"));
  ck("数据处理区", !!$("s-xp"));

  console.log("[图片批处理+ 模式切换]");
  const bModes = [...d.querySelectorAll("#b-mode button")].map((b) => b.dataset.v);
  ck("4 种模式齐全", ["qr", "wm2", "stitch", "gif"].every((m) => bModes.includes(m)), bModes.join(","));
  [...d.querySelectorAll("#b-mode button")].find((b) => b.dataset.v === "stitch").click();
  await sleep(120);
  ck("拼接参数出现", !!$("b-dir") && !!$("b-gap") && !!$("b-cols"));
  [...d.querySelectorAll("#b-mode button")].find((b) => b.dataset.v === "wm2").click();
  await sleep(120);
  ck("水印参数出现", !!$("b-wmt") && !!$("b-wmp") && !!$("b-wmo"));

  console.log("[数据处理 模式切换]");
  const xpModes = [...d.querySelectorAll("#xp-mode button")].map((b) => b.dataset.v);
  ck("6 种模式齐全", ["clean", "sort", "filter", "stats", "splitcol", "mergecol"].every((m) => xpModes.includes(m)), xpModes.join(","));
  [...d.querySelectorAll("#xp-mode button")].find((b) => b.dataset.v === "filter").click();
  await sleep(120);
  ck("筛选参数出现", !!$("xp-key") && !!$("xp-cond") && !!$("xp-val"));
  [...d.querySelectorAll("#xp-mode button")].find((b) => b.dataset.v === "mergecol").click();
  await sleep(120);
  ck("合并列多选框", !!$("xp-keys") && $("xp-keys").multiple === true);

  console.log("[导航切换]");
  [...d.querySelectorAll("nav button")].find((b) => b.dataset.t === "ib").click();
  await sleep(80);
  ck("切到图片批处理+后该区显示", $("s-ib").classList.contains("on"));
  [...d.querySelectorAll("nav button")].find((b) => b.dataset.t === "xp").click();
  await sleep(80);
  ck("切到数据处理后该区显示", $("s-xp").classList.contains("on"));

  console.log("[端到端：选真实表格 → 排序 → 出结果]");
  const ws = XLSX.utils.aoa_to_sheet([["名字", "分数"], ["甲", 60], ["乙", 90], ["丙", 75]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const file = new w.File([new Uint8Array(buf)], "成绩.xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const xpInput = $("xp-files");
  Object.defineProperty(xpInput, "files", { value: [file], configurable: true });
  xpInput.dispatchEvent(new w.Event("change", { bubbles: true }));
  await sleep(600);
  ck("列名已读出", /名字/.test($("xp-info").textContent), $("xp-info").textContent);

  [...d.querySelectorAll("#xp-mode button")].find((b) => b.dataset.v === "sort").click();
  await sleep(250);
  ck("排序参数出现", !!$("xp-key") && !!$("xp-dir"));
  $("xp-key").value = "分数";
  $("xp-dir").value = "desc";
  $("xp-run").click();
  await sleep(3000);
  ck("执行的是排序（不是清理）", /已按「分数」排序/.test($("xp-log").textContent), $("xp-log").textContent.slice(0, 100));
  ck("输出文件名带_排序", /_排序\.xlsx/.test($("xp-res").innerHTML), $("xp-res").innerHTML.slice(0, 120));

  [...d.querySelectorAll("#xp-mode button")].find((b) => b.dataset.v === "filter").click();
  await sleep(250);
  ck("再次切换后筛选参数出现", !!$("xp-key") && !!$("xp-cond") && !!$("xp-val"));

  console.log("\n" + "=".repeat(44));
  console.log("通过 " + pass + "，失败 " + fail);
  w.close();
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("FATAL", e); process.exit(1); });
