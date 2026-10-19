/**
 * 判卷夹具校准：故意写坏的样本必须在受测函数上 FAIL。
 * 目的：证明本测试不是在白送分（AGENTS.md 第 1 条：双向验证）。
 * 用法：node _calib.js
 */
const path = require("path");
const Module = require("module");
const fs = require("fs");
const NM = path.join(__dirname, "..", "webapp", "node_modules");
const req = Module.createRequire(path.join(NM, "jsdom", "package.json"));
const XLSX = req("xlsx");

const src = fs.readFileSync(path.join(__dirname, "office-core.js"), "utf8");
const mod = { exports: {} };
new Function("module", "exports", "require", src)(mod, mod.exports, require);
const TB = mod.exports({ PDFLib: req("pdf-lib"), XLSX, JSZip: req("jszip"), mammoth: req("mammoth"), docxLib: null });

function mkXlsx(header, rows) {
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}
function readBack(data) {
  const wb = XLSX.read(data, { type: "array" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
}

(async () => {
  let bad = 0, good = 0;

  // 1) xlSort 的 desc 必须真的生效：升序/降序应不同
  const t = mkXlsx(["名字", "分数"], [["甲", 60], ["乙", 90], ["丙", 75]]);
  const asc = readBack((await TB.xlSort({ name: "s.xlsx", data: t }, { key: "分数" })).data);
  const desc = readBack((await TB.xlSort({ name: "s.xlsx", data: t }, { key: "分数", desc: true })).data);
  const ascTop = Number(asc[1][1]), descTop = Number(desc[1][1]);
  if (ascTop === descTop) { bad++; console.log("  BAD xlSort 方向未生效（两者首行都是 " + ascTop + "）"); }
  else { good++; console.log("  OK  xlSort 方向生效 asc=" + ascTop + " desc=" + descTop); }

  // 2) xlFilter 必须真的筛掉行：条件不匹配时应 count < total
  const tf = mkXlsx(["城市", "金额"], [["北京", 100], ["上海", 200], ["北京", 300], ["广州", 400]]);
  const rf = await TB.xlFilter({ name: "f.xlsx", data: tf }, { key: "城市", cond: "eq", val: "北京" });
  if (rf.count === rf.total) { bad++; console.log("  BAD xlFilter 没筛掉任何行 " + rf.count + "/" + rf.total); }
  else { good++; console.log("  OK  xlFilter 生效 " + rf.count + "/" + rf.total); }

  // 3) xlClean 必须真的去掉重复行
  const dirty = mkXlsx(["名称", "数量"], [["苹果", 10], ["苹果", 10], ["", ""], ["香蕉", 20]]);
  const rc = await TB.xlClean({ name: "d.xlsx", data: dirty }, {});
  if (rc.after >= rc.before) { bad++; console.log("  BAD xlClean 没减少行 " + rc.before + "->" + rc.after); }
  else { good++; console.log("  OK  xlClean 生效 " + rc.before + "->" + rc.after); }

  // 4) xlStats 求和必须真的加总
  const ts = mkXlsx(["地区", "销售额"], [["华东", 100], ["华南", 200], ["华东", 300]]);
  const rst = await TB.xlStats({ name: "st.xlsx", data: ts }, { groupBy: 0 });
  const back = readBack(rst.data);
  const hd = back.find((r) => r[0] === "华东");
  if (!hd || Number(hd[2]) !== 400) { bad++; console.log("  BAD xlStats 求和错 " + JSON.stringify(hd)); }
  else { good++; console.log("  OK  xlStats 求和=400"); }

  console.log("\n校准：可用样本 " + good + "，异常 " + bad);
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error("FATAL", e.message); process.exit(1); });
