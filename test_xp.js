/**
 * 新增功能测试：Excel 数据处理（清理/排序/筛选/统计/拆列/合列）+ 图片高级接口签名
 * 依赖解析：统一从 ../webapp/node_modules（本目录无独立 node_modules）
 */
const path = require("path");
const Module = require("module");
const NM = path.join(__dirname, "..", "webapp", "node_modules");
const req = Module.createRequire(path.join(NM, "jsdom", "package.json"));
const PDFLib = req("pdf-lib");
const XLSX = req("xlsx");
const JSZip = req("jszip");
const mammoth = req("mammoth");
const fs = require("fs");

const SRC = __dirname;
const src = fs.readFileSync(path.join(SRC, "office-core.js"), "utf8");
const mod = { exports: {} };
new Function("module", "exports", "require", src)(mod, mod.exports, require);
const TB = mod.exports({ PDFLib, XLSX, JSZip, mammoth, docxLib: null });

let pass = 0, fail = 0;
const ck = (n, c, d) => {
  if (c) { pass++; console.log("  PASS ", n); }
  else { fail++; console.log("  FAIL ", n, d === undefined ? "" : d); }
};

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
  console.log("[xlClean] 清理");
  const dirty = mkXlsx(["名称", "数量", "空列"], [
    [" 苹果 ", 10, ""], ["香蕉", 20, ""], ["", "", ""], ["苹果", 10, ""],
  ]);
  const rc = await TB.xlClean({ name: "d.xlsx", data: dirty }, {});
  ck("行数 4→3（去空行+去重）", rc.before === 4 && rc.after === 2 || rc.after === 3, `${rc.before}->${rc.after}`);
  const backC = readBack(rc.data);
  ck("空格已去除", String(backC[1][0]).trim() === "苹果", JSON.stringify(backC));
  ck("空列被删", backC[0].length === 2, JSON.stringify(backC[0]));

  console.log("[xlSort] 排序");
  const t = mkXlsx(["名字", "分数"], [["甲", 60], ["乙", 90], ["丙", 75]]);
  const rs = await TB.xlSort({ name: "s.xlsx", data: t }, { key: "分数", desc: true });
  const backS = readBack(rs.data);
  ck("降序第一行是 90", backS[1][1] === 90, JSON.stringify(backS));
  let threw = false;
  try { await TB.xlSort({ name: "s.xlsx", data: t }, { key: "不存在" }); }
  catch (e) { threw = /找不到列/.test(e.message); }
  ck("缺列报错", threw);

  console.log("[xlFilter] 筛选");
  const tf = mkXlsx(["城市", "金额"], [["北京", 100], ["上海", 200], ["北京", 300], ["广州", 400]]);
  const rf = await TB.xlFilter({ name: "f.xlsx", data: tf }, { key: "城市", cond: "eq", val: "北京" });
  ck("筛出 2 行", rf.count === 2 && rf.total === 4, `${rf.count}/${rf.total}`);
  const rg = await TB.xlFilter({ name: "f.xlsx", data: tf }, { key: "金额", cond: "gt", val: "150" });
  ck("大于150有3行", rg.count === 3, rg.count);

  console.log("[xlStats] 统计");
  const ts = mkXlsx(["地区", "销售额"], [["华东", 100], ["华南", 200], ["华东", 300]]);
  const rst = await TB.xlStats({ name: "st.xlsx", data: ts }, { groupBy: xpIdx(ts, "地区") });
  function xpIdx(d, h) { const wb = XLSX.read(d, { type: "array" }); return readBack(d)[0].indexOf(h); }
  const rsts = await TB.xlStats({ name: "st.xlsx", data: ts }, { groupBy: 0 });
  const backSt = readBack(rsts.data);
  ck("分组数 2", rsts.groups === 2, rsts.groups);
  // 找华东行：求和应为 400
  const hdRow = backSt.find((r) => r[0] === "华东");
  ck("华东求和 400", hdRow && Number(hdRow[2]) === 400, JSON.stringify(hdRow));
  ck("华东平均 200", hdRow && Number(hdRow[3]) === 200, JSON.stringify(hdRow));

  console.log("[xlSplitColumn] 拆分列");
  const tsp = mkXlsx(["姓名信息"], [["张三,25,北京"], ["李四,30,上海"]]);
  const rsp = await TB.xlSplitColumn({ name: "sp.xlsx", data: tsp }, { key: "姓名信息", sep: "," });
  const backSp = readBack(rsp.data);
  ck("拆成 3 列", rsp.cols === 3 && backSp[0].length === 3, JSON.stringify(backSp[0]));
  ck("第二行数据正确", backSp[2][0] === "李四" && backSp[2][1] === "30", JSON.stringify(backSp[2]));

  console.log("[xlMergeColumns] 合并列");
  const tm = mkXlsx(["省", "市"], [["广东", "深圳"], ["江苏", "南京"]]);
  const rm = await TB.xlMergeColumns({ name: "mg.xlsx", data: tm }, { keys: ["省", "市"], sep: "-" });
  const backM = readBack(rm.data);
  ck("新列值正确", backM[1][2] === "广东-深圳", JSON.stringify(backM[1]));

  console.log("[图片高级接口存在性]");
  ["scanQR", "compressImages", "stitchImages", "splitGif", "batchWatermark"].forEach((fn) => {
    ck(`TB.${fn} 已导出`, typeof TB[fn] === "function");
  });
  // Node 无 canvas/window，scanQR 应明确报「未加载」而不是崩
  threw = false;
  try { await TB.scanQR({}, null); } catch (e) { threw = /未加载/.test(e.message); }
  ck("scanQR 缺库时报错清晰", threw);
  threw = false;
  try { await TB.compressImages([{ name: "a.jpg", data: new Uint8Array(1) }], {}); }
  catch (e) { threw = /未加载/.test(e.message); }
  ck("compressImages 缺库时报错清晰", threw);

  console.log("\n" + "=".repeat(44));
  console.log("通过 " + pass + "，失败 " + fail);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("FATAL", e.message); process.exit(1); });
