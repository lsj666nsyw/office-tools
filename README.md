# 全能办公工具箱（纯前端）

图片 / PDF / Word / Excel / 格式转换 / 图片批处理+ / 数据处理，全部在浏览器本地完成，
文件不上传服务器，手机电脑都能用，断网可用。

## 运行

静态站点，任意 HTTP 服务器即可（不能用 `file://`，浏览器会拦脚本）：

```bash
cd products/office
python -m http.server 8099 --bind 127.0.0.1
# 打开 http://127.0.0.1:8099/
```

## 文件

| 文件 | 作用 |
|---|---|
| `index.html` | 页面结构 + 全部样式（内联 CSS） |
| `office-core.js` | 核心逻辑，导出 `createOfficeToolbox()`，不碰 DOM 元素 |
| `office-app.js` | 前端交互，读 `office-core.js` 的接口 |
| `libs/` | 第三方库：pdf-lib、pdfjs、jsQR、bic（图片压缩）、xlsx、jszip、mammoth、docx |

### 8 个功能区

图片、PDF、Word、Excel、格式转换、图片→PDF、图片批处理+（二维码/批量水印/拼接/GIF 拆帧）、
数据处理（清理/排序/筛选/统计/拆列/合列）

## 测试

```bash
npm test           # 主测试：校准 + 单测 + 冒烟 + 全模块端到端
npm run test:env   # 环境限制行为（超时保护），约 2 分钟
```

| 文件 | 覆盖 |
|---|---|
| `_calib.js` | 判卷夹具校准：证明被测函数真的在工作（AGENTS.md 第 1 条） |
| `test_xp.js` | Excel 数据处理 20 项（清理/排序/筛选/统计/拆列/合列） |
| `test_br2.js` | 冒烟 24 项：库加载、UI 结构、模式切换、端到端排序 |
| `test_all.js` | 全模块端到端 26 项，每个模块走完整 UI 流程 |
| `test_env_limits.js` | 受限环境下不静默挂死（约 2 分钟） |

测试用 jsdom 内联脚本跑，不走网络——通过 HTTP 加载多个大脚本时有资源竞态，
同一条命令反复跑结果不稳定（曾一次 1 FAIL、一次 9 FAIL）。

### 依赖解析

本目录没有独立 `node_modules`（pnpm 布局无法跨目录复制），测试统一从
`../webapp/node_modules` 解析，见各测试顶部的 `Module.createRequire`。

## 已知环境限制（实测结论）

以下三条都在 jsdom 里复现过，**纯 Node 或真实浏览器下均正常**，属环境差异：

| 现象 | 原因 | 处理 |
|---|---|---|
| `URL.createObjectURL is not a function` | jsdom 未实现 | 图片预览在 jsdom 下不可用，真实浏览器正常 |
| docx `Packer.toBlob` 静默挂起 | jsdom 的 Blob 实现不完整 | 已改优先走 `toBuffer`，并加 60s 超时 |
| JSZip `generateAsync` 静默挂起 | DEFLATE 依赖 `setImmediate`，jsdom window 作用域没有 | 已加 45s 超时 |
| pdf.js `getDocument().promise` 永不 resolve | jsdom 的 worker 是 fake worker | 已加 45s 超时 |

共同点：**没有超时保护时，用户点完按钮会面对一个永久转圈、没有任何提示的按钮，
且按钮不恢复**。`test_env_limits.js` 专门守这个行为。

## 踩过的坑（2026-10-19 实测）

1. **`seg()` 缓存 `.on` 元素**：切换模式后 `value` 仍返回初始值，
   表现为「点了排序却执行清理」。影响全部 8 个模块。必须每次动态查询。
2. **picker 的 accept 只测文件名**：图片 picker 用 `/^image\//`，拖拽进来的
   文件只带 `name`（`图.png`），被静默丢弃。需同时测 MIME。
3. **jsdom 的 "Could not load script" 是 stderr 噪音**：文件其实 200 可达，
   据它判断"库没加载"会误判。要用就绪信号而非猜。
4. **测试用例之间会互相污染**：picker 是多选的，前一个用例选的文件还在，
   后一个用例的文件数就变了。需要在用例间清空。
