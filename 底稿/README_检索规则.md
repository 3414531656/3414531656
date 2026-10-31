# 政府公开信息检索 · 统一规则（6 个任务全部适用）

## 0. 你的身份与边界
- 你是「任务 N」的独立检索单元。只做分配给你的那一个任务。
- 与其他 5 个任务完全隔离：禁止引用其他任务的数字、链接、结论；禁止用其他任务的数据补齐本任务。
- 只写本任务文件：`C:\Users\341453\Downloads\ln_marine_data\taskN.md`；下载物放 `C:\Users\341453\Downloads\ln_marine_data\taskN\`。不得改其他文件。
- 找不到就写「检索不到」并列出检索式。**宁可空手，不要编造**。编造 URL、数字、页码、文件编号，整份结果作废。

## 1. 每条结论必须凑齐五要素（缺一项就不写进结论表）
发布单位 | 文件/公告全称 | 发布日期 | 页码或表号 | 可打开的原始 URL
- 网页无页码时，写定位到章节/段落标题（如「网页无页码；位于『二、农业』第3段」）。
- PDF 必须自己下载并抽文本确认真实页码（见第4节），不得凭想象写页码。
- 每条数字后附 1–2 句**照抄原文**（不改写、不转述）。

## 2. 只认原始出处
- 可作 A 级来源的域名：`*.gov.cn`（含 ln.gov.cn、tjj.ln.gov.cn、sthj.ln.gov.cn、zrzy.ln.gov.cn、moa.gov.cn、mnr.gov.cn、mee.gov.cn、stats.gov.cn、ccgp.gov.cn、ccgp-liaoning.gov.cn、ggzy.ln.gov.cn 等）、`*.gov.cn` 下的 PDF、政府公报、统计年鉴官方版。
- 二手站：chinabidding、招标网、搜狐、百家号、公众号、道客巴巴、百度文库、各类"年鉴数据站"、新闻网站（东北新闻网等）。**只能当线索**，必须标「二手，原始出处待查」，不得作为结论来源。
- 证据分级：A＝政府原始文件可直接打开（你亲自打开过，含 PDF 已下载抽取）；B＝原始文件存在但只能看到检索摘要或转载全文；C＝仅第三方转述。

## 3. 检索工具与用法
- `web_search`：多组关键词并行搜，支持 `site:`；每条检索式记入「未能获取」表。
- `web_fetch`：只支持 HTML；抓 PDF 会报 `unsupported content type "application/pdf"`，不要反复试。
- 页面打不开时：试 `https://web.archive.org/web/2026/<原URL>`、搜索引擎快照、上级栏目页。

## 4. PDF 下载与逐字核对 SOP（重要）
Python：`C:\Users\341453\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`
工具脚本：`C:\Users\341453\Downloads\ln_marine_data\pdftool.py`（已装 pymupdf）

```powershell
# 1) 下载（带 UA，避免 403）
Invoke-WebRequest -Uri "<PDF直链>" -OutFile "C:\Users\341453\Downloads\ln_marine_data\taskN\x.pdf" -Headers @{'User-Agent'='Mozilla/5.0'} -TimeoutSec 60
# 2) 看页数、是否有文本层
& "<python路径>" "C:\Users\341453\Downloads\ln_marine_data\pdftool.py" info "...\x.pdf"
# 3) 转 UTF-8 文本（务必用 save，不要把中文直接打到控制台，会乱码）
& "<python路径>" "C:\Users\341453\Downloads\ln_marine_data\pdftool.py" save "...\x.pdf" "...\x.txt"
```
- 然后用 `grep` / `read` 工具读 `.txt`，搜关键词，得到 `===== PAGE 8/58 =====` 形式的真实页码。
- `pages_without_text_layer` > 0 说明是扫描版：如实写「扫描版，无文本层，无法逐字核对」，不要猜内容。
- 页脚印刷页码与 PDF 页序不一致时，两个都写（如「PDF 第3页，页脚·3」）。

## 5. 输出文件模板（严格照此写 taskN.md）
```
# 任务 N｜<标题>

## N-1 结论表
| 指标 | 年份 | 数值 | 单位 | 口径 | 来源文件全称 | 发布单位 | 发布日期 | 页码/表号 | URL | 证据等级 |
|---|---|---|---|---|---|---|---|---|---|---|

## N-2 原文摘录
1. <指标>（<年份>）
   > <照抄原文 1–2 句>
   出处：<文件全称>，<PDF第X页/表X/网页章节>，<URL>

## N-3 URL 核验记录
| URL | 打开方式 | 结果 | 命中关键词/页码 |
|---|---|---|---|

## N-4 未能获取
| 指标 | 试过的检索式（≥3 条） | 建议去哪个网站哪个栏目 |
|---|---|---|

## N-5 口径冲突与提醒
- <同一指标不同来源的数字/单位/年份差异；注明哪个可用、为什么>
```

## 6. 禁区
- 不检索辽宁入海排污口（已完成，145／420／2 088 为累计口径）。
- 不写引言、总结、建议、对策、"可写入说明书"之类段落。
- 不用全国数冒充辽宁数；全国数必须标「仅全国口径」。
- 不把不同年份的面积与占比强行配对。
- 不引用其他任务的数字。

## 7. 完成后汇报（send_message 给 lead，≤15 行）
状态（找到/部分/未找到）｜本任务最关键的 1–3 个数值｜A 级原始 URL｜文件路径｜发现的冲突或异常（如「92.8% 出自某单页PDF」）。
