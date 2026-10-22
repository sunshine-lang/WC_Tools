# WC 翻译准备工作台

项目保留三个入口：翻译准备工作台、标签空格清洗、译后质量检查。准备工作台支持项目、资料库、批次、六节点流水线、逐节点确认、失败重试、产物下载，状态落盘在 `workspace/pipeline/`。

## 启动

```bash
.venv/bin/python app.py
```

打开 **http://127.0.0.1:5001/pipeline**（`/` 也会跳到这里）。标签清洗在 `/tag_clean`，译后检查在 `/qa`，三个入口可通过侧栏互相切换。原五个准备页的书签会跳到主工作台，旧上传和资料库 API 已移除。

已有项目继续读取 `workspace/pipeline/projects.json` 中的资料库。全新启动时会建立「听风城堡」项目，可从上级目录初始化游戏介绍和客户要求；术语库、TM 在资料库页上传，不再读取旧 uploads 知识库。分类规则位于 `profiles/game/wc_key_rules.json`。

## 使用流程

1. 打开 `/pipeline`，拖入或点击选择 XLSX / CSV 待译文件。
2. 资料库的 **术语库 + 记忆库** 为必需项；缺项时顶栏会变成「补齐资料库 →」。
3. 点「开始准备」。自动节点连续执行，需要判断的节点停在「待确认」。
4. 每个确认节点下方有「确认并继续」「重跑本步」「下载产物」。
5. 任一节点失败会保留前序产物，给「重试本步」「重试并继续」；重跑本步会作废下游。
6. 跑完后右上角「下载最终产物」拿到 `05-TM.xlsx`（源文件 + speaker/分类/Context/生效术语/无效术语/TM参考 等新增列），下载文件名为「上传文件名_ready.xlsx」（替换原扩展名）。
7. 「准备自检」节点已暂时停用（`pipeline_service.PREFLIGHT_ENABLED = False`），不再作为一个确认步骤；`core_tools/preflight.py` 代码保留，需要时把开关改回 `True` 并重启即可恢复。
8. 顶栏「重置」现在打开确认面板：优先 **新建批次重跑**（保留旧批次和产物，基于同一源文件新建一个 0/6 批次），也可以 **清空当前批次**（删除 00-05，只留源文件，不可撤销）；不再使用浏览器原生 confirm。

批次和项目都可以在左侧切换：批次下拉框切换历史批次，左上角项目名打开项目列表 / 新建项目。

## 六节点

| # | 节点 | 来源 | 是否确认 |
|---|---|---|---|
| 00 | 规范化 | `core_tools/normalizer.py` | 自动 |
| 01 | 分类 | `core_tools/key_classifier.py` | 自动 |
| 02 | 上下文 | `core_tools/context_processor.py` + `core_tools/dialogue_context.py`（活动对白 LLM） | 需确认 |
| 03 | 术语 | `core_tools/term_scanner.py` | 自动 |
| 04 | 术语筛选 | `core_tools/term_filter.py`（LLM） | 需确认 |
| 05 | TM 参考 | `core_tools/tm_matcher.py` + `core_tools/tm_semantic.py` | 需确认 |

`06 准备自检` 目前不在流程里；代码和九条断言保留在 `core_tools/preflight.py`，由 `pipeline_service.PREFLIGHT_ENABLED` 控制是否挂回流程。

资料库里的游戏背景、人物人设、人物关系和风格指南会全文作为术语筛选的参考语境，不再截取前 1000 字符；术语库和记忆库直接决定第 03/05 步的数据。

「人物人设」与「人物关系」是两个独立上传入口，分别支持 UTF-8 `.md` / `.txt` 文件，可单独替换和移除。每个入口对应一份文档，可在文档内部按角色组织。旧版「人设 / 关系」合并文件原样保留在「人物人设」入口并显示待拆分提示；用户上传新的人设文件后提示消失，「人物关系」需单独上传。升级不自动拆写文件，不改现有批次产物。

对白支持客户的 `note = 完整 Key、key = 说话人` 格式：上下文和术语筛选读取同一身份信息，按活动、段号分组、句号数字排序，原始 `note`、`key` 和已有译文均保留。缺完整 Key 时不猜分段。更新后旧批次需重跑上下文及下游；若旧源文件丢失 `note`，应使用客户原始文件新建批次。

## 活动对白场景 Context（2026-10-22）

第 02 步只为「活动对白」生成场景分析。按完整 Key 分组、按数字句序通读该组全部已提供英文，并全文传入当前项目的人物人设、人物关系、游戏背景与风格资料，不摘要或截断人物原文。模型只选用与本组参与者、明确提及对象及当前剧情相关的信息；同一内容同时上传到多个入口时合并传入一次，并提示不能视为多份独立证据。

每组分析一次，向本组各行写入相同的中文 Context：

```text
【场景结构】
参与者、对白或独白的组织方式、明确的交流媒介与段内切换。

【情境与基调】
本段事件与发展、情绪和语气变化、需要保留的幽默、反讽、双关或线索。
```

当前英文决定剧情事实，人物背景帮助理解语气与关系，不把后续剧情提前写进台词。情绪 `thinking` 或角色后缀 `_remote` 不单独证明内心独白，气泡字段缺失时不默认媒介。KR、JP、FR 附录不自动成为德语规则。人物资料冲突、缺句或关键字段缺失时保留必要的不确定性。

预览按组展示完整结果，可展开同组英文、已核验的关联名称／跨段线索、模型引用的对白和人物资料原文。对白引用须为英文中的连续片段；人物资料引用允许忽略 Markdown 行引用符、行内代码、加粗标记及排版空白，文字本身仍必须连续匹配，不能改写。语义解释仍需在本步人工确认。表格只写原有 `Context` 列，原文、已有译文与其他字段保持；UI、集卡和装饰物仍使用原文关联方式。现有术语筛选继续读取本行 Context。

场景分析沿用 `config/config.json` 的 `llm` 配置，临时网络、限流及服务端错误最多尝试 3 次。无可靠分组、同段源文异常、分类冲突或完整 Key 重复时留空待确认；模型失败、输出格式或引用异常时本步失败，不发布本次不完整结果。通过校验的场景保存在批次的 `context_scenes/`，重试时复用相同输入的成功结果；缓存绑定全部资料、对白、提示词、模型和接口，任一变化都会重新分析。失败响应单独保存供诊断，只有重新通过全部校验后才能复用。选择「重跑本步」则清空本步场景缓存和旧 Context，重新分析；下游产物作废，规范化与分类结果保留。纯非对白批次不调用场景模型。

旧批次不会自动重算。可在更新服务后重跑上下文及下游，或复制批次保留原产物再运行。底层 `ContextProcessor.process` 保留原关联模式，工作台和命令行适配器显式启用场景分析；人物资料由工作台当前项目资料库传入。

## TM 语义匹配（2026-10-22）

第 05 节点保留原文完全一致和正文一致的优先通道；其余句子使用英文本身进行语义向量与 BM25 关键词混合召回，再调用重排模型。沿用 RO 的默认参数：70% 向量 + 30% 关键词（BM25 最高分大于 5 时启用关键词分数），召回 30 条，重排分数至少 0.65，最多选择 5 条历史英文。分数不是翻译正确率。历史原文存在多个德语版本时保留全部版本并提示冲突，所以参考条目可能超过 5 条。

使用 `config/config.json` 现有的 `embedding` 和 `rerank` 配置（各自的 `api_url`、`api_key`、`model_name`）；`embedding` 可配置 `batch_size`、`max_workers`、`timeout`，`rerank` 可配置 `recall_k`、`final_k`、`threshold`、`max_workers`、`timeout`、`request_retries`、`retry_backoff`。无须另传 NPY。配置完整仅表示可以发起请求，不保证外部接口可用。

首次处理非精确命中的句子时自动生成记忆库向量；后续复用 `workspace/pipeline/tm_cache/` 中的缓存（命令行默认 `workspace/tm_cache/`）。缓存绑定 TMX 内容、英文顺序、向量模型、接口和处理版本，修改任一项会重建，不复用 RO 向量或旧上传目录。上传 TMX 只校验语料，不立即调用模型。全为精确命中的批次无需模型请求。

输出仍是一个 `TM参考` 列和原有 `_ready.xlsx` 文件。参考保留原始英德文本、标签和参数，并提示数字、否定、参数或格式差异；不会回填目标译文。当前与 RO 一样，只用源句匹配，Context、分类、术语不参与重排，交给后续翻译环节结合判断。

临时网络错误最多尝试 3 次，响应漏项、零向量、维度不一致等会明确失败，不用空参考冒充未命中，不导出本次不完整结果。重试会复用已完成的整库向量缓存，查询向量和重排重新执行；尚未实现逐行恢复。重跑 TM 本步会清除旧的第 05 步产物，避免失败时下载到旧 ready；第 00–04 步保留。更新服务后，可从已完成批次的 TM 节点选择「重跑本步」应用新匹配。

## 数据落盘

```
workspace/pipeline/
  projects.json                        项目、资料库路径、当前项目
  projects/<project_id>/library/       项目上传的资料
  projects/<project_id>/batches/<batch_id>/
    source.xlsx                        源文件
    00-规范化.xlsx … 05-TM.xlsx        各节点产物
    state.json                         节点状态、统计、预览、确认时间
```

`state.json` 在节点边界原子写入。服务重启时，正在执行的节点会被标成失败，可重试。UI 通过轮询 `state` 接口跟进长节点（术语筛选通常最慢）。

## 关键接口

```
GET  /api/pipeline/bootstrap
POST /api/pipeline/projects
POST /api/pipeline/projects/<id>/activate
POST /api/pipeline/projects/<id>/library/<slot>       # termbase|tm|cast|relationships|world|style
POST /api/pipeline/projects/<id>/batches
GET  /api/pipeline/projects/<id>/batches/<batch>/state
POST /api/pipeline/projects/<id>/batches/<batch>/run  # {stage?, mode?: run|rerun|retry, then_continue?}
POST /api/pipeline/projects/<id>/batches/<batch>/confirm
POST /api/pipeline/projects/<id>/batches/<batch>/clone      # 保留旧批次，新建同源待处理批次
POST /api/pipeline/projects/<id>/batches/<batch>/reset
GET  /api/pipeline/projects/<id>/batches/<batch>/final      # 05-TM.xlsx，最终准备产物
GET  /api/pipeline/projects/<id>/batches/<batch>/artifact/<stage>
```

## 下载排查

- 下载不会弹「另存为」框，文件直接进浏览器下载目录；先看 Downloads。
- 优先用阶段面板右上角的「下载产物」，它始终在可视区；底部的「下载产物」在长面板里可能被滚出视口。
- 如果浏览器或内嵌预览拦截下载，点旁边的「复制链接」，把地址粘贴到独立浏览器的新标签页打开。
- 页面已返回 `Cache-Control: no-store`；如果界面像旧版，用 `http://127.0.0.1:5001/pipeline?v=2` 或硬刷新。

## 当前限制

- 只接 WC 的 **en → de**；新项目语言对固定，暂不支持其它语种。
- 主工作台负责翻译前准备；翻译交给外部工具，返回文件可进入独立的标签清洗和译后 QA 页。QA 检查结构与一致性，不代表语言质量或客户验收。
- 单机单用户，无登录；后台任务在 Flask 进程内跑，不要用多 worker 部署。
- 活动对白场景分析与术语筛选需要 `config/config.json` 里有效的 `llm` 配置；没有配置时对应节点会失败并给出提示。
- TM 的非精确匹配需要有效的向量和重排模型配置；首次生成整库向量会耗时并产生接口费用。
- 准备自检暂时停用；恢复后它的“长度余量”按源文达到 limit 的 85% 计为告警，阈值在 `core_tools/preflight.py` 顶部。

## 测试

```bash
.venv/bin/python -m unittest discover -s tests -v
```

`tests/test_preflight.py` 覆盖独立保留的九条断言；`tests/test_pipeline_service.py` 用假节点覆盖当前六节点的逐节点确认、重试并继续、重跑作废下游和重置。

## 目录与历史数据

- `core_tools/`：六步准备、标签清洗、QA 与可选自检的处理模块；旧页面移除不影响这些核心功能。
- `templates/`、`static/`：三个保留入口及其共用布局、样式。
- `profiles/`：分类与语向配置；`config/config.json` 为本机模型配置，不进 Git。
- `prompts/`：当前提示词源文件、生成脚本和完整版本；具体使用方法见 [提示词说明](prompts/README.md)。
- `workspace/pipeline/`：项目资料库、源文件和批次结果，不进 Git。
- `uploads/`：标签清洗与 QA 的本地运行文件，不进 Git。
- `.venv/`：运行依赖，不能当作历史数据删除；`tests/` 为验证代码。

2026-10-21 已将旧单页数据、旧演示稿、交接包草案、提示词初稿及五个旧/空批次移至仓库外 `../../../WC_Tools归档/`。归档含逐文件 SHA-256 清单与清理前源码，可按需恢复；当前保留两份包含完整对白 Key 的已完成批次。暂时停用的准备自检仍保留代码和测试。
