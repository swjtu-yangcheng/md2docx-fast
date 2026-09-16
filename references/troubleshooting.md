# md2docx 排错手册（2026-09-16 实战记录）

## 一、性能瓶颈剖析（为什么旧做法慢）

| 瓶颈 | 实测 | 说明 |
|---|---|---|
| 每步一次 `python edsdk.py call` | 0.6-0.8s/次 | 其中 python 冷启动 0.58s，真正 RPC 仅几十毫秒 |
| 反复 `doc_find` 重新定位 | 每次 ~17ms × 多次 | 插入后索引漂移，旧流程只能查一次、插一次 |
| 中途 `schema` 查询 | 0.5-0.8s/次 | 纯属浪费，参数本可预置 |
| `doc_to_image` 云端渲染 | 数十秒～失败 | 中文路径有编码 bug，云端服务还不稳定 |
| Agent 每步一次工具往返 | 显著 | 逐步调用法需要几十轮对话 |

**核心优化**：把整条流水线收进**一个 python 进程**，`import edsdk` 复用其 `_rpc`（HTTP JSON-RPC），
进程启动只付一次；索引只解析一次、按逆序/已知偏移插入；不做云端渲染。

## 二、已修复的 Bug

### Bug 1：`doc_insert_toc` 打崩 editor_sdk（最严重）
- 现象：`doc_insert_toc` 报 `[WinError 10054] 远程主机强迫关闭了一个现有的连接`，
  随后所有调用变成 `10061 目标计算机积极拒绝` —— 服务进程已退出。
- 定位：在「刚插入 `doc_insert_page_break` 的同一索引处」再插入目录域时触发。
- 修复：改用顺序 `insert_text(目录标签) → insert_paragraph → insert_toc → doc_find(锚点) → insert_page_break(锚点) → insert_page_break(标签前)`。
- 影响：旧流程每步间隔 0.8s，一旦踩中就表现为「后面步骤莫名其妙全失败」。

### Bug 2：create_doc 竞态
- 现象：create 成功返回 file_id，紧接着的操作全报 `[-1]DocEditor::*: document is not open`。
- 原因：文档真正 open 有约 0.5s 延迟；旧流程因每步耗时 0.8s 而侥幸掩盖。
- 修复：`wait_ready()` 轮询 `doc_get_outline`（100ms 间隔，上限 15s）+ 任意调用遇
  `is not open / not opened` 自动等待后重试一次。

### Bug 3：字段名误判（静默失效，不报错）
- 标题文本字段是 **`text_preview`**（并非 `text`），且会被截断到约 3-4 字 → 目录锚点匹配失败、目录整段被跳过。
- `doc_list_tables` 的列数是 **`col_count`**（不是 `cols`/`columns`）→ 表头加粗静默不执行。
- `doc_find` 返回 **`{"locations":[{begin,end,...}]}`**（不是 `ranges`）→ 解析失败。
- `heading_level` 可能返回字符串 → 需 `int()` 归一。

### Bug 4：`schema doc_insert_toc` 触发服务内部异常
- `list object has no attribute get`；工具本身可正常调用，别查它的 schema。

### Bug 5：`doc_to_image` 中文路径编码 bug + 云端不稳定
- 中文路径报 file not found；复制到 ASCII 路径后仍可能因云端服务繁忙失败。
- 结论：**不做图片校验**，改用 zipfile 直读 XML 校验结构。

## 三、服务挂掉如何恢复

`editor_sdk` 由 WorkBuddy 宿主拉起，崩溃后**不会自愈**。手动恢复：

```bash
# 二进制位置（可用环境变量 EDITOR_SDK_EXE 覆盖）
D:\Users\yangc\AppData\Local\Programs\WorkBuddy\resources\app.asar.unpacked\
  node_modules\@tencent\tencent-docs-ai-engine\bin\win32-x64\editor_sdk.exe --port 39099
```
- 以 `DETACHED_PROCESS|CREATE_NEW_PROCESS_GROUP` 后台启动，cwd 设为 exe 所在目录；
- 探测 `POST http://127.0.0.1:39099/mcp` 的 `tools/list`，约 **0.5-0.8s 就绪**；
- 脚本 `ensure_service()` 已内置该逻辑，无需手工操作。
- 注意：服务在空闲一段时间后可能自行退出，流水线开头务必先探测。

## 四、字段名速查

| 工具 | 关键返回/入参 |
|---|---|
| `create_doc` | 文本里 `file_id=xxx`（非 JSON） |
| `doc_resolve_document_structure` | nodes[].`type/heading_level/start_index/end_index/text_preview` |
| `doc_insert_markdown` | `markdown=file://绝对路径` |
| `doc_insert_toc` | `idx`, `max_level`（别查 schema） |
| `doc_find` | `{"locations":[{"begin","end","related_text"}]}` |
| `doc_list_tables` | `{"tables":[{"table_id","col_count",...}]}` |
| `doc_set_table_properties` | `borders/cell_fills[{condition:"first_row",fill}]` / `cell_margin` |
| `doc_set_table_cells` | `cells:[{row,col,text_format:{bold}}]` |
| `save_file` | `file_path`（中文路径可用），最慢约 1-2s |
