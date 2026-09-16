# md2docx-fast

Markdown / 长文报告 → 排版美观 Word（.docx）的**极速**流水线。

基于 WorkBuddy 内置 `editor_sdk`（本地腾讯文档编辑内核），**单进程批处理**调用：
自动目录、分级标题样式、表格美化、页眉页码一次成型。
实测 2.1 万字符报告：**35 次 RPC、约 3–6 秒**完成（逐步调用法需数分钟）。

## 快速开始

```bash
python scripts/md2docx_fast.py "<源.md>" "<输出.docx>" \
       [--title "页眉文字"] [--toc-anchor 摘要] [--toc-level 3] [--render]
```

| 参数 | 默认 | 说明 |
|---|---|---|
| `md_path` | 必填 | 源 Markdown 绝对路径 |
| `out_path` | 必填 | 输出 .docx 绝对路径（中文路径可用） |
| `--title` | 首个一级标题 | 页眉文字 |
| `--toc-anchor` | `摘要` | 目录页插在该标题之前 |
| `--toc-level` | 3 | 目录层级 |
| `--render` | 关 | 额外渲染逐页图片（云端转换，慢且不稳定，建议不开） |

路径可用环境变量覆盖：`EDSDK_DIR`、`EDITOR_SDK_EXE`、`EDITOR_SDK_PORT`（默认 39099）。

## 流水线

1. `ensure_service` — 服务存活探测，挂了就地拉起 `editor_sdk.exe --port`
2. `create_doc` + `wait_ready` — 规避「create 后文档尚未 open」竞态
3. `doc_insert_markdown`（整篇一次写入）
4. `doc_resolve_document_structure` — 只解析一次，后续全部基于索引计算
5. 标题样式：按级别批量下发 ranges
6. 目录页：标签 → 空段 → 目录域 → 定位分页 → 标签前分页
7. `doc_set_document_style` — 宋体 12pt / 1.5 倍行距 / 页边距
8. 表格：统一边框、表头浅蓝底纹 + 加粗
9. 页眉 + 页码 → `save_file`

## 踩过的坑（详见 references/troubleshooting.md）

1. **不要在刚插入分页符的同一索引处插目录域** —— 会让 editor_sdk 进程崩溃（WinError 10054 → 服务退出 → 后续全部 10061）。
2. **`create_doc` 后有约 0.5s 竞态窗口**，直接操作会报 `document is not open`。
3. **字段名坑**：标题文本是 `text_preview`（且会截断）、表列数是 `col_count`、`doc_find` 返回 `locations` 而非 `ranges`、`heading_level` 可能是字符串。
4. **别查 `doc_insert_toc` 的 schema** —— 触发服务内部异常，但工具本身可正常调用。
5. **不要用 `doc_to_image` 做校验** —— 中文路径有编码 bug 且云端不稳；改用 zipfile 直读 `word/document.xml`。

## 性能对比（2.1 万字符报告，实测）

| 方案 | RPC 次数 | 进程启动次数 | 总耗时 |
|---|---|---|---|
| 逐步 `python edsdk.py call ...` | ~35 | ~35（每次 0.6s 冷启动） | 数分钟 |
| 单进程批处理（本方案） | 35 | 1 | **3–6 s** |

## 目录结构

```
md2docx-fast/
├── SKILL.md                        # 供 Agent 调用的 skill 定义
├── README.md
├── scripts/md2docx_fast.py         # 流水线主脚本
└── references/troubleshooting.md   # 排错手册（含服务恢复与字段名速查）
```
