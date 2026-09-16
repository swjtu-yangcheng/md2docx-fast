---
name: md2docx-fast
description: Markdown/长文报告 → 排版美观 Word(.docx) 的**极速**流水线。单进程批处理调用本地 editor_sdk，自动插入目录、分级标题样式、表格美化、页眉页码，含服务崩溃自愈与竞态规避。当用户要求把 md/文本报告写成 word 文档、转换 docx、排版报告、批量生成 Word 时使用。
agent_created: true
---

# md2docx-fast —— Markdown → 排版 Word（快速流水线）

## 何时用
- 把 `.md` / 长文报告转成**排版清晰美观**的 `.docx`（目录、标题层级、表格、页眉页码齐全）。
- 对**速度**敏感：本流水线约 **5-6 秒**完成 2 万字符报告的转换（旧逐步调用法需数分钟）。

## 一条命令
```bash
python "C:/Users/yangc/.workbuddy/skills/md2docx-fast/scripts/md2docx_fast.py" \
       "<源.md>" "<输出.docx>" \
       [--title "页眉文字"] [--toc-anchor 摘要] [--toc-level 3] [--render]
```
只需 **1 次工具调用**即可产出成品；不要拆成多次 edsdk 调用。

参数说明：
| 参数 | 默认 | 说明 |
|---|---|---|
| `md_path` | 必填 | 源 Markdown 绝对路径 |
| `out_path` | 必填 | 输出 .docx 绝对路径（中文路径可用） |
| `--title` | 首个一级标题 | 页眉文字 |
| `--toc-anchor` | `摘要` | 目录页插在该标题之前 |
| `--toc-level` | 3 | 目录层级 |
| `--render` | 关 | 额外渲染逐页图片（云端转换，慢且不稳定，默认不要开） |

路径可用环境变量覆盖：`EDSDK_DIR`、`EDITOR_SDK_EXE`、`EDITOR_SDK_PORT`。

## 流水线（脚本内 9 步，35 次 RPC）
1. `ensure_service` — 服务存活探测，挂了就地拉起 `editor_sdk.exe --port`（冷启动 ~0.7s）
2. `create_doc` + `wait_ready` — 规避「create 后文档尚未 open」竞态
3. `doc_insert_markdown`（`markdown=file://路径`，整篇一次写入）
4. `doc_resolve_document_structure` — **只解析一次**，后续全部基于索引计算
5. 标题样式：按级别**批量下发 ranges**（H1-H4 各 2 次调用，而非逐段）
6. 目录页：标签 → 空段 → 目录域 → 定位分页 → 标签前分页（顺序不可改，见下）
7. `doc_set_document_style` — 宋体 12pt / 1.5 倍行距 / 页边距
8. 表格：`doc_list_tables` → 逐表 `doc_set_table_properties` + 表头加粗
9. 页眉 + 页码 → `save_file`（最慢一步，约 1-2s）

## 必须遵守的 5 条硬约束（踩过的坑）
1. **不要在刚插入分页符的同一索引处插目录域** —— 会让 editor_sdk 进程崩溃（连接被强制关闭 → 端口拒绝 → 后续全部失败）。正确顺序：`insert_text(标签)` → `insert_paragraph` → `insert_toc` → `doc_find(锚点)` → `insert_page_break`。
2. **create_doc 后必须等文档就绪**（~0.5s 竞态窗口），否则批量报 `document is not open`。脚本已内置 `wait_ready` 轮询 + 「not open」自动重试。
3. **别用 `schema doc_insert_toc`** —— 该工具的 schema 查询会触发服务内部异常（`list object has no attribute get`），但工具本身可正常调用。
4. **字段名坑**：标题文本是 `text_preview`（可能被截断，匹配需双向包含）；`doc_list_tables` 的列数是 `col_count`；`doc_find` 返回的是 `locations` 而非 `ranges`。
5. **不要用 `doc_to_image` 做校验**（中文路径有编码 bug，云端转换常失败）。改用 zipfile 直读 `word/document.xml` 校验段落/表格/TOC/页眉/页码，秒级完成。

## 产出校验（推荐，1 次调用）
```python
import zipfile
z = zipfile.ZipFile(out); xml = z.read('word/document.xml').decode('utf-8','replace')
print('段落', xml.count('<w:p ')+xml.count('<w:p>'), '表格', xml.count('<w:tbl>'),
      'TOC', 'TOC' in xml, '页眉', any(n.startswith('word/header') for n in z.namelist()))
```

## 性能对比（2.1 万字符报告，实测）
| 方案 | RPC 次数 | 进程启动次数 | 总耗时 |
|---|---|---|---|
| 旧：逐步 `python edsdk.py call ...` | ~35 | ~35（每次 0.6s 冷启动） | 数分钟（含失败重试/渲染） |
| 新：单进程批处理 | 35 | 1 | **4.7-6.2 s** |

单步耗时参考：insert_markdown ~0.2s、insert_toc ~0.17s、save_file 1.4-2.2s（最慢）、其余多为 3-30ms。

## 排错
详见 `references/troubleshooting.md`（含服务重启方法、崩溃定位过程、字段名速查）。
