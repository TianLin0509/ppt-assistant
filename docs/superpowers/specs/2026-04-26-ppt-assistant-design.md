# PPT Assistant — 单页 PPT 填充助手设计文档

> 日期: 2026-04-26
> 状态: 已审查，待实现

## 1. 项目定位

本地 Python 工具，辅助用户精雕细琢单页 PPT。**模板决定版式，工具只做内容填充**。

核心约束：
- 公司电脑无法访问任何 AI API
- 不能上传 .pptx 到 AI Web（合规要求）
- 通过剪贴板桥接 AI Web（Claude/Gemini/豆包）做纯文字对话和图片生成
- 单人本地使用，双机部署（家用电脑开发测试，公司电脑实践）

## 2. 技术栈

| 组件 | 选型 | 理由 |
|------|------|------|
| UI | Streamlit | 单人本地工具，1 天搭完；`st.fragment` 解决性能和定时保存 |
| PPT 操作 | python-pptx + lxml | lxml 处理 python-pptx 无法直接操作的底层 OOXML 属性 |
| 预览渲染 | pywin32 COM（Phase 1） | 调用本机 PowerPoint 导出 PNG；Phase 2+ 加 LibreOffice headless 兜底 |
| 图像处理 | Pillow | 画红框标注、缩略图 |
| 数据模型 | Pydantic v2 | 强类型 + JSON 序列化 |
| Prompt 模板 | Jinja2 | 比字符串拼接更清晰 |
| 剪贴板 | st-copy-to-clipboard | Streamlit 复制按钮 |
| 存储 | 本地文件系统 + JSON | 不用数据库 |

## 3. 架构与数据流

### 3.1 模板预处理（一次性，每个模板做一次）

```
.pptx 模板
  → pptx_parser（递归遍历 shape，提取元数据 + FontInfo 快照）
  → pptx_renderer（pywin32 COM → PNG）
  → shape_annotator（PIL 红框 + 编号）
  → role_inferencer（字号/位置/类型规则推断 ~70%）
  → 用户 Review（Streamlit 修正表单，逐个确认角色）
  → 输出：metadata JSON + shape.name 加角色后缀
```

### 3.2 单次任务执行

```
用户选模板 + 写任务描述
  → text_prompt_builder（Jinja2 渲染 prompt）
  → 用户复制 prompt → AI Web → 拿到 JSON → 粘贴回工具
  → json_parser（TD-5 宽容解析）
  → 用户在 candidate_picker 中逐项挑选备选
  → pptx_filler（TD-1 run 级安全替换文本）
  → pptx_renderer（COM → 预览 PNG）
  → 用户检查预览 → 满意则导出 .pptx
  → 用户在 PowerPoint 做最后精修
```

## 4. 目录结构

```
ppt-assistant/
├── .gitignore
├── README.md
├── requirements.txt
├── config.yaml                        # 可移植配置
│
├── templates/                         # 用户的 .pptx 模板
│   └── .gitkeep
├── templates_meta/                    # 预处理产出
│   └── .gitkeep
├── runs/                              # 每次任务的完整产物 (TD-6)
│   └── .gitkeep
├── prompts/                           # AI Prompt 模板
│   ├── text_generation.md.j2
│   └── image_brief.md.j2
│
├── src/
│   ├── __init__.py
│   ├── schema.py                      # Pydantic v2 数据模型
│   ├── core/
│   │   ├── __init__.py
│   │   ├── pptx_parser.py            # 解析 .pptx → TemplateMeta
│   │   ├── pptx_filler.py            # 填回模板 (TD-1, TD-2, TD-4)
│   │   ├── pptx_renderer.py          # pywin32 COM → PNG
│   │   ├── shape_annotator.py        # PIL 红框 + 编号
│   │   └── role_inferencer.py         # 本地规则推断角色
│   ├── prompt/
│   │   ├── __init__.py
│   │   ├── text_prompt_builder.py    # 生成文本 prompt
│   │   ├── image_prompt_builder.py   # 生成图片 brief
│   │   └── json_parser.py            # TD-5 宽容解析
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── app.py                     # Streamlit 主入口
│   │   ├── pages/
│   │   │   ├── 1_template_library.py
│   │   │   ├── 2_new_task.py
│   │   │   └── 3_task_history.py
│   │   └── components/
│   │       ├── __init__.py
│   │       ├── candidate_picker.py
│   │       └── template_card.py
│   └── utils/
│       ├── __init__.py
│       ├── state_manager.py           # TD-7 自动保存
│       └── config.py                  # 配置加载
│
├── docs/
│   └── superpowers/
│       └── specs/                     # 设计文档
│
└── tests/
    ├── __init__.py
    ├── conftest.py
    ├── samples/
    │   └── .gitkeep
    ├── test_pptx_parser.py
    ├── test_pptx_filler.py
    └── test_json_parser.py
```

## 5. 数据模型（审查定稿版）

### 5.1 枚举

- `ShapeType`: text / image / decoration / group / smartart / chart / table / unknown
- `TextSubtype`: title（整段替换）/ bullet（按行拆 paragraph）/ body（整段替换）

### 5.2 BBox

shape 的位置和尺寸，EMU 单位。字段：left, top, width, height（均 int）。

### 5.3 FontInfo

单个 run 的格式快照，TD-1 安全替换核心结构。

| 字段 | 类型 | 说明 |
|------|------|------|
| name | Optional[str] | 西文字体名 |
| name_east_asian | Optional[str] | CJK 字体名（rPr ea 属性） |
| size_pt | Optional[float] | 字号（磅） |
| bold | Optional[bool] | 粗体 |
| italic | Optional[bool] | 斜体 |
| color_rgb | Optional[str] | "#RRGGBB" 格式 |
| underline | Optional[bool] | 下划线 |

### 5.4 ImageSlotInfo

图片位的裁剪和旋转信息（Phase 2 使用）。字段：rotation, crop_left/right/top/bottom（float），aspect_ratio（Optional[str]）。

### 5.5 ShapeRole

单个 shape 的完整元数据，核心模型。

| 字段组 | 字段 | 说明 |
|--------|------|------|
| 标识 | shape_id, shape_name_original, shape_name_with_role, role_key, role_zh, role_confirmed | role_confirmed 标记用户是否已确认 |
| 类型 | type (ShapeType), is_editable | 默认可编辑，group/smartart/chart/table 设为 False |
| 定位 | bbox (BBox), text_hash | TD-4 多重定位锚点 |
| 文本 | text_subtype, max_chars, max_lines, current_content, first_run_font, paragraph_fonts, paragraph_count | paragraph_fonts: bullet 类型保存每 paragraph 首 run 格式 |
| 图片 | image_slot (ImageSlotInfo) | Phase 2 |
| Group | is_in_group, group_path, z_order_index | group_path 如 "01-02" |
| 计算属性 | display_label（@property，不序列化） | UI 显示标签 |

### 5.6 TemplateMeta

单个模板的完整元数据，存储到 `templates_meta/{template_id}.json`。

| 字段 | 说明 |
|------|------|
| template_id | 文件名去后缀，空格替换为下划线，保留中文 |
| file_path | 绝对路径 |
| file_mtime | 修改时间（一致性检查） |
| preview_image, annotated_image | PNG 路径 |
| slide_width_emu, slide_height_emu | slide 尺寸 |
| elements: list[ShapeRole] | 所有 shape 的元数据 |
| editable_text_elements（@property） | 可编辑文本元素（不序列化） |
| editable_image_elements（@property） | 可编辑图片元素（不序列化） |

### 5.7 TextCandidates

AI 返回的文本生成结果（json_parser 解析后）。字段：template_id, task_description, candidates: dict[str, list[str]]。

注意：schema 保持严格类型，宽容解析在 json_parser.py 层面处理。

### 5.8 TaskRun

一次任务的完整记录，存储到 `runs/{task_id}/task.json`。

| 字段 | 说明 |
|------|------|
| task_id | "2026-04-26_XX算法介绍" 格式 |
| created_at | ISO 格式时间字符串 |
| status | 任务状态：created / prompting / selecting / rendering / completed / failed |
| current_step | 当前步骤编号 (0-6) |
| task_description | 任务描述 |
| template_id | 使用的模板 ID |
| template_mtime | 模板当时的修改时间（冗余快照，检测模板变更） |
| text_prompt, image_prompt | 生成的 prompt |
| ai_response_raw | AI 返回原文 |
| text_candidates | 解析后的备选 |
| text_choices | role_key → 选定文案 |
| image_choices | role_key → 图片相对路径 |
| output_pptx, preview_image | 产出路径 |
| run_dir（@property） | 任务产物目录相对路径（不序列化） |

### 5.9 AppConfig

可移植配置，从 config.yaml 加载。

| 字段 | 默认值 | 说明 |
|------|--------|------|
| templates_dir | "templates" | 模板目录（相对路径） |
| templates_meta_dir | "templates_meta" | 元数据目录 |
| runs_dir | "runs" | 任务历史目录 |
| prompts_dir | "prompts" | Prompt 模板目录 |
| powerpoint_path | None (自动检测) | PowerPoint 可执行文件路径 |
| libreoffice_path | None | Phase 2+ |
| auto_save_interval_sec | 10 | 自动保存间隔 |
| candidates_per_element | 3 | 每元素备选数量 |

## 6. 关键技术决策

### TD-1：文本替换的 run 级安全策略

按 text_subtype 分三种策略：

- **title / body**：保留第一个 run 的格式（first_run_font），删除多余 run，写入新文本
- **bullet**：按 `\n` 拆行，每行对应一个 paragraph，从 paragraph_fonts 继承对应格式。多出的行复用最后一个格式，少的行删除多余 paragraph

### TD-2：图片替换保留 Z-Order 和裁剪（Phase 2）

记录原 shape 在 spTree 中的 z_order_index，替换后用 lxml 移回原位置。crop/rotation 从 ImageSlotInfo 恢复。

### TD-3：Group Shape 递归处理

遍历时递归展开 group 内部 shape，编号为 `01-01`、`01-02`。**第一版不允许编辑 group 内部 shape**，UI 标红提示。SmartArt/Chart/Table 同理。

### TD-4：shape 多重定位

回填时按优先级匹配：shape.name 后缀 → shape_id → bbox + text_hash 模糊匹配 → 全部失败则 UI 提示。

### TD-5：JSON 宽容解析

解析链：提取 ```json 代码块 → 提取首尾花括号 → 修复尾逗号 → json.loads → ast.literal_eval → 返回 None 让用户手动编辑。

### TD-6：中间产物全部落盘

```
runs/{task_id}/
├── task.json          # TaskRun 完整状态
├── prompts/
│   ├── text_prompt.txt
│   └── image_prompt.txt
├── ai_response_raw.txt
├── candidates.json
├── selections.json
├── images/
├── preview.png
└── output.pptx
```

### TD-7：状态自动保存

`@st.fragment(run_every=10)` 定时序列化 TaskRun 到 JSON。启动时自动检测并恢复未完成任务。

## 7. Phase 1 详细任务拆解

**目标**：跑通"模板预处理 → prompt 生成 → JSON 粘贴 → 文本填充 → 导出 .pptx"完整链路。

**验收标准**：用 1 个野生模板，输入任务描述，粘贴手编 JSON，导出的 .pptx 文字格式不丢失。

### T1：项目骨架（0.5h）

- requirements.txt: streamlit, python-pptx, lxml, Pillow, pywin32, pydantic, Jinja2, st-copy-to-clipboard
- config.yaml: 默认配置
- src/utils/config.py: `load_config() -> AppConfig`
- src/schema.py: 全部 Pydantic 模型（审查修复后版本）
- src/ui/app.py: Streamlit 主入口 + 侧边栏导航

### T2：模板解析器（1h）— src/core/pptx_parser.py

| 函数 | 职责 |
|------|------|
| `parse_template(pptx_path) -> TemplateMeta` | 主入口，遍历第一张 slide 所有 shape |
| `_parse_shape(shape, z_index, group_path) -> ShapeRole` | 单 shape 解析：类型判断、bbox、文本、FontInfo 快照 |
| `_parse_group(group_shape, z_index, parent_path) -> list[ShapeRole]` | 递归展开 group |
| `_classify_shape_type(shape) -> ShapeType` | MSO_SHAPE_TYPE 映射 |
| `_snapshot_font(run) -> FontInfo` | 提取 run 格式（含 CJK 字体 ea 属性） |
| `_compute_text_hash(text) -> str` | MD5 摘要 |

### T3：pywin32 渲染器（1h）— src/core/pptx_renderer.py

| 函数 | 职责 |
|------|------|
| `render_slide_to_png(pptx_path, output_png, slide_index=0) -> str` | COM 导出 slide 为 PNG |
| `_get_powerpoint_app() -> Dispatch` | 获取/启动 PowerPoint COM 单例 |
| `_ensure_powerpoint_closed()` | atexit 注册安全关闭 |

要点：必须 `pythoncom.CoInitialize()`；失败抛明确异常不静默降级。

### T4：形状标注器（0.5h）— src/core/shape_annotator.py

| 函数 | 职责 |
|------|------|
| `annotate_preview(preview_png, elements, slide_w_emu, slide_h_emu) -> str` | 红框 + 编号标注 |
| `_emu_to_pixel(emu_val, emu_total, pixel_total) -> int` | 坐标转换 |

跳过 DECORATION 和 GROUP 容器本身。

### T5：角色推断器（0.5h）— src/core/role_inferencer.py

| 函数 | 职责 |
|------|------|
| `infer_roles(elements, slide_w, slide_h) -> list[ShapeRole]` | 规则推断 + role_confirmed=False |

规则：字号 ≥24pt 且上 1/4 → title_main；含换行 → bullet；图片 → image_*；面积 <2% → DECORATION。

### T6：Prompt 生成器（0.5h）— src/prompt/text_prompt_builder.py

| 函数 | 职责 |
|------|------|
| `build_text_prompt(meta, task, n_candidates=3) -> str` | Jinja2 渲染文本 prompt |

### T7：JSON 宽容解析器（0.5h）— src/prompt/json_parser.py

| 函数 | 职责 |
|------|------|
| `parse_ai_json(raw_text) -> dict \| None` | TD-5 完整实现 |
| `_extract_code_block(text) -> str \| None` | 提取代码块 |
| `_extract_braces(text) -> str` | 提取首尾花括号 |
| `_fix_trailing_commas(text) -> str` | 正则修复尾逗号 |

### T8：文本填充器（1.5h）— src/core/pptx_filler.py

| 函数 | 职责 |
|------|------|
| `fill_template(pptx_path, meta, text_choices, output_path) -> str` | 主入口：加载 → 匹配 → 替换 → 保存 |
| `_find_shape(slide, element) -> Shape \| None` | TD-4 多重定位 |
| `_replace_title(text_frame, new_text, font_info)` | title 策略 |
| `_replace_bullet(text_frame, new_text, para_fonts)` | bullet 策略 |
| `_replace_body(text_frame, new_text, font_info)` | body 策略 |
| `_apply_font(run, font_info)` | 格式恢复 |

Phase 1 最核心最高风险模块。

### T9：Streamlit 模板库页面（1h）— src/ui/pages/1_template_library.py

- 缩略图墙（3 列 st.columns）
- 预处理按钮 → parse → render → annotate → infer_roles
- 角色修正表单：每元素一行，下拉选 role_key + text_subtype
- 保存 metadata JSON + shape.name 后缀

### T10：Streamlit 新任务页面（1.5h）— src/ui/pages/2_new_task.py

7 步线性流程：选模板 → 写任务 → 生成 prompt（复制按钮）→ 粘贴 JSON → 挑选备选 → 生成预览 → 导出。

### T11：备选挑选组件（0.5h）— src/ui/components/candidate_picker.py

`@st.fragment` 包裹每个元素块。Radio + text_area 联动，动态 session_state key。一键全选 A 按钮。

### T12：状态管理（0.5h）— src/utils/state_manager.py

`@st.fragment(run_every=10)` 定时 auto_save。启动时检测未完成任务自动恢复。

### T13：集成测试（1h）

- test_pptx_parser.py: shape 数量、类型识别、group 递归
- test_pptx_filler.py: TD-1 三种策略格式不丢
- test_json_parser.py: 正常/代码块/尾逗号/垃圾文本

**总耗时：~10h 编码**

## 8. Phase 1 风险与对策

| 风险 | 对策 |
|------|------|
| Streamlit 备选挑选 UI 交互卡顿 | st.fragment 局部重跑 + 动态 session_state key，Phase 1 验证后评估是否换框架 |
| pywin32 COM 在公司电脑被 IT 策略禁用 | Phase 1 只做 pywin32，Phase 2 加 LibreOffice headless 兜底 |
| 野生模板 shape name 杂乱 | 简单规则推断 ~70% + 用户手动修正确认 |

## 9. Phase 路线图

| Phase | 目标 | 预计耗时 |
|-------|------|---------|
| 1 | 最小闭环：文本填充主链路 | 2 天 |
| 2 | 预处理完善：图片替换、Group 处理、多重定位 | 1.5 天 |
| 3 | 关键体验：图片拖拽上传、校验页、自动预览 | 1.5 天 |
| 4 | 鲁棒性：一致性检查、历史存档、状态保存、LibreOffice 兜底 | 1 天 |
| 5 | 按实际反馈优化 | 持续 |

## 10. 多方审查记录

2026-04-26 schema 四路审查通过。

已修复：computed_field → @property、新增 paragraph_fonts、新增 name_east_asian、删除未使用 import、TaskRun 新增 status/current_step/template_mtime。

已确认不采纳：角色分配 per-task（角色 per-template 正确）、FontInfo 主题色支持（Phase 2）、model_config extra='forbid'（Phase 2）。
