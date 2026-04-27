# Phase 2 Batch 1 Design: Prompt Quality + Style Presets + UX Enhancements

**Date:** 2026-04-27
**Status:** Approved
**Scope:** 4 features, estimated half-day

---

## 1. Style Preset System

### Goal
Built-in PPT style presets that inject writing-style rules into AI prompts. Default: Huawei-style technical report.

### File Structure
```
prompts/
  styles/
    huawei.md      # Huawei technical report style (default)
    minimal.md     # Minimal/clean style (placeholder for future)
```

### Style File Content (huawei.md)
Extracted from the full 11-section Huawei PPT spec, keeping only sections that affect AI text generation:

1. **Title Rules** — conclusion-oriented titles, formula: "Subject + colon + method/action + quantified result", embed metrics like +20%, 1.2X, 35%->60%
2. **Language & Wording** — high-frequency verbs (construct, enable, enhance, optimize, evolve, cover), nouns (architecture, foundation, scenario, gain, mechanism), sentence patterns ("facing XX scenario, construct XX capability", "compared to XX, improved XX%")
3. **Bullet Principles** — concise short sentences, technical terminology, English abbreviations used directly, conclusion-oriented
4. **Negative Constraints** — no marketing slogans, no noun-only titles, no long prose, no vague statements
5. **Layout-Aware Hints** — based on template's layout_type, suggest content structure (e.g., dual-column -> As-Is vs To-Be comparison content)

### Schema Changes

**AppConfig:**
```python
default_style: str = "huawei"
```

**TaskRun:**
```python
style_id: Optional[str] = None
```

**TemplateMeta:**
```python
layout_type: Optional[str] = None  # "dual_column" | "top_summary" | "flow_chain" | "card_grid" | None
```

### UI Changes (2_new_task.py, Step 2)
- Scan `prompts/styles/` for `.md` files
- Dropdown to select style (default from config), option "None" for no style injection
- Selected style_id saved to TaskRun

### Prompt Template Changes (text_generation.md.j2)
- `{% if style_content %}` block before element list: inject style rules
- `{% if layout_type %}` block: describe current layout type with content suggestions

### Template Library UI (1_template_library.py)
- In the role correction section, add dropdown for `layout_type` selection per template
- Saved to TemplateMeta JSON

---

## 2. BBox-Based Character Count Estimation

### Goal
Replace the naive `2x original text length` heuristic with physics-based estimation using text box dimensions and font size.

### Algorithm (in role_inferencer.py)
```python
def estimate_char_capacity(bbox: BBox, font_size_pt: float, slide_w_emu: int, slide_h_emu: int) -> tuple[int, int]:
    """Returns (min_chars, max_chars) suggested range."""
    bbox_width_pt = bbox.width / 12700
    bbox_height_pt = bbox.height / 12700
    
    chars_per_line = bbox_width_pt / font_size_pt
    lines = bbox_height_pt / (font_size_pt * 1.5)
    
    theoretical_max = chars_per_line * lines
    max_chars = int(theoretical_max * 0.85)
    min_chars = int(theoretical_max * 0.55)
    
    return (min_chars, max_chars)
```

**Title special case:** max_chars capped to single line (`chars_per_line * 1`).

### Schema Changes

**ShapeRole:**
- `min_chars: Optional[int] = None` (new field)
- `max_chars` — now populated by bbox estimation instead of 2x heuristic

### Font Size Source Priority
1. `first_run_font.size_pt` (if available)
2. Median of all body elements' font sizes in the template
3. Fallback: 12pt

### Prompt Template Changes
- Old: `最多{{ e.max_chars }}字`
- New: `建议{{ e.min_chars }}-{{ e.max_chars }}字，尽量充实内容不留空白`
- Title elements: append `控制在一行内`

---

## 3. Template Consistency Check

### Goal
Detect if template .pptx file was modified after preprocessing, warn user before generating results on stale metadata.

### Trigger Points
- Step 1: after selecting template
- Step 5: before generating preview

### Behavior
- Compare `TaskRun.template_mtime` vs `os.path.getmtime(template_file)`
- Mismatch → `st.warning("template file has been modified, recommend re-preprocessing")`
- Two buttons: "Re-preprocess" (navigate to template library) / "Ignore and continue"
- Non-blocking: user can proceed at own risk

### Code Changes
- `state_manager.py`: add `check_template_consistency(task: TaskRun) -> bool`
- `2_new_task.py`: call at Step 1 and Step 5

---

## 4. Prompt Retry with Feedback

### Goal
When AI returns poor quality JSON (format errors, too short, wrong style), allow user to generate a revision prompt with specific feedback.

### UI Changes (2_new_task.py, Step 4)
Below the JSON paste area, add "Retry" section:
- `st.text_area("Revision notes")` with quick-fill buttons:
  - "Content too short, fill to suggested character count for each text box"
  - "Title not conclusion-oriented enough, needs quantified data"
  - "Bullets too generic, need specific technical details"
  - Free-form input
- "Regenerate Prompt" button → produces new prompt with revision context

### New Prompt Template (prompts/text_revision.md.j2)
```jinja2
以下是你上次生成的内容（有问题需要修改）：

```json
{{ previous_response }}
```

修改要求：
{{ revision_notes }}

请基于以上要求重新生成，格式不变。
```

### Schema Changes

**TaskRun:**
```python
revision_count: int = 0
revision_notes: Optional[str] = None
```

### File Persistence
- On retry, rename current `ai_response_raw.txt` to `ai_response_raw_v{revision_count}.txt`
- New paste overwrites `ai_response_raw.txt`
- Maintains clipboard-bridging mode (no direct AI API calls)

---

## Files to Create / Modify

### New Files
| File | Purpose |
|------|---------|
| `prompts/styles/huawei.md` | Huawei style preset |
| `prompts/styles/minimal.md` | Minimal style placeholder |
| `prompts/text_revision.md.j2` | Revision prompt template |

### Modified Files
| File | Changes |
|------|---------|
| `config.yaml` | Add `default_style` |
| `src/schema.py` | TaskRun: style_id, revision_count, revision_notes; ShapeRole: min_chars; TemplateMeta: layout_type; AppConfig: default_style |
| `src/core/role_inferencer.py` | Add `estimate_char_capacity()`, update `infer_roles()` to use bbox estimation |
| `src/prompt/text_prompt_builder.py` | Load style file, inject style_content and layout_type into template |
| `prompts/text_generation.md.j2` | Add style and layout_type blocks, update char count display |
| `src/ui/pages/1_template_library.py` | Add layout_type dropdown in role correction |
| `src/ui/pages/2_new_task.py` | Add style dropdown (Step 2), template consistency check (Step 1, 5), retry section (Step 4) |
| `src/utils/state_manager.py` | Add `check_template_consistency()` |

---

## Out of Scope (Phase 2 Batch 2)
- Image replacement (TD-2)
- Image drag-upload UI
- Group shape editing
- LibreOffice headless fallback
