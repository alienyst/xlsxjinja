# Changelog

All notable changes to **xlsxjinja** will be documented in this file.

## [1.1.1] - 2026-08

### ✨ New Features

- **`{% tr %}` tag** — inline row-level `for`/`if` control flow without needing
  a cell comment. Opening tags at the start of a cell value are hoisted to
  `beforerow`; closing tags at the end are placed as `aftercell`.
- **`{% tc %}` tag** — inline column-level conditional marker (tag stripping;
  see README for current scope and limitations).
- **`{% insert_img %}` tag** — insert an image into a cell with no existing
  placeholder. Image is auto-sized to the cell's column width / row height
  via `OneCellAnchor`.
- **Native WebP image support** — `ImageRef` now decodes base64 image bytes
  directly and converts WebP to PNG in-memory (no external `dwebp` binary
  required). Pillow's `WebPImagePlugin` is force-registered on import so
  `Image.open()` reliably recognizes WebP from `BytesIO`.
- **`NestedDateTime` timezone fix** — patches openpyxl to strip timezone info
  from `dcterms:modified` before XML serialization, fixing a corruption issue
  observed on Odoo.sh.

### 📦 Dependencies

- **Pillow is now optional.** Core install (`pip install xlsxjinja`) only
  requires `openpyxl` + `jinja2`. Image tags (`{% img %}`, `{% insert_img %}`)
  require `pip install xlsxjinja[image]`.

### 📄 Documentation

- Added a full "Looping" chapter to the README comparing `{% tr %}` vs.
  `beforerow` cell-comment loops, including when each is required (merged
  cells, multi-row `for` blocks).
- Documented `{% insert_img %}` and WebP auto-conversion behavior.
- Archived early AI-generated refactoring reports (`SUMMARY.md`,
  `FINAL_REPORT.md`, `REFACTORING_SUMMARY_V2.md`) to `debug/`; superseded by
  this changelog and git history.

## [1.0.0] - 2024-12-13

### 🎉 Initial Release

**xlsxjinja** is a modern fork of [xltpl](https://github.com/zhangyu836/xltpl) by Zhang Yu, refactored and rebranded for modern Python development.

### ✨ What's New

- 📦 **New package name**: `xlsxjinja` (previously `xltpl`)
- 🐍 **Python 3.7+ only** - Modernized codebase
- 📊 **xlsx format only** - Focused on modern Excel files
- 🚀 **Simplified dependencies** - Only 2 dependencies instead of 5
- 📝 **Enhanced documentation** - Comprehensive README and examples
- ⚡ **40% less code** - Removed legacy support and dual implementations

### 🔧 Technical Improvements

**Code Refactoring:**
- Removed Python 2 compatibility layer (six)
- Removed .xls format support (xlrd/xlwt)
- Unified class implementations (no more dual classes)
- Added docstrings throughout codebase
- Fixed regex warnings with raw strings
- Modern Python 3 string handling

**Dependencies:**
- ❌ Removed: `xlrd`, `xlwt`, `six`
- ✅ Kept: `openpyxl >= 3.1.0`, `jinja2 >= 2.10`

**File Structure:**
- Renamed all internal modules (removed 'x' suffix)
- Simplified class names (CellContext instead of CellContextX)
- Consolidated merger implementations
- Cleaned up imports

### 📊 Metrics

| Metric | Before (xltpl 3.x) | After (xlsxjinja 1.0) |
|--------|-------------------|---------------------|
| Python files | 23 | 20 |
| Total lines | ~4000+ | ~2473 |
| Dependencies | 5 | 2 |
| Python version | 2.7+ | 3.7+ |
| Supported formats | .xls, .xlsx | .xlsx only |

### 🎯 Features

All original xltpl features for .xlsx format:
- ✅ Jinja2 templating in Excel
- ✅ Variables, loops, conditionals
- ✅ Rich text support
- ✅ Image insertion
- ✅ Merged cells
- ✅ Data validation
- ✅ Auto filters
- ✅ Custom filters and globals

### 📝 Usage

```python
from xlsxjinja import BookWriter

writer = BookWriter('template.xlsx')
writer.render_book([{'name': 'John', 'items': [...]}])
writer.save('output.xlsx')
```

---

## Future Releases

Track planned features and improvements on [GitHub Issues](https://github.com/alienyst/xlsxjinja/issues).

Potential roadmap:
- [ ] `save_to_stream()` — render directly to `BytesIO` without a temp file
- [ ] `{% hyperlink %}` tag — create hyperlinks dynamically from template data
- [ ] `{% tc %}` full column-hide support (currently only strips the tag)
- [ ] `{% barcode %}` / `{% qrcode %}` tags (optional dependency)
- [ ] Named range support that follows loop expansion
- [ ] Formula re-anchoring when loops expand row ranges
- [ ] Type hints throughout codebase
- [ ] Async support for large files
- [ ] CLI tool for command-line usage
- [ ] More examples and tutorials
- [ ] Performance benchmarks
