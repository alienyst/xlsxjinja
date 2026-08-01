# Changelog

All notable changes to **xlsxjinja** will be documented in this file.

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

### 🙏 Credits

This project is based on [xltpl](https://github.com/zhangyu836/xltpl) by Zhang Yu.
Licensed under MIT License.

### 📜 License

MIT License

- Original work Copyright (c) 2020 Zhang Yu (xltpl)
- Modified work Copyright (c) 2024 [Your Name] (xlsxjinja)

---

## Future Releases

Track planned features and improvements on [GitHub Issues](https://github.com/yourusername/xlsxjinja/issues).

Potential roadmap:
- [ ] Type hints throughout codebase
- [ ] Async support for large files
- [ ] Streaming mode for memory efficiency
- [ ] CLI tool for command-line usage
- [ ] More examples and tutorials
- [ ] Performance benchmarks
