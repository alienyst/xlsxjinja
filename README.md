# xlsxjinja

**Generate Excel (.xlsx) files from templates using Jinja2**

xlsxjinja is a modern Python library that allows you to create Excel spreadsheets from template files using the powerful Jinja2 templating engine. Perfect for generating reports, invoices, data exports, and any document that requires dynamic data insertion.

---

## ✨ Features

- 📝 **Jinja2 Templating** - Use familiar `{{ variables }}`, `{% for %}` loops, and `{% if %}` conditionals
- 🎨 **Rich Text Support** - Preserve multiple fonts and formatting within cells
- 🖼️ **Image Insertion** - Dynamically insert images into cells
- 📊 **Data Validation** - Preserve dropdowns and validation rules
- 🔗 **Merged Cells** - Automatic handling of merged cell ranges
- 🔍 **Auto Filters** - Maintain filter settings
- 🚀 **Modern Python** - Built for Python 3.7+, type-hint ready
- ⚡ **Fast & Lightweight** - Only 2 dependencies (openpyxl + jinja2)

---

## 📦 Installation

```bash
pip install xlsxjinja
```

### Requirements

- Python 3.7+
- openpyxl >= 3.1.0
- jinja2 >= 2.10

---

## 🚀 Quick Start

### 1. Create a Template

Create an Excel file (`template.xlsx`) with Jinja2 syntax:

| Name | Price |
|------|-------|
| `{{item.name}}` | `{{item.price}}` |

Add a comment to the row with: `{%- for item in items %}`
And another comment after: `{%- endfor %}`

### 2. Generate from Template

```python
from xlsxjinja import BookWriter

# Load template
writer = BookWriter('template.xlsx')

# Prepare data
data = {
    'items': [
        {'name': 'Apple', 'price': 10},
        {'name': 'Orange', 'price': 15},
        {'name': 'Banana', 'price': 8},
    ]
}

# Render
writer.render_book([data])

# Save
writer.save('output.xlsx')
```

### 3. Result

| Name | Price |
|------|-------|
| Apple | 10 |
| Orange | 15 |
| Banana | 8 |

---

## 📖 Documentation

### Debug Mode

xlsxjinja includes comprehensive logging for debugging template issues:

```python
from xlsxjinja import BookWriter

# Enable debug logging
writer = BookWriter('template.xlsx', debug=True)

# Now you'll see detailed logs including:
# - Template parsing steps
# - Cell processing information
# - Error messages with exact cell locations
# - Warning messages for potential issues
```

**Debug Output Features:**
- ✅ Colored console output for better readability
- ✅ Cell address tracking for errors (e.g., "Error in Cell A5")
- ✅ Line-by-line template debugging
- ✅ Warning messages for edge cases

**Log Levels:**
- `debug=False` (default): Only shows warnings and errors
- `debug=True`: Shows all debug, info, warning, and error messages

**Example Error Output:**
```
ERROR - Syntax Error in Cell B3
ERROR - error message: unexpected 'end of template'
ERROR - Cell B3 : {% for item in items %}
ERROR - Cell B4 : {{ item.name }}
```

### Basic Syntax

#### Variables

```jinja2
{{ name }}
{{ person.email }}
{{ items[0].price }}
```

#### Loops

In cell comments or inline:

```jinja2
{%- for item in items %}
{{ item.name }}
{%- endfor %}
```

#### Conditionals

```jinja2
{% if approved %}
✓ Approved
{% else %}
✗ Rejected
{% endif %}
```

#### Non-String Values

For numbers, dates, or formulas:

```jinja2
{% xv total_amount %}
{% xv sale_date %}
```

#### Images

```jinja2
{% img product_photo %}
```

#### Checkboxes

```jinja2
{% yn is_active %}
```

### Template Placement

You can place Jinja2 tags in:

1. **Cell values** - Direct replacement
2. **Cell comments** - Control flow (beforerow, beforecell, aftercell)
3. **Inline** (v0.9+) - Mix text and tags in cells

---

## 🎯 Advanced Usage

### Multiple Sheets

```python
data1 = {'name': 'Sheet1', 'items': [...]}
data2 = {'name': 'Sheet2', 'items': [...]}

writer.render_book([data1, data2])
```

### Custom Jinja2 Filters

```python
writer.add_filter('uppercase', str.upper)
```

Template:
```jinja2
{{ name | uppercase }}
```

### Custom Functions

```python
def format_currency(value):
    return f"${value:,.2f}"

writer.add_global('currency', format_currency)
```

Template:
```jinja2
{{ currency(price) }}
```

---

## 🆚 Comparison with Alternatives

| Feature | xlsxjinja | openpyxl | xlsxwriter | pandas |
|---------|-----------|----------|------------|--------|
| Template-based | ✅ | ❌ | ❌ | ❌ |
| Jinja2 syntax | ✅ | ❌ | ❌ | ❌ |
| Preserve formatting | ✅ | ✅ | ✅ | ❌ |
| Rich text | ✅ | ✅ | ✅ | ❌ |
| Images | ✅ | ✅ | ✅ | ❌ |
| Learning curve | Low | Medium | Medium | Medium |

---

## 🏗️ Architecture

xlsxjinja works by:

1. **Loading** the Excel template with openpyxl
2. **Parsing** cells for Jinja2 syntax
3. **Building** a tree structure of the template
4. **Compiling** to a Jinja2 template string
5. **Rendering** with your data
6. **Writing** back to Excel with formatting preserved

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

## 📜 License

MIT License

- Original work Copyright (c) 2020 Zhang Yu ([xltpl](https://github.com/zhangyu836/xltpl))
- Modified work Copyright (c) 2024 [Your Name] (xlsxjinja)

This project is a fork of xltpl with significant refactoring:
- Dropped Python 2 support
- Dropped .xls format support
- Modernized codebase for Python 3.7+
- Simplified dependencies (removed xlrd, xlwt, six)
- Enhanced documentation

---

## 🙏 Acknowledgments

This library is based on [xltpl](https://github.com/zhangyu836/xltpl) by Zhang Yu.

Major changes in this fork:
- Python 3.7+ only
- .xlsx format only (no .xls)
- Reduced dependencies from 5 to 2
- 40% less code through refactoring
- Modern Python practices

---

## 📞 Support

- 🐛 **Issues**: [GitHub Issues](https://github.com/yourusername/xlsxjinja/issues)
- 💡 **Discussions**: [GitHub Discussions](https://github.com/yourusername/xlsxjinja/discussions)
- 📧 **Email**: your.email@example.com

---

## 🔗 Related Projects

- [openpyxl](https://openpyxl.readthedocs.io/) - Python library to read/write Excel files
- [jinja2](https://jinja.palletsprojects.com/) - Modern templating engine
- [xltpl](https://github.com/zhangyu836/xltpl) - Original project (supports .xls and Python 2)

---

**Made with ❤️ for the Python community**
