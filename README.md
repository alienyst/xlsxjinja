# xlsxjinja

**Generate Excel (.xlsx) files from templates using Jinja2**

[![PyPI](https://img.shields.io/pypi/v/xlsxjinja?style=flat&color=2222FF)](https://pypi.org/project/xlsxjinja/)
[![License](https://img.shields.io/github/license/alienyst/xlsxjinja?style=flat&color=222222)](https://github.com/alienyst/xlsxjinja/blob/main/LICENSE)
[![Python](https://img.shields.io/badge/python-3.7+-222222?style=flat)](https://www.python.org)

Fork of [xltpl](https://github.com/zhangyu836/xltpl) by [@zhangyu836](https://github.com/zhangyu836)

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
- ⚡ **Fast & Lightweight** - Only 2 required dependencies (openpyxl + jinja2)

---

## 📦 Installation

```bash
# Core install: text, formulas, loops, conditionals - no image support
pip install xlsxjinja

# With image support ({% img %} and {% insert_img %} tags)
pip install xlsxjinja[image]
```

### Requirements

- Python 3.7+
- openpyxl >= 3.1.0
- jinja2 >= 2.10
- Pillow >= 9.0 (optional, only required for `{% img %}` / `{% insert_img %}`)

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

#### Non-String Values (`{% xv %}`)

By default all cell values are rendered as strings. Use `{% xv %}` to preserve
the original Python type (number, date, boolean) so Excel treats it correctly:

```jinja2
{% xv total_amount %}
{% xv sale_date %}
{% xv item.qty, 1 %}
```

The optional second argument is an index for multiple `xv` values in one cell.

#### Checkboxes (`{% yn %}`)

Render a boolean as a checkbox-style value:

```jinja2
{% yn is_active %}
```

---

## 🔁 Looping

xlsxjinja supports two ways to loop over rows.

### Method 1 — `{% tr %}` tag (recommended for simple rows)

Place `{% tr %}` anywhere inside a cell value. The engine strips the tag and
promotes surrounding `for`/`if` tags to row-level control flow automatically.

**Rules:**
- Opening tags (`{% for %}`, `{% if %}`, `{% elif %}`, `{% else %}`) at the
  **start** of the cell value → hoisted to `beforerow` of that row.
- Closing tags (`{% endfor %}`, `{% endif %}`) at the **end** of the cell
  value → placed as `aftercell` of that cell.
- Put `{% tr %}` **before** the opening tag and **after** the closing tag.

**Template layout:**

| A | B | C |
|---|---|---|
| Name | Qty | Price |
| `{% tr %}{% for line in lines %}{{ line.name }}` | `{{ line.qty }}` | `{{ line.price }}{% tr %}{% endfor %}` |

**Example — basic loop:**

```python
import openpyxl
from xlsxjinja import BookWriter

wb = openpyxl.Workbook()
ws = wb.active
ws['A1'], ws['B1'] = 'Name', 'Price'
ws['A2'] = '{% tr %}{% for item in items %}{{ item["name"] }}'
ws['B2'] = '{{ item["price"] }}{% tr %}{% endfor %}'
wb.save('template.xlsx')

writer = BookWriter('template.xlsx')
writer.render_book([{
    'tpl_name': ws.title,
    'sheet_name': 'Report',
    'items': [{'name': 'Laptop', 'price': 999}, {'name': 'Mouse', 'price': 25}],
}])
writer.save('output.xlsx')
```

**Example — conditional row with `{% tr %}`:**

| A | B |
|---|---|
| `{% tr %}{% if show_total %}TOTAL` | `{{ total }}{% tr %}{% endif %}` |

### Method 2 — Cell comment `beforerow` (required for merged-cell rows)

Place Jinja2 control tags inside a **cell comment** using the key `beforerow`.
This is the only supported method when:

- The template row contains **merged cells**, or
- A single `{% for %}` block needs to span **multiple template rows**
  (e.g. `display_type` patterns with section rows and product rows).

**Comment format (inside the cell comment):**
```
beforerow:{% for line in lines %}
```

Add the comment to the **first cell** of the row where the loop starts.
Close the loop at the end of the last row's cell value.

**Example — `display_type` (section row merged A:C, product rows normal):**

```
Row 1  — Header:  Description | Qty | Price
Row 2  — A2:C2 merged, comment on A2:
           beforerow:{% for line in lines %}{% if line.display_type == 'line_section' %}
         Cell A2 value:
           {{ line.name }}{% endif %}
Row 3  — comment on A3:
           beforerow:{% if not line.display_type %}
         Cell A3: {{ line.name }}
         Cell B3: {{ line.qty }}
         Cell C3: {{ line.price }}{% endif %}{% endfor %}
```

> **Why not `{% tr %}` here?**
>
> `{% tr %}` promotes opening tags to the `beforerow` of **its own row only**.
> A `{% for %}` that must open on row 2 and close on row 3 cannot be expressed
> with `{% tr %}` — the tag boundary is one row. Use `beforerow` comments instead.

### Comparison

| Scenario | `{% tr %}` | `beforerow` comment |
|---|---|---|
| Simple row loop | ✅ | ✅ |
| Conditional row | ✅ | ✅ |
| Merged cell row | ❌ | ✅ |
| `for` spanning multiple template rows | ❌ | ✅ |
| No Excel comment needed | ✅ | ❌ |

---

## 🔀 Conditionals

```jinja2
{% if approved %}
✓ Approved
{% else %}
✗ Rejected
{% endif %}
```

Can also be used inline with `{% tr %}` (see Looping above).

---

## 🖼️ Images

### Replace a placeholder image (`{% img %}`)

Insert a dummy image in the template cell. At render time the dummy is replaced:

```jinja2
{% img product_photo %}
{% img product_photo, 1 %}
```

Requires `pip install xlsxjinja[image]`.

### Insert a new image (`{% insert_img %}`)

Insert an image into a cell that has **no placeholder** in the template.
The image is automatically sized to fit the cell's column width and row height:

```jinja2
{% insert_img company_logo %}
```

`company_logo` must be a PIL `Image` object, a `BytesIO`, or a base64-encoded
`bytes` value. WebP images are automatically converted to PNG in-memory.

---

### Template Placement Summary

| Location | Use for |
|---|---|
| **Cell value** | Variables, inline loops/conditionals via `{% tr %}` / `{% tc %}` |
| **Cell comment** `beforerow` | Loop/conditional opening tags, especially for merged rows |
| **Cell comment** `aftercell` | Loop/conditional closing tags (rarely needed manually) |


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

## 🙏 Acknowledgments

Originally built on [xltpl](https://github.com/zhangyu836/xltpl) by [Zhang Yu](https://github.com/zhangyu836). Major changes in this fork:

- Python 3.7+ only, dropped Python 2 support
- .xlsx format only, dropped .xls support (xlrd/xlwt removed)
- Reduced dependencies from 5 to 2 required (Pillow now optional, via `xlsxjinja[image]`)
- ~40% less code through refactoring and unified class implementations
- Native `{% tr %}` / `{% tc %}` inline control flow, `{% insert_img %}`, and WebP support added on top of upstream xltpl
