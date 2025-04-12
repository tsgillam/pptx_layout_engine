# pptx_layout_engine

Generate PowerPoint presentations from YAML files using custom `.yml` slide layout templates and placeholder-based content injection. Built on top of `python-pptx`.

## ⚙️ Features

- Supports `.yml` layout files with shape and styling config
- Uses standard PowerPoint slide placeholders (`placeholder0`, `placeholder1`, etc.)
- Add text, tables, and images dynamically
- Compatible with design templates (`.pptx`)
- Custom shape defaults and color schemes

---

## 🚀 Quick Start

### 📦 Installation

Install via GitHub:

```bash
pip install git+https://github.com/tsgillam/pptx-layout-engine.git
```

Or for development:

```bash
git clone https://github.com/yourusername/pptx-layout-engine.git
cd pptx-layout-engine
pip install -e .
```

---

## 📄 Example Usage

```python
from pptx_layout_engine.builder import PresentationBuilder

builder = PresentationBuilder(template_path="resources/powerpoint_templates/template.pptx")

builder.load_presentation("examples/example_deck.yml")

builder.save("output.pptx")
```

---

## 📝 Slide Deck YAML Format

```yaml
template: resources/powerpoint_templates/template.pptx
defaults:
  slide_layout_path: resources/slide_layouts/
  shape_layout_path: resources/shape_layouts/

slides:
  - layout: title_slide.yml
    content:
      placeholder0: "Welcome to Kansas"
      placeholder1: ["Explore a mix of historical and cultural attractions."]
```

---

## 🧱 Project Structure

```
pptx_layout_engine/
├── builder.py             # Core class for slide generation
├── utils.py               # Color parsing, alignment, etc.
├── resources/
│   ├── slide_layouts/     # .yml files defining each layout
│   ├── shape_layouts/     # shape_defaults.yml and overrides
│   └── powerpoint_templates/  # base .pptx templates
├── tests/                 # Unit tests
└── examples/              # Optional: example decks or test slides
```

---

## 🧪 Testing

```bash
pytest tests/
```

---

## 📄 License

MIT

---

## 👤 Author

Tom Gillam – [github.com/tsgillam](https://github.com/tsgillam)
```

---
