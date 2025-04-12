import yaml
import os
from importlib.resources import files
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

from .utils import (
    parse_color,
    parse_rgb_string,
    get_alignment,
    get_vertical_anchor
)

class PresentationBuilder:
    def __init__(self, template_path):
        self.prs = Presentation(template_path)
        self._clear_all_slides()
        self.template_defaults = self._load_layout_defaults()
        self.shape_defaults = self._load_shape_defaults("resources/shape_layouts/shape_defaults.yml")

    def _load_shape_defaults(self, resource_path="resources/shape_layouts/shape_defaults.yml"):
        try:
            full_path = files("pptx_layout_engine").joinpath(resource_path)
            with full_path.open("r", encoding="utf-8") as f:
                return yaml.safe_load(f)
        except Exception:
            return {}

    def _clear_all_slides(self):
        while self.prs.slides:
            self.prs.slides._sldIdLst.remove(self.prs.slides._sldIdLst[0])

    def _load_layout_defaults(self, layout_path=None):
        if not layout_path:
            return {
                "font": "Avenir LT  Next Pro",
                "font_size": 18,
                "left": 0.67,
                "top": 0.4,
                "horizontal_align": "left",
                "vertical_align": "top"
            }
        with open(layout_path, 'r') as f:
            data = yaml.safe_load(f)
            return data.get("defaults", {})
    
    def _resolve_template_layout(self, template_layout):
        if isinstance(template_layout, int):
            return self.prs.slide_layouts[template_layout]
        elif isinstance(template_layout, str):
            for layout in self.prs.slide_layouts:
                if layout.name.strip().lower() == template_layout.strip().lower():
                    return layout
        return self.prs.slide_layouts[6]  # fallback to Blank

    def _fill_placeholder(self, slide, placeholder_name, text, shape_cfg, defaults, color_schemes):
        idx = int(placeholder_name.replace("placeholder", ""))
        if idx >= len(slide.placeholders):
            print(f"⚠️ No placeholder {idx} in this layout.")
            return

        ph = slide.placeholders[idx]
        frame = ph.text_frame
        frame.clear()
        frame.word_wrap = True

        # Only override vertical alignment if specified
        if "vertical_align" in shape_cfg:
            frame.vertical_anchor = get_vertical_anchor(shape_cfg["vertical_align"])

        def apply_run_style(run):
            if "font" in shape_cfg:
                run.font.name = shape_cfg["font"]
            if "font_size" in shape_cfg:
                run.font.size = Pt(shape_cfg["font_size"])
            if shape_cfg.get("font_weight", "").lower() == "bold":
                run.font.bold = True
            if shape_cfg.get("font_italic", False):
                run.font.italic = True
            if "font_color" in shape_cfg:
                run.font.color.rgb = parse_color(shape_cfg["font_color"], color_schemes)

        if isinstance(text, list):
            for i, line in enumerate(text):
                para = frame.add_paragraph() if i > 0 else frame.paragraphs[0]
                para.text = line
                para.level = 0

                if "horizontal_align" in shape_cfg:
                    para.alignment = get_alignment(shape_cfg["horizontal_align"])

                run = para.runs[0]
                apply_run_style(run)
        else:
            para = frame.paragraphs[0]
            para.text = text

            if "horizontal_align" in shape_cfg:
                para.alignment = get_alignment(shape_cfg["horizontal_align"])

            run = para.runs[0]
            apply_run_style(run)

    def _add_text_shape(self, slide, text, shape_cfg, defaults, color_schemes):
        left = Inches(shape_cfg.get("left", defaults["left"]))
        top = Inches(shape_cfg.get("top", defaults["top"]))
        width = Inches(shape_cfg.get("width", 5))
        height = Inches(shape_cfg.get("height", 1))

        textbox = slide.shapes.add_textbox(left, top, width, height)
        frame = textbox.text_frame
        frame.clear()
        frame.word_wrap = True

        frame.vertical_anchor = get_vertical_anchor(
            shape_cfg.get("vertical_align", defaults["vertical_align"])
        )

        # Text can be a string or a list of bullet points
        if isinstance(text, list):
            for i, line in enumerate(text):
                para = frame.add_paragraph() if i > 0 else frame.paragraphs[0]
                para.text = f"• {line}" if shape_cfg.get("bullet", False) else line
                para.alignment = get_alignment(shape_cfg.get("horizontal_align", defaults["horizontal_align"]))
                para.level = 0  # bullet level

                run = para.runs[0]
                run.font.name = shape_cfg.get("font", defaults["font"])
                run.font.size = Pt(shape_cfg.get("font_size", defaults["font_size"]))
                if shape_cfg.get("font_weight", "").lower() == "bold":
                    run.font.bold = True
                if shape_cfg.get("font_italic", False):
                    run.font.italic = True
                if "font_color" in shape_cfg:
                    run.font.color.rgb = parse_color(shape_cfg["font_color"], color_schemes)

        else:
            para = frame.paragraphs[0]
            para.text = text
            para.alignment = get_alignment(shape_cfg.get("horizontal_align", defaults["horizontal_align"]))

            run = para.runs[0]
            run.font.name = shape_cfg.get("font", defaults["font"])
            run.font.size = Pt(shape_cfg.get("font_size", defaults["font_size"]))
            if shape_cfg.get("font_weight", "").lower() == "bold":
                run.font.bold = True
            if shape_cfg.get("font_italic", False):
                run.font.italic = True
            if "font_color" in shape_cfg:
                run.font.color.rgb = parse_color(shape_cfg["font_color"], color_schemes)

    def _add_table_shape(self, slide, table_data, shape_cfg, defaults, color_schemes):
        rows = len(table_data)
        cols = len(table_data[0]) if rows > 0 else 1

        left = Inches(shape_cfg.get("left", defaults.get("left", 1.0)))
        top = Inches(shape_cfg.get("top", defaults.get("top", 1.0)))
        width = Inches(shape_cfg.get("width", 5.5))
        height = Inches(shape_cfg.get("height", 3.0))

        col_widths = shape_cfg.get("column_widths", None)
        row_height = shape_cfg.get("row_height", None)

        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table

        font_name = shape_cfg.get("font", defaults["font"])
        font_size = Pt(shape_cfg.get("font_size", defaults["font_size"]))
        header_bold = shape_cfg.get("header_bold", True)
        cell_padding = Inches(shape_cfg.get("cell_padding", 0.05))

        horiz_align = get_alignment(shape_cfg.get("horizontal_align", defaults["horizontal_align"]))
        vert_align = get_vertical_anchor(shape_cfg.get("vertical_align", defaults["vertical_align"]))

        # Row height
        if row_height:
            for i in range(rows):
                table.rows[i].height = Inches(row_height)

        # Column widths
        if col_widths and len(col_widths) == cols:
            for i, w in enumerate(col_widths):
                table.columns[i].width = Inches(w)

        # Color scheme
        header_fill = None
        row_fills = []
        if "table_colors" in shape_cfg:
            scheme_name = shape_cfg["table_colors"]
            scheme = color_schemes.get(scheme_name, [])
            if isinstance(scheme, list):
                for color_dict in scheme:
                    if "header_green" in color_dict:
                        header_fill = parse_color(color_dict["header_green"], color_schemes)
                    elif "row_light" in color_dict:
                        row_fills.append(parse_color(color_dict["row_light"], color_schemes))
                    elif "row_dark" in color_dict:
                        row_fills.append(parse_color(color_dict["row_dark"], color_schemes))

        # Fill cells
        for r, row_data in enumerate(table_data):
            for c, cell_text in enumerate(row_data):
                cell = table.cell(r, c)
                cell.text = str(cell_text)
                cell.text_frame.word_wrap = True
                cell.text_frame.vertical_anchor = vert_align
                cell.margin_left = cell.margin_right = cell.margin_top = cell.margin_bottom = cell_padding

                p = cell.text_frame.paragraphs[0]
                p.alignment = horiz_align
                run = p.runs[0]
                run.font.name = font_name
                run.font.size = font_size

                if r == 0:
                    if header_bold:
                        run.font.bold = True
                    if header_fill:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = header_fill
                else:
                    if row_fills:
                        fill_color = row_fills[(r - 1) % len(row_fills)]
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = fill_color

    def _add_image_shape(self, slide, image_path, shape_cfg):
        import os
        if not image_path or not os.path.exists(image_path):
            print(f"Warning: missing image at path: {image_path}")
            return

        left = Inches(shape_cfg.get("left", 1.0))
        top = Inches(shape_cfg.get("top", 1.0))
        width = Inches(shape_cfg.get("width", 4.0))
        height = Inches(shape_cfg.get("height", 3.0))

        slide.shapes.add_picture(image_path, left, top, width, height)

    def load_presentation(self, deck_yaml_path):
        with open(deck_yaml_path, "r") as f:
            deck = yaml.safe_load(f)

        # If the YAML file defines a template, replace current presentation
        template_path = deck.get("template", None)
        if template_path:
            from pptx import Presentation
            self.prs = Presentation(template_path)
            self._clear_all_slides()

        # Handle global defaults
        defaults = deck.get("defaults", {})
        layout_root = defaults.get("slide_layout_path", "")
        shape_root = defaults.get("shape_layout_path", "")
        self.template_defaults.update(defaults)
        self.template_defaults.update(deck.get("defaults", {}))

        slides = deck.get("slides", [])
        for idx, slide_def in enumerate(slides):
            layout_file = slide_def.get("layout")
            content = slide_def.get("content", {})

            if not layout_file:
                print(f"[Slide {idx}] Missing layout file. Skipping.")
                continue

            full_layout_path = os.path.join(layout_root, layout_file)
            if not os.path.exists(full_layout_path):
                print(f"[Slide {idx}] Layout not found: {full_layout_path}. Skipping.")
                continue

            self.add_slide(layout=full_layout_path, content=content)

    def _add_placeholder(self, slide, name, value, shape_cfg, defaults, color_schemes):
        try:
            idx = int(name.replace("placeholder", ""))
            placeholder = slide.placeholders[idx]
        except (ValueError, IndexError):
            print(f"⚠️ Could not find placeholder '{name}'")    
            return

        frame = placeholder.text_frame
        frame.clear()
        frame.word_wrap = True
        frame.vertical_anchor = get_vertical_anchor(shape_cfg.get("vertical_align", defaults.get("vertical_align", "top")))

        if isinstance(value, list):
            for i, line in enumerate(value):
                para = frame.add_paragraph() if i > 0 else frame.paragraphs[0]
                para.text = f"• {line}" if shape_cfg.get("bullet", False) else line
                para.level = 0
                para.alignment = get_alignment(shape_cfg.get("horizontal_align", defaults.get("horizontal_align", "left")))

                run = para.runs[0]
                run.font.name = shape_cfg.get("font", defaults.get("font", "Arial"))
                run.font.size = Pt(shape_cfg.get("font_size", defaults.get("font_size", 18)))
        else:
            para = frame.paragraphs[0]
            para.text = value
            para.alignment = get_alignment(shape_cfg.get("horizontal_align", defaults.get("horizontal_align", "left")))

            run = para.runs[0]
            run.font.name = shape_cfg.get("font", defaults.get("font", "Arial"))
            run.font.size = Pt(shape_cfg.get("font_size", defaults.get("font_size", 18)))

    def add_slide(self, layout, content):
        with open(layout, 'r') as f:
            layout_cfg = yaml.safe_load(f)

        template_layout = layout_cfg.get("template_layout")
        slide_layout = self._resolve_template_layout(template_layout)
        slide = self.prs.slides.add_slide(slide_layout)
        layout_defaults = {**self.template_defaults, **layout_cfg.get("defaults", {})}
        color_schemes = layout_cfg.get("color_schemes", {})
        shapes = layout_cfg.get("shapes", {})

        for shape_name, value in content.items():
            layout_shape = shapes.get(shape_name, {})
            default_shape = self.shape_defaults.get(shape_name, {})
            shape_cfg = {**default_shape, **layout_shape}

            if shape_name.startswith("placeholder"):
                self._fill_placeholder(slide, shape_name, value, shape_cfg, layout_defaults, color_schemes)
                continue

            shape_type = shape_cfg.get("type", "text")

            if shape_type == "text":
                self._add_text_shape(slide, value, shape_cfg, layout_defaults, color_schemes)
            elif shape_type == "table":
                self._add_table_shape(slide, value, shape_cfg, layout_defaults, color_schemes)
            elif shape_type == "image":
                self._add_image_shape(slide, value, shape_cfg)

    def save(self, path):
        self.prs.save(path)
