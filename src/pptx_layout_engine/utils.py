from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR

def parse_color(value, color_schemes):
    if isinstance(value, (list, tuple)) and len(value) == 3:
        return RGBColor(*value)

    if isinstance(value, str):
        if "." in value:
            scheme, key = value.split(".", 1)
            for item in color_schemes.get(scheme, []):
                if key in item:
                    return parse_color(item[key], color_schemes)

        if value.startswith("#") and len(value) == 7:
            try:
                return RGBColor(int(value[1:3], 16), int(value[3:5], 16), int(value[5:7], 16))
            except:
                pass

        return RGBColor(*parse_rgb_string(value))

    return RGBColor(0, 0, 0)

def parse_rgb_string(s):
    try:
        return tuple(int(x.strip()) for x in s.strip("()").split(","))
    except:
        return (0, 0, 0)

def get_alignment(alignment):
    return {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT
    }.get(alignment.lower(), PP_ALIGN.LEFT)

def get_vertical_anchor(alignment):
    return {
        "top": MSO_VERTICAL_ANCHOR.TOP,
        "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
        "bottom": MSO_VERTICAL_ANCHOR.BOTTOM
    }.get(alignment.lower(), MSO_VERTICAL_ANCHOR.TOP)
