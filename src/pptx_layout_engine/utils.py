import logging
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR

logger = logging.getLogger(__name__)

def parse_color(value, color_schemes) -> RGBColor:
    """Parses a color value into an RGBColor object.

    Args:
        value: Can be a list/tuple of RGB values, a hex string, a named color from a color scheme, or a string like '(255,255,255)'.
        color_schemes: Dictionary of named color scheme mappings.

    Returns:
        RGBColor instance, defaulting to black if parsing fails.
    """
    try:
        if isinstance(value, (list, tuple)) and len(value) == 3:
            return RGBColor(*value)

        if isinstance(value, str):
            if "." in value:
                scheme, key = value.split(".", 1)
                for item in color_schemes.get(scheme, []):
                    if key in item:
                        return parse_color(item[key], color_schemes)

            if value.startswith("#") and len(value) == 7:
                return RGBColor(
                    int(value[1:3], 16),
                    int(value[3:5], 16),
                    int(value[5:7], 16)
                )

            return RGBColor(*parse_rgb_string(value))
    except Exception as e:
        logger.warning(f"Failed to parse color value: {value} — {e}")

    return RGBColor(0, 0, 0)


def parse_rgb_string(s: str) -> tuple[int, int, int]:
    """Parses a string like '(255, 255, 255)' into a tuple."""
    try:
        return tuple(int(x.strip()) for x in s.strip("() ").split(",") if x.strip())
    except Exception as e:
        logger.warning(f"Failed to parse RGB string: {s} — {e}")
        return (0, 0, 0)


def get_alignment(alignment: str) -> PP_ALIGN:
    """Maps alignment string to PowerPoint horizontal alignment."""
    return {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY
    }.get(alignment.lower(), PP_ALIGN.LEFT)


def get_vertical_anchor(alignment: str) -> MSO_VERTICAL_ANCHOR:
    """Maps alignment string to PowerPoint vertical anchor."""
    return {
        "top": MSO_VERTICAL_ANCHOR.TOP,
        "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
        "bottom": MSO_VERTICAL_ANCHOR.BOTTOM
    }.get(alignment.lower(), MSO_VERTICAL_ANCHOR.TOP)
