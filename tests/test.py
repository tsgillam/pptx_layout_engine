#%%
import sys
import os

ROOT = os.path.dirname(os.path.dirname(__file__))  # project root
SRC_PATH = os.path.join(ROOT, "src")
sys.path.insert(0, SRC_PATH)
from pptx_layout_engine.builder import PresentationBuilder

SLIDE_LAYOUTS_DIR = os.path.join(ROOT, "src", "pptx_layout_engine", "resources", "slide_layouts")
TEMPLATE_PATH = os.path.join(ROOT, "src", "pptx_layout_engine", "resources", "powerpoint_templates", "template.pptx")
TEST = os.path.dirname(__file__)
OUTPUT_PATH = os.path.join(TEST, "test_output.pptx")

builder = PresentationBuilder(template_path=TEMPLATE_PATH)
#%%
builder.load_presentation("slide_deck.yml")
builder.save("test_output.pptx")
# %%
