#%%
from pptx import Presentation

prs = Presentation("test_output.pptx")
layout = prs.slide_layouts[1]  # usually title slide

print(f"Layout:{layout.name}")
for shape in layout.placeholders:
    print(f"Index: {shape.placeholder_format.idx}, Type: {shape.placeholder_format.type}")

# %%
