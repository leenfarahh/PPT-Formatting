"""
Generates a small synthetic "rough content" .pptx - a few slides with
inconsistent, off-brand formatting, standing in for what a designer's
first draft looks like before Stage 3 cleans it up.

Run: python examples/make_sample_content.py
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

OUT_PATH = Path(__file__).resolve().parent / "sample_content.pptx"


def main():
    prs = Presentation()
    blank_layout = prs.slide_masters[0].slide_layouts[6]  # "Blank"

    # Slide 1: off-brand title + body text, slightly off-grid.
    slide1 = prs.slides.add_slide(blank_layout)
    title_box = slide1.shapes.add_textbox(Emu(731520), Emu(365760), Inches(8), Inches(1))
    title_box.text_frame.text = "Market Overview"
    title_box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    title_box.text_frame.paragraphs[0].runs[0].font.name = "Arial"

    body_box = slide1.shapes.add_textbox(Emu(800100), Inches(1.8), Inches(7.5), Inches(2))
    body_box.text_frame.text = "Placeholder body copy describing the market landscape and key trends."
    body_box.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
    body_box.text_frame.paragraphs[0].runs[0].font.name = "Times New Roman"

    # An off-brand colored rectangle (not one of the theme's accent colors) -
    # this is exactly what apply_color_mapping() should remap.
    rect = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4), Inches(2), Inches(1))
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0x33, 0x99, 0xFF)  # arbitrary off-theme blue
    rect.text_frame.text = "Highlight"

    # Slide 2: a simple table, off-brand fonts, to exercise format_table().
    slide2 = prs.slides.add_slide(blank_layout)
    title_box2 = slide2.shapes.add_textbox(Emu(731520), Emu(365760), Inches(8), Inches(1))
    title_box2.text_frame.text = "Key Metrics"
    title_box2.text_frame.paragraphs[0].runs[0].font.size = Pt(30)

    rows, cols = 3, 3
    table_shape = slide2.shapes.add_table(rows, cols, Inches(0.7), Inches(1.8), Inches(8), Inches(2))
    table = table_shape.table
    headers = ["Metric", "This Quarter", "Last Quarter"]
    data = [["Revenue", "$4.2M", "$3.8M"], ["Margin", "22%", "19%"]]
    for c, h in enumerate(headers):
        table.cell(0, c).text = h
    for r, row in enumerate(data, start=1):
        for c, val in enumerate(row):
            table.cell(r, c).text = val

    prs.save(str(OUT_PATH))
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
