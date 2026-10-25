import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# 1. Page & Tab Configuration
st.set_page_config(page_title="MLT Research Formatter", page_icon="📄", layout="centered")

# 2. Sidebar Branding (Suneel Yadav)
st.sidebar.title("👨‍🔬 About the Developer")
st.sidebar.markdown("Developed by **Suneel Yadav**")
st.sidebar.markdown("🔬 **[Learn MLT Online (YouTube)](#)**") # Aap yahan apne channel ka link daal sakte hain
st.sidebar.markdown("---")
st.sidebar.info("Yeh tool MLT professionals aur students ko research papers, syllabi, aur scripts instantly format karne mein madad karta hai.")

# 3. Main App UI
st.title("📄 MLT Research Formatter (Pro)")
st.markdown("Apna AI text paste karein aur 1-click mein professional, perfectly aligned Word file banayein.")

# 4. Settings (Columns)
col1, col2 = st.columns(2)
with col1:
    font_name = st.selectbox("Font Type", ["Times New Roman", "Arial", "Calibri", "Garamond"])
    font_size = st.number_input("Font Size", min_value=8, max_value=24, value=12)
with col2:
    line_spacing = st.selectbox("Line Spacing", [1.0, 1.15, 1.5, 2.0], index=2)
    alignment = st.selectbox("Paragraph Alignment", ["Left", "Center", "Right", "Justify"], index=3)

align_dict = {
    "Left": WD_ALIGN_PARAGRAPH.LEFT,
    "Center": WD_ALIGN_PARAGRAPH.CENTER,
    "Right": WD_ALIGN_PARAGRAPH.RIGHT,
    "Justify": WD_ALIGN_PARAGRAPH.JUSTIFY
}

# 5. Inputs
heading_text = st.text_input("Paper Title / Main Heading (Optional):")
body_text = st.text_area("Apna Research Text Yahan Paste Karein:", height=250)

# 6. Generate Logic with Error Handling & Animation
if st.button("Generate Formatted Word File", type="primary"):
    if not body_text.strip():
        st.warning("⚠️ Please enter your research text first!")
    else:
        doc = Document()

        # Heading Logic
        if heading_text.strip():
            heading = doc.add_heading(heading_text, level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in heading.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size + 4)
                run.bold = True
                run.font.color.rgb = None  # Default black color

        # Body Text Logic
        paragraphs = body_text.split('\n')
        for para_text in paragraphs:
            if para_text.strip():
                p = doc.add_paragraph()
                p.alignment = align_dict[alignment]
                p.paragraph_format.line_spacing = line_spacing
                run = p.add_run(para_text.strip())
                run.font.name = font_name
                run.font.size = Pt(font_size)

        # File Processing
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)

        # Success Animation
        st.balloons()
        st.success("✅ Aapki Word file successfully ban gayi hai! Neeche click karke download karein.")

        st.download_button(
            label="⬇️ Download Word Document",
            data=bio,
            file_name="MLT_Formatted_Paper.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
