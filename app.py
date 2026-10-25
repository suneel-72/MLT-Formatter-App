import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
import io

# App ka Header
st.title("📄 MLT Research Formatter")
st.write("Apna AI text paste karein aur 1-click mein formatted Word file banayein.")

# Professional Dropdowns
col1, col2, col3 = st.columns(3)
with col1:
    font_name = st.selectbox("Font Type", ["Times New Roman", "Arial", "Calibri"])
with col2:
    font_size = st.number_input("Font Size", min_value=10, max_value=14, value=12)
with col3:
    line_spacing = st.selectbox("Spacing", [1.0, 1.5, 2.0], index=1)

# Text Box
client_text = st.text_area("Apna Research Text Yahan Paste Karein:", height=250)

# Format & Download Button
if st.button("✨ Format & Generate Word File"):
    if client_text.strip():
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = font_name
        style.font.size = Pt(font_size)

        for line in client_text.split('\n'):
            line = line.strip()
            if line:
                p = doc.add_paragraph(line)
                
                # Spacing logic
                if line_spacing == 1.0:
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                elif line_spacing == 2.0:
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
                else:
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                
                # Auto-Bold Headings
                headings = ['Abstract', 'Background', 'Methods', 'Results', 'Conclusion']
                if any(line.startswith(h) for h in headings):
                    for run in p.runs:
                        run.bold = True
        
        # Download ke liye file save karna
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("✅ File Ready!")
        st.download_button(
            label="⬇️ Download Document",
            data=bio.getvalue(),
            file_name="Formatted_Research_Paper.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("Pehele text paste karein!")
