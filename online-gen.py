import streamlit as st
from generate import parse_text, create_doc
from io import BytesIO, StringIO

st.title("Text to Docx Menu Generator")

uploaded = st.file_uploader(
    "Upload text file",
    type=["txt"]
)

if uploaded:
    text = uploaded.read().decode("utf-8")
        
    items = parse_text(StringIO(text))
    output_doc = create_doc('\\is_template.docx', items)
    print(f"\nSuccessfully generated {output_file_name}\n")
    
    output = BytesIO()
    output_doc.save(output)
    output.seek(0)
    
    st.download_button(
        "Download Menu",
        output,
        file_name="Menu-Iron_Skillet.docx",
        mime = "application/vnd.openxmlformats-officedocuments.wordprocessingxml.document"
    )
