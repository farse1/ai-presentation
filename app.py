import streamlit as st
# ... (other imports)

# 1. Access Secrets securely
# It checks Streamlit Cloud settings first, then falls back to sidebar if empty
OPENAI_API_KEY = st.secrets.get("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")
TAVILY_API_KEY = st.secrets.get("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")

with st.sidebar:
    st.header("API Configuration")
    # If keys aren't in secrets, show the input boxes
    if not OPENAI_API_KEY:
        o_api = st.text_input("OpenAI API Key", type="password")
    else:
        o_api = OPENAI_API_KEY
        st.success("OpenAI Key loaded from Secrets")

    if not TAVILY_API_KEY:
        t_api = st.text_input("Tavily API Key", type="password")
    else:
        t_api = TAVILY_API_KEY
        st.success("Tavily Key loaded from Secrets")
    
    num_slides = st.slider("Number of Slides", 5, 20, 10)

# Use o_api and t_api in the rest of your script...
from pptx import Presentation
from fpdf import FPDF
import fitz  
from langchain_openai import ChatOpenAI
from langchain_community.tools.tavily_search import TavilySearchResults
import json
import os
import re

st.set_page_config(page_title="Presentation Generator", page_icon="📊")

# --- PPTX GENERATOR ---
def create_pptx(slides_data):
    prs = Presentation()
    for slide_info in slides_data:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_info.get("title", "Presentation Slide")
        body_shape = slide.placeholders[1]
        body_shape.text = slide_info.get("content", "")
    
    path = "generated_presentation.pptx"
    prs.save(path)
    return path

# --- PDF GENERATOR ---
def create_pdf(slides_data):
    pdf = FPDF(orientation="landscape", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    
    for slide_info in slides_data:
        pdf.add_page()
        
        # Title Background
        pdf.set_fill_color(200, 220, 255)
        pdf.rect(0, 0, 297, 30, 'F')
        
        # Title Text
        pdf.set_font("Arial", 'B', 24)
        pdf.set_xy(10, 10)
        pdf.cell(0, 10, slide_info.get("title", "Slide"), ln=True)
        
        # Content Text
        pdf.set_font("Arial", size=14)
        pdf.set_xy(10, 40)
        # multi_cell handles line breaks (\n) automatically
        pdf.multi_cell(0, 10, slide_info.get("content", ""))
        
    path = "generated_presentation.pdf"
    pdf.output(path)
    return path

# --- REFERENCE EXTRACTION ---
def extract_text(file):
    if file.name.endswith("pptx"):
        prs = Presentation(file)
        return "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
    else:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        return "\n".join([page.get_text() for page in doc])

# --- UI LAYOUT ---
st.title("📊 AI Presentation & PDF Generator")

with st.sidebar:
    st.header("API Keys")
    o_api = st.text_input("OpenAI API Key", type="password")
    t_api = st.text_input("Tavily API Key", type="password")
    num_slides = st.slider("Number of Slides", 5, 20, 10)
    st.info("Check OpenAI balance at platform.openai.com")

topic = st.text_input("Enter Topic:", placeholder="e.g. Modern Architecture")
ref_file = st.file_uploader("Upload reference (Optional)", type=["pdf", "pptx"])

if st.button("Generate Files"):
    if not o_api or not t_api or not topic:
        st.warning("Please enter keys and topic.")
    else:
        try:
            with st.spinner("Searching and generating content..."):
                os.environ["TAVILY_API_KEY"] = t_api
                
                # Context gathering
                ref_text = extract_text(ref_file)[:1500] if ref_file else ""
                search = TavilySearchResults(max_results=3)
                web_data = search.invoke(topic)
                
                # AI Logic
                llm = ChatOpenAI(model="gpt-4o-mini", api_key=o_api)
                prompt = f"""
                Create a presentation on: {topic}
                Reference: {ref_text}
                Web: {web_data}
                Slides: {num_slides}
                Return ONLY a JSON array: [{{ "title": "...", "content": "..." }}]
                """
                
                res = llm.invoke(prompt)
                clean_json = re.sub(r"```json|```", "", res.content).strip()
                slides_json = json.loads(clean_json)
                
                # File creation
                pptx_file = create_pptx(slides_json)
                pdf_file = create_pdf(slides_json)
                
                st.success(f"✅ Generated {len(slides_json)} slides!")
                
                col1, col2 = st.columns(2)
                with col1:
                    with open(pptx_file, "rb") as f:
                        st.download_button("📥 Download PPTX", f, file_name=f"{topic}.pptx")
                with col2:
                    with open(pdf_file, "rb") as f:
                        st.download_button("📥 Download PDF", f, file_name=f"{topic}.pdf")
                        
        except Exception as e:
            st.error(f"Error: {str(e)}")
