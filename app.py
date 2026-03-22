import streamlit as st
from pptx import Presentation
import fitz  
from langchain_openai import ChatOpenAI
from langchain_community.tools.tavily_search import TavilySearchResults
import json
import os
import re

st.set_page_config(page_title="AI Presentation Creator", page_icon="📊")

def extract_text_from_pptx(file):
    try:
        prs = Presentation(file)
        return "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
    except Exception as e:
        return f"Error reading PPTX: {e}"

def extract_text_from_pdf(file):
    try:
        doc = fitz.open(stream=file.read(), filetype="pdf")
        return "\n".join([page.get_text() for page in doc])
    except Exception as e:
        return f"Error reading PDF: {e}"

def create_pptx(slides_data):
    prs = Presentation()
    for slide_info in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1]) # Title and Content layout
        slide.shapes.title.text = slide_info.get("title", "Slide")
        placeholder = slide.placeholders[1]
        placeholder.text = slide_info.get("content", "")
    
    output_path = "generated_presentation.pptx"
    prs.save(output_path)
    return output_path

# --- UI Layout ---
st.title("📊 AI Presentation Generator")

with st.sidebar:
    st.header("API Configuration")
    o_api = st.text_input("OpenAI API Key", type="password")
    t_api = st.text_input("Tavily API Key", type="password")
    num_slides = st.slider("Number of Slides", 5, 20, 10)
    st.info("Note: Ensure your OpenAI account has a paid balance ($5 min).")

topic = st.text_input("What is the presentation topic?", placeholder="e.g. Impact of AI on Healthcare")
ref_file = st.file_uploader("Optional: Upload reference (PDF/PPTX)", type=["pdf", "pptx"])

if st.button("Generate Presentation"):
    if not o_api or not t_api or not topic:
        st.warning("⚠️ Please provide both API keys and a topic.")
    else:
        try:
            with st.spinner("Step 1: Searching the web and reading references..."):
                os.environ["TAVILY_API_KEY"] = t_api
                
                # Extract text from reference
                ref_content = ""
                if ref_file:
                    if ref_file.name.endswith("pptx"):
                        ref_content = extract_text_from_pptx(ref_file)
                    else:
                        ref_content = extract_text_from_pdf(ref_file)
                
                # Search Internet
                search = TavilySearchResults(max_results=3)
                web_data = search.invoke(topic)

            with st.spinner("Step 2: AI is writing your slides..."):
                # Use gpt-4o-mini: Cheaper, faster, and higher rate limits
                llm = ChatOpenAI(model="gpt-4o-mini", api_key=o_api, temperature=0.7)
                
                prompt = f"""
                You are a professional presentation creator.
                Topic: {topic}
                Reference Materials: {ref_content[:2000]}
                Web Search Results: {web_data}
                
                Task: Create {num_slides} slides.
                Return ONLY a valid JSON array of objects. 
                Format: [{"title": "Slide Title", "content": "Bullet point 1\\nBullet point 2"}]
                """
                
                res = llm.invoke(prompt)
                
                # Clean JSON formatting (removes markdown backticks)
                clean_content = re.sub(r"```json|```", "", res.content).strip()
                data = json.loads(clean_content)
                
                path = create_pptx(data)
                
                st.success(f"✅ Created {len(data)} slides successfully!")
                with open(path, "rb") as f:
                    st.download_button(
                        label="📥 Download PowerPoint File",
                        data=f,
                        file_name=f"{topic.replace(' ', '_')}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

        except Exception as e:
            if "insufficient_quota" in str(e) or "RateLimitError" in str(e):
                st.error("❌ OpenAI Error: Your API key has no credits or has hit its limit. Please check your billing balance at platform.openai.com.")
            elif "invalid_api_key" in str(e):
                st.error("❌ Invalid API Key. Please check your keys.")
            else:
                st.error(f"⚠️ An error occurred: {e}")
