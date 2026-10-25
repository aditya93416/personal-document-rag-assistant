import os
import streamlit as st
from config import RAGConfig
from ingestion import IngestionPipeline
from query_engine import RAGQueryEngine

st.set_page_config(
    page_title="Personal Document Q&A Assistant (RAG Pipeline)",
    page_icon="📚",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-title { font-size: 2.1rem; color: #102C57; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-title { font-size: 1.05rem; color: #64748B; margin-bottom: 1.5rem; }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 12px; }
    .citation-box { background-color: #EEF2F6; border-left: 4px solid #102C57; padding: 8px 12px; margin: 6px 0; border-radius: 4px; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

@st.cache_resource
def load_rag_engine():
    pipeline = IngestionPipeline()
    engine = RAGQueryEngine(pipeline)
    return pipeline, engine

pipeline, engine = load_rag_engine()

# Sidebar: Controls & Ingestion
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/artificial-intelligence.png", width=64)
    st.title("RAG Controls")
    st.markdown("**Intern:** Aditya Kumar (024101010012)")
    st.markdown("**Organization:** InternCourse")
    st.divider()
    
    st.subheader("1. Ingest Documents")
    uploaded_files = st.file_uploader(
        "Upload PDF, TXT or Markdown files",
        type=["pdf", "txt", "md"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("Process & Index Documents", type="primary"):
            with st.spinner("Extracting, Chunking & Embedding..."):
                for uploaded_file in uploaded_files:
                    save_path = os.path.join(pipeline.config.DOCS_DIR, uploaded_file.name)
                    with open(save_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    stats = pipeline.ingest_file(save_path)
                    st.success(f"Indexed: {stats['filename']} ({stats['total_chunks']} chunks)")

    st.divider()
    st.subheader("2. Hyperparameters")
    top_k = st.slider("Top-K Retrievals", min_value=1, max_value=8, value=4)
    st.caption("Embedding: `sentence-transformers/all-MiniLM-L6-v2` (384-dim)")
    
    stats = pipeline.get_collection_stats()
    st.metric("Total Indexed Chunks", stats["total_vectors"])

# Main Chat View
st.markdown('<div class="main-title">Personal Document Q&A Assistant</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Privacy-Preserving Retrieval-Augmented Generation (RAG Pipeline)</div>', unsafe_allow_html=True)

# Chat Session Management
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "Hello! I am your personal document assistant. Upload your files via the sidebar and ask me anything."}
    ]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])
        if "citations" in msg and msg["citations"]:
            with st.expander("🔍 View Verified Source Citations & Scores"):
                for c in msg["citations"]:
                    st.markdown(f"- **{c['source']}** (Page {c['page']}) — Semantic Score: `{c['relevance_score']}`")

# Query Input
if prompt := st.chat_input("Ask a question about your indexed documents..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.write(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Searching vector index & generating response..."):
            result = engine.generate_answer(prompt, top_k=top_k)
            st.write(result["answer"])
            
            if result["citations"]:
                with st.expander("🔍 View Verified Source Citations & Scores"):
                    for c in result["citations"]:
                        st.markdown(f"- **{c['source']}** (Page {c['page']}) — Semantic Score: `{c['relevance_score']}`")
                st.caption(f"⚡ Retrieval latency: {result['retrieval_latency_ms']} ms")

    st.session_state.messages.append({
        "role": "assistant",
        "content": result["answer"],
        "citations": result.get("citations", [])
    })
