import os
from dataclasses import dataclass

@dataclass
class RAGConfig:
    # Model parameters
    EMBEDDING_MODEL_NAME: str = "sentence-transformers/all-MiniLM-L6-v2"
    CHUNK_SIZE: int = 500
    CHUNK_OVERLAP: int = 50
    TOP_K: int = 4
    SIMILARITY_THRESHOLD: float = 0.65
    
    # Paths
    BASE_DIR: str = os.path.dirname(os.path.abspath(__file__))
    PERSIST_DIR: str = os.path.join(BASE_DIR, "data", "chroma_db")
    DOCS_DIR: str = os.path.join(BASE_DIR, "data", "documents")

    # System instruction
    SYSTEM_PROMPT: str = (
        "You are an expert academic and technical Personal Document Q&A Assistant. "
        "Your task is to answer user queries strictly based on the provided retrieved context. "
        "Rules:\n"
        "1. Never fabricate information or hallucinate facts beyond the given context.\n"
        "2. If the context does not contain sufficient information, state clearly: "
        "'Information not found in reference documents.'\n"
        "3. Explicitly cite the document source and page number for every factual statement."
    )
