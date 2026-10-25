import os
from typing import List, Dict, Any
from langchain_community.document_loaders import PyPDFLoader, TextLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

from config import RAGConfig

class IngestionPipeline:
    """Document Ingestion, Chunking, and Embedding subsystem."""

    def __init__(self, config: RAGConfig = None):
        self.config = config or RAGConfig()
        os.makedirs(self.config.PERSIST_DIR, exist_ok=True)
        os.makedirs(self.config.DOCS_DIR, exist_ok=True)
        
        print(f"[INFO] Initializing Embedding Model: {self.config.EMBEDDING_MODEL_NAME}")
        self.embeddings = HuggingFaceEmbeddings(
            model_name=self.config.EMBEDDING_MODEL_NAME,
            model_kwargs={"device": "cpu"}
        )
        
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=self.config.CHUNK_SIZE,
            chunk_overlap=self.config.CHUNK_OVERLAP,
            separators=["\n\n", "\n", ". ", " ", ""],
            length_function=len
        )
        
        self.vector_store = Chroma(
            persist_directory=self.config.PERSIST_DIR,
            embedding_function=self.embeddings
        )

    def load_document(self, file_path: str) -> List[Any]:
        ext = os.path.splitext(file_path)[-1].lower()
        if ext == ".pdf":
            loader = PyPDFLoader(file_path)
        elif ext in [".txt", ".md"]:
            loader = TextLoader(file_path, encoding="utf-8")
        else:
            raise ValueError(f"Unsupported file format: {ext}")
        return loader.load()

    def ingest_file(self, file_path: str) -> Dict[str, Any]:
        """Ingest a single document into the vector store."""
        print(f"[INFO] Ingesting file: {file_path}")
        docs = self.load_document(file_path)
        chunks = self.text_splitter.split_documents(docs)
        
        # Add metadata tags
        filename = os.path.basename(file_path)
        for chunk in chunks:
            chunk.metadata["filename"] = filename
            chunk.metadata["char_count"] = len(chunk.page_content)

        # Store in ChromaDB
        self.vector_store.add_documents(chunks)
        self.vector_store.persist()
        
        return {
            "filename": filename,
            "raw_pages": len(docs),
            "total_chunks": len(chunks),
            "status": "Indexed Successfully"
        }

    def get_collection_stats(self) -> Dict[str, Any]:
        count = self.vector_store._collection.count()
        return {
            "total_vectors": count,
            "persist_directory": self.config.PERSIST_DIR,
            "embedding_dimension": 384
        }
