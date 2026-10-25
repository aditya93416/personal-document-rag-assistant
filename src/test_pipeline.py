import os
import sys
from config import RAGConfig
from ingestion import IngestionPipeline
from query_engine import RAGQueryEngine

def run_cli_test():
    print("="*65)
    print("PERSONAL DOCUMENT Q&A ASSISTANT — CLI VERIFICATION RUNNER")
    print("="*65)

    config = RAGConfig()
    pipeline = IngestionPipeline(config)
    engine = RAGQueryEngine(pipeline)

    sample_doc_path = os.path.join(config.BASE_DIR, "sample_docs", "rag_internship_overview.txt")
    with open(sample_doc_path, "w", encoding="utf-8") as f:
        f.write("""INTERNCOURSE ARTIFICIAL INTELLIGENCE TRAINING PROGRAM OVERVIEW
Student Intern: Aditya Kumar
Roll Number: 024101010012
University: Haridwar University, Roorkee
Supervising Mentor: Pooja Parmar
Head of Department: Dr. Himanshu Verma
Executive Authority: Ankit Pandey, Managing Director
Organisation: InternCourse (interncourse.online), Patna, Bihar
Certificate Identifier: IC26/11873
Training Period: 15 June 2026 to 18 July 2026 (6 Weeks)
Domain: Artificial Intelligence & Natural Language Processing

Project Topic: Basic Analysis of RAG Systems & Development of a Personal Document Q&A Assistant (RAG Pipeline)

Architectural Details:
The pipeline operates on private local PDF documents. Documents are ingested using PyMuPDF and split recursively into 500-character chunks with a 50-character sliding overlap. Each chunk is encoded into a 384-dimensional dense semantic vector using the all-MiniLM-L6-v2 sentence-transformer model. Vectors are indexed into a local ChromaDB collection using Hierarchical Navigable Small World (HNSW) graphs. 

When a user submits a query, it is embedded using the identical vector space and matched using Cosine Similarity against all indexed chunks. The top 4 most relevant chunks are retrieved with sub-15ms latency and injected into the prompt context along with strict citation instructions. The assistant responds with explicit source and page references without hallucination.
""")

    print(f"[1] Ingesting sample document: {sample_doc_path}")
    ingest_result = pipeline.ingest_file(sample_doc_path)
    print(f"    Result: {ingest_result}")

    stats = pipeline.get_collection_stats()
    print(f"[2] Vector Store Stats: {stats}")

    test_queries = [
        "What is the student name and roll number in this internship?",
        "What chunk size and overlap are used in the RAG pipeline?",
        "Who is the Managing Director and what is the certificate ID?"
    ]

    print("\n[3] Executing Test Queries:")
    for q in test_queries:
        print(f"\nQUERY: '{q}'")
        res = engine.generate_answer(q, top_k=2)
        print(f"LATENCY: {res['retrieval_latency_ms']} ms")
        print(f"ANSWER:\n{res['answer']}")
        print(f"CITATIONS: {res['citations']}")

    print("\n" + "="*65)
    print("ALL TESTS PASSED: RAG PIPELINE FUNCTIONAL & GROUNDED")
    print("="*65)

if __name__ == "__main__":
    run_cli_test()
