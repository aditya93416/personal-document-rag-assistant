# Personal Document Q&A Assistant (RAG Pipeline)

**Author:** Aditya Kumar (Roll No: `024101010012`)  
**Degree Program:** B.Tech. Computer Science & Engineering (5th Semester)  
**Institution:** Haridwar University, Roorkee, Uttarakhand – 247667  
**Host Organisation:** InternCourse (`interncourse.online`), Patna, Bihar  
**Accreditations:** ISO 9001:2015 Certified, AICTE Partnered, MCA Registered  
**Mentors:** Pooja Parmar (Faculty Supervisor) & Dr. Himanshu Verma (HOD, CSE)  
**Industry Authority:** Ankit Pandey, Managing Director, InternCourse  
**Training Certificate ID:** `IC26/11873` (15 June 2026 – 18 July 2026)

---

## 📌 Project Overview
This repository contains the complete implementation and formal documentation for the **Personal Document Q&A Assistant**, built upon a privacy-preserving local **Retrieval-Augmented Generation (RAG)** pipeline.

The system addresses the fundamental limitation of Large Language Models (LLMs)—hallucination and temporal knowledge boundaries—by dynamically grounding responses strictly on private, local user documents (PDF, DOCX, TXT) with exact source and page citations.

---

## 🏗️ Architecture & Dataflow
1. **Document Ingestion:** Multi-format parsing using PyPDF & PyMuPDF with header/footer normalization.
2. **Recursive Text Splitting:** 500-token chunks with 50-token sliding overlap to prevent boundary context loss.
3. **Dense Vector Embeddings:** Encoded into 384-dimensional dense vectors using `sentence-transformers/all-MiniLM-L6-v2`.
4. **Vector Database:** Local embedded persistent vector indexing using **ChromaDB** with HNSW graphs.
5. **Similarity Retrieval:** Sub-15ms Cosine Similarity search retrieving Top-K ($K=4$) candidate chunks.
6. **Prompt Augmentation:** Injects retrieved passages under strict negative anti-hallucination constraints.
7. **Interactive UI Dashboard:** Full-stack dashboard implemented in Streamlit with real-time citation inspector.

---

## 📂 Repository Structure
```
Aditya_Kumar_RAG_Project/
├── src/
│   ├── config.py                 # System hyperparameters and configuration
│   ├── ingestion.py              # Parsing, chunking and ChromaDB ingestion pipeline
│   ├── query_engine.py           # Cosine similarity retrieval and grounded QA engine
│   ├── app.py                    # Streamlit web user interface dashboard
│   ├── test_pipeline.py          # Automated verification test suite
│   └── generate_aditya_report.py # Report compilation script
├── docs/
│   ├── Aditya_Kumar_Internship_Report.docx  # Full 23-page academic internship report (Word)
│   ├── Aditya_Kumar_Internship_Report.pdf   # Compiled PDF report
│   └── Aditya_Kumar_Cert.pdf                # Official InternCourse certificate
├── sample_docs/
│   └── rag_internship_overview.txt          # Sample document for indexing benchmarks
├── artifacts/
│   ├── aditya_cert_page_1.png               # High-resolution rendered certificate
│   └── screenshots/                         # UI dashboard & citation inspector snapshots
├── requirements.txt
└── README.md
```

---

## 🚀 Quickstart Guide

### 1. Installation
Clone the repository and set up a virtual environment:
```bash
git clone https://github.com/aditya93416/personal-document-rag-assistant.git
cd personal-document-rag-assistant
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Run Automated Verification Tests
```bash
python3 src/test_pipeline.py
```

### 3. Launch the Interactive Web Dashboard
```bash
streamlit run src/app.py
```

---

## 📊 Empirical Benchmarks (Chunk Size vs Accuracy)
| Chunk Size | Sliding Overlap | Latency | Hit Rate @ K=4 | Qualitative Assessment |
| :--- | :--- | :--- | :--- | :--- |
| 250 tokens | 25 tokens | 11.2 ms | 68.4% | Frequent sentence fragmentation |
| **500 tokens** | **50 tokens** | **14.8 ms** | **89.6%** | **Optimal balance & high precision** |
| 1000 tokens | 100 tokens | 22.6 ms | 74.2% | Diluted semantic density |
| 2000 tokens | 200 tokens | 38.1 ms | 59.8% | Exceeds context focus, attention dispersion |
