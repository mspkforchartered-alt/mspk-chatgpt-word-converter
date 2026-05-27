# MSPK AI to Word Converter

Professional AI-powered converter for transforming ChatGPT, AI-generated notes, equations, financial calculations, and markdown content into properly formatted Microsoft Word documents with native Word equations.

---

# Features

✅ Native Microsoft Word equations (OMML)

✅ LaTeX equation conversion

✅ Multiline equation support

✅ Financial management formulas support

✅ Markdown table conversion

✅ Professional document formatting

✅ AI-generated content cleanup

✅ FastAPI backend architecture

✅ Deployable on Render

✅ Supports CA / CMA / Finance notes

---

# Supported Content

- Financial Management formulas
- WACC calculations
- Leverages
- Cost of Capital
- Statistics equations
- Calculus equations
- Probability formulas
- Matrix equations
- Markdown tables
- Mixed AI-generated text

---

# Tech Stack

## Backend
- FastAPI
- python-docx
- latex2mathml
- lxml
- MML2OMML.XSL

## Frontend
- HTML
- CSS
- JavaScript

---

# Project Structure

```text
root/
│
├── backend/
│   ├── main.py
│   ├── requirements.txt
│   ├── MML2OMML.XSL
│
├── frontend/
│
├── LICENSE
└── README.md
```

---

# Installation

## Clone Repository

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
```

---

## Backend Setup

```bash
cd backend
pip install -r requirements.txt
```

---

## Run Backend

```bash
uvicorn main:app --reload
```

---

# API Endpoint

## Generate DOCX

```http
POST /generate-docx
```

---

# Deployment

## Backend
Recommended:
- Render

## Frontend
Recommended:
- Netlify

---

# License

This project is licensed under the MIT License.

---

# Author

Muhammed Sinan P K

MSPK FOR CHARTERED