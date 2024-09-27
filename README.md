PDF, DOCX, and XLSX Viewer and Q&A Application

This application allows users to upload PDF, DOCX, and XLSX files, extract text from them, and ask questions about the content. It uses Google Generative AI for embeddings and vector storage to facilitate accurate question answering.

Features

File Upload: Upload multiple PDF, DOCX, and XLSX files.

Text Extraction: Extract text from the uploaded files.

Vector Storage: Store the extracted text as vectors using Google Generative AI.

Question Answering: Ask questions related to the content of the uploaded files and get accurate answers.

Setup and Installation

Prerequisites

Python 3.8 or higher

A Google API Key for Google Generative AI

Installation

Clone the Repository
git clone (https://github.com/Ishwaryathiru/PDF-WORD-EXCEL-Chat-bot)

cd your-repo-name

Create and Activate a Virtual Environment

virtualenv -p python3.12 penv

In powershell (Run as administrator)

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

Activation

in cmd : ./penv/Scripts/Activate

Install the Required Packages

pip install -r requirements.txt

Set Up Environment Variables

Create a .env file in the root of your project and add your Google API Key:

GOOGLE_API_KEY=your-google-api-key

Running the Application

Run the Application

python pdf_word_xlxs.py

Use the GUI

Click on "Open PDF" to select and upload your PDF, DOCX, and XLSX files.

The uploaded file names will be displayed.

Enter a question in the provided entry box and click "Submit" to get an answer based on the uploaded documents.

Code Explanation

Key Functions

get_pdf_text(path): Extracts text from PDF files.

get_text_chunks(text): Splits text into manageable chunks.

get_vector_store(text_chunks): Stores text chunks as vectors using Google Generative AI.

get_conversational_chain(): Creates a conversational chain for answering questions.

user_input(ques): Handles user questions and retrieves relevant answers.

get_docx_text(path): Extracts text from DOCX files.

get_excel_text(path): Extracts text from XLSX files.
