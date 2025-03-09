# SharePoint Data Query with FAISS and LangChain

This project demonstrates how to retrieve data from a SharePoint site using the Microsoft Graph API, convert the data into vector embeddings with OpenAI, index them using FAISS, and then query the data interactively using LangChain's RetrievalQA chain with a ChatOpenAI model.

## Overview

The script performs the following steps:

1. **Authentication:**  
   Retrieves an access token from Microsoft using your tenant credentials.

2. **Data Retrieval:**  
   Fetches lists or other data from a specified SharePoint site via the Microsoft Graph API.

3. **Vectorization & Indexing:**  
   Converts the SharePoint data into vector embeddings using OpenAI embeddings, creates a FAISS index, and saves it locally.

4. **Querying:**  
   Loads (or creates) the FAISS index and uses LangChain's RetrievalQA chain to allow interactive querying of your SharePoint data.

## Features

- **Microsoft Graph API Integration:** Securely fetch SharePoint data.
- **OpenAI Embeddings:** Convert textual data into meaningful vector representations.
- **FAISS Indexing:** Efficiently index and search vector embeddings.
- **Interactive Querying:** Ask questions and retrieve answers based on your SharePoint content.

## Prerequisites

- **Python Version:** 3.7 or higher
- **Libraries:**  
  - `faiss-cpu` (or `faiss-gpu` if you have GPU support)  
  - `numpy`  
  - `requests`  
  - `langchain_community` (for embeddings, vectorstores, and chat models)  
  - `langchain`  
- **APIs & Credentials:**  
  - Microsoft Graph API credentials:
    - `TENANT_ID`
    - `CLIENT_ID`
    - `CLIENT_SECRET`
    - `SHAREPOINT_SITE_ID`
  - OpenAI API key (`OPENAI_API_KEY`)

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/yourusername/your-repo-name.git
   cd your-repo-name
