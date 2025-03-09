import os
import faiss
import numpy as np
import traceback
import requests
from langchain_community.embeddings import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain.schema import Document
from langchain_community.chat_models import ChatOpenAI
from langchain.chains import RetrievalQA

# Your constants (ensure these are defined somewhere in your code)
TENANT_ID =" "
CLIENT_ID ="  "
CLIENT_SECRET =" "
SHAREPOINT_SITE_ID =""
OPENAI_API_KEY =""


# Step 1: Get Access Token
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(url, data=data)

    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        print(f"❌ Error getting token: {response.text}")
        return None


# Step 2: Fetch SharePoint Data
def fetch_sharepoint_data():
    access_token = get_access_token()
    if not access_token:
        print("❌ Authentication failed. Exiting...")
        return []

    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/lists"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print(f"❌ Error fetching SharePoint data: {response.text}")
        return []


# Step 3: Convert Text Data to Vectors and Create FAISS Index
def create_vector_store():
    sharepoint_data = fetch_sharepoint_data()

    if not sharepoint_data:
        print("❌ No data fetched from SharePoint. Exiting...")
        return None

    embedding_model = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)

    # Convert SharePoint data into Documents
    documents = [Document(page_content=item['name'], metadata={"source": "SharePoint"}) for item in sharepoint_data]

    # Convert Documents to Embeddings (Vectors)
    embeddings = []
    for doc in documents:
        vector = embedding_model.embed([doc.page_content])[0]  # Embedding the text
        embeddings.append(vector)

    embeddings = np.array(embeddings).astype("float32")  # Ensure correct type for FAISS

    # Create FAISS Index
    index = faiss.IndexFlatL2(embeddings.shape[1])  # L2 similarity (Euclidean distance)
    index.add(embeddings)  # Add embeddings to the index

    # Save the index to disk
    index_path = "sharepoint_index/index.faiss"
    os.makedirs(os.path.dirname(index_path), exist_ok=True)  # Ensure directory exists
    faiss.write_index(index, index_path)

    print("✅ Vector database created and saved.")
    return index


# Step 4: Query SharePoint Data
def query_sharepoint():
    embedding_model = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)

    try:
        # Ensure directory exists for FAISS index
        if not os.path.exists("sharepoint_index"):
            os.makedirs("sharepoint_index")

        # Try to load the FAISS index
        index = faiss.read_index("sharepoint_index/index.faiss")
        print("✅ Loaded existing FAISS index.")

    except Exception as e:
        print(traceback.format_exc())
        print(f"❌ Error loading FAISS: {e}")
        print("⚠️ FAISS index not found. Creating a new one...")

        # If the index doesn't exist, create a new one
        index = create_vector_store()
        if index is None:
            print("❌ Failed to create FAISS index. Exiting...")
            exit(1)

    # Create a retriever from the FAISS index
    retriever = FAISS(index)
    qa_chain = RetrievalQA.from_chain_type(
        llm=ChatOpenAI(model="gpt-4", openai_api_key=OPENAI_API_KEY),
        retriever=retriever
    )

    # Query loop
    while True:
        query = input("\n🔎 Ask a question (or type 'exit' to quit): ")
        if query.lower() == 'exit':
            break

        response = qa_chain.run(query)
        print(f"\n💡 Answer: {response}")


# Run the full process
if __name__ == "__main__":
    query_sharepoint()
