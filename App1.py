from langchain.document_loaders import UnstructuredExcelLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_google_genai import ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings
from langchain.vectorstores.faiss import FAISS
from langchain_google_genai import GoogleGenerativeAIEmbeddings

# Build a function for Excel data
def prepare_excel(file_path):
    loader = UnstructuredExcelLoader(file_path)
    data = loader.load()

    # Split the text into chunks
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=10000,
        chunk_overlap=1000
    )
    chunks = text_splitter.split_documents(data)
    # Create embeddings with the OpenAI model
    embeddings = GoogleGenerativeAIEmbeddings(
        model="models/text-embedding-004",
        google_api_key="AIzaSyB_sMbpqCpd9u0HX-Fj7P9-x6S8YmA_fm4"  # <<-- quan trọng!
  )
    db_faiss = FAISS.from_documents(chunks, embeddings)
    return db_faiss

def rag(db_faiss, query, k=5):
    # Try the retrieval system
    output_retrieval = db_faiss.similarity_search(query, k=k)

    # Merge the context
    output_retrieval_merged = "\n".join([doc.page_content for doc in output_retrieval])

    # Define the prompt
    # Create a prompt for the rag system
    prompt = f"""
    based on this context: {output_retrieval_merged}
    answer the following question: {query}
    if you don't have information on the answer, say you don't know
    """
    model = ChatGoogleGenerativeAI(
    google_api_key="AIzaSyB_sMbpqCpd9u0HX-Fj7P9-x6S8YmA_fm4",
    model="models/gemini-2.0-flash-lite",  # hoặc "gemini-1.5-pro-latest"
    temperature=0
)
    response_text = model.invoke(prompt)
    return response_text.content

