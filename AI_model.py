import os
from langchain.document_loaders import UnstructuredExcelLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_google_genai import ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings
from langchain.vectorstores.faiss import FAISS
import langdetect


EMBEDDING_MODEL = "models/text-embedding-004"
CHAT_MODEL = "models/gemini-2.0-flash-lite"
GOOGLE_API_KEY = "AIzaSyB_sMbpqCpd9u0HX-Fj7P9-x6S8YmA_fm4"

def get_faiss_cache_path(file_path):
    return file_path + ".faiss"

def prepare_excel(file_path):
    cache_path = get_faiss_cache_path(file_path)
    embeddings = GoogleGenerativeAIEmbeddings(
        model=EMBEDDING_MODEL,
        google_api_key=GOOGLE_API_KEY
    )
    if os.path.exists(cache_path):
        db_faiss = FAISS.load_local(
            cache_path, 
            embeddings, 
            allow_dangerous_deserialization=True  # <-- add this!
        )
        print("Loaded FAISS index from cache.")
        return db_faiss

    loader = UnstructuredExcelLoader(file_path)
    data = loader.load()

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=5000, 
        chunk_overlap=200,
    )
    chunks = text_splitter.split_documents(data)
    db_faiss = FAISS.from_documents(chunks, embeddings)

    db_faiss.save_local(cache_path)
    print("Saved FAISS index to cache.")

    return db_faiss

def rag(db_faiss, query, k=5):
    lang = langdetect.detect(query)
    output_retrieval = db_faiss.similarity_search(query, k=k)
    output_retrieval_merged = "\n".join([doc.page_content for doc in output_retrieval])
    if lang == 'vi':
        prompt = f"""
    Dựa trên nội dung sau: {output_retrieval_merged}
    Hãy trả lời câu hỏi sau:{query}
    Nếu bạn không có đủ thông tin để đưa ra câu trả lời, hãy nói: "Tôi không biết."
"""
    else:
        prompt = f"""
    Based on this context: {output_retrieval_merged}  
    Answer the following question: {query}  
    If you don't have enough information to answer, say "I don't know."
"""
    model = ChatGoogleGenerativeAI(
        google_api_key=GOOGLE_API_KEY,
        model=CHAT_MODEL,
        temperature=0
    )
    response_text = model.invoke(prompt)
    return response_text.content

