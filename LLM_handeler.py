import os
from dotenv import load_dotenv
from langchain_openai import AzureChatOpenAI
from langchain.vectorstores import Chroma
from langchain.embeddings import AzureOpenAIEmbeddings
from langchain.chains import RetrievalQA
import logging 
from typing import List
from langchain_core.documents import Document
from langchain.prompts import ChatPromptTemplate, SystemMessagePromptTemplate, HumanMessagePromptTemplate
from langchain.chains import ConversationalRetrievalChain
from langchain.schema import SystemMessage
from langchain.memory import ConversationBufferMemory
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.runnables import RunnableSequence
from langchain.chains.combine_documents import create_stuff_documents_chain
from langchain.chains import create_history_aware_retriever, create_retrieval_chain
from langchain_community.chat_message_histories import ChatMessageHistory

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
    handlers=[
        logging.FileHandler('converter.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

load_dotenv()

llm=AzureChatOpenAI(
        api_key=os.getenv("AZURE_OPENAI_API_KEY"),
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
        api_version=os.getenv("AZURE_API_VERSION"),
        azure_deployment=os.getenv("DEPLOYMENT_NAME"),
        temperature=0,
        max_tokens=512,
        timeout=None
    )



embeddings=AzureOpenAIEmbeddings(
    azure_endpoint=,
    azure_deployment=,
    openai_api_key=,
    openai_api_version=,
    chunk_size=1000
)

def save_to_chroma(docs: List[Document]):
    try:
        vectorstore = Chroma.from_documents(
            documents=docs,
            embedding=embeddings,
            persist_directory="chroma_db2"
        )
        vectorstore.persist()
    except Exception as e:
        logger.error(f"❌ Failed to save to Chroma: {e}")
        raise

def create_rag_chain(vectorstore):
    # 1. Chat history placeholder
    chat_history = ChatMessageHistory()

    # 2. Prompt for context-injected answering
    qa_prompt = ChatPromptTemplate.from_messages([
    ("system", """
    You are a highly reliable assistant that strictly answers questions based on the provided context. This context consists of standard operating manuals containing step-by-step procedures.

    Your goal is to guide users—especially those new to the system—by providing clear, structured, and instructional responses.

    When responding:
    - Use point-wise, step-by-step instructions whenever applicable.
    - Provide detailed explanations only if they are explicitly present in the documents.
    - Avoid assumptions, general knowledge, or unsupported inferences. Zero hallucination.

    Strict response guidelines:
    1. Base all responses solely on the retrieved content (marked as {context}). Never rely on external or general knowledge.
    2. If the answer is not found in the documents, reply politely with:
    "I'm sorry, but I couldn't find enough information in the provided documents to answer your question. Please refer to the appropriate manual or rephrase your query."
    3. Always use clear, simple, and beginner-friendly language.
    4. If the question relates to multiple procedures, mention all relevant ones before answering.
    5. If steps or procedures are listed in the documents, reproduce them exactly or summarize accurately while preserving their meaning and order.

    Be concise, polite, and completely grounded in the provided documentation.
    """),
        MessagesPlaceholder("chat_history"),
        ("human", "{input}"),
         ("user", "{context}")
        
    ])


    # 3. Context injection step (stuffing retrieved docs)
    question_answer_chain = create_stuff_documents_chain(llm, qa_prompt)

    # 4. Smart retriever for follow-ups
    
    question_prompt = ChatPromptTemplate.from_messages([
        ("system", "Given the conversation and a follow-up question, rephrase the question to be standalone and optimized for document search."),
        MessagesPlaceholder("chat_history"),
        ("human", "{input}")
    ])


    retriever = create_history_aware_retriever(
        llm=llm,
        retriever = vectorstore.as_retriever(
            search_type="mmr",  # 🧠 Improves semantic diversity
            search_kwargs={"k": 10, "fetch_k": 20}  # Fetch 5 chunks
        ),
        prompt=question_prompt  # ✅ Fixes the error,
    )

    # 5. Final RAG chain with chat memory
    rag_chain = create_retrieval_chain(
        retriever=retriever,
        combine_docs_chain=question_answer_chain
    )

    return rag_chain, chat_history




