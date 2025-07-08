from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import AzureOpenAIEmbeddings
# from langchain_community.embeddings import HuggingFaceEmbeddings
from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import os
import tempfile
import asyncio
import logging
import sys
import io
from agent_prog import (
    agent_file_processor,
    check_pdf_type,
    normal_pdf_processor,
    extract_text_to_markdown,
    convert_docx_to_temp_pdf,
    ppt_to_pdf_win32com,
    xlsx_to_mrkdwn,
    csv_to_mrkdwn,
    txt_to_mrkdwn,
    extract_text_to_tempfile,
    refine_markdown_structure,
    _format_text_to_markdown,
    is_rag_compatible,
    # convert_pdf_to_structured_pdf
)
from LLM_handeler import (
    create_rag_chain,
    save_to_chroma,
    embeddings
)
from splitter import (
    chunk_by_markdown_headers
)

from dotenv import load_dotenv
load_dotenv()


sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


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

app = FastAPI(
    title="Folder to Markdown Converter",
    description="API to convert all supported files in a folder to markdown text",
    version="2.0.0"
)

class FolderRequest(BaseModel):
    folder_path: str

@app.post("/convert-folder/")
async def convert_folder(request: FolderRequest):
    folder_path = request.folder_path
    logger.debug(f"Starting folder conversion: {folder_path}")
    if not os.path.isdir(folder_path):
        logger.error(f"Invalid folder path: {folder_path}")
        raise HTTPException(status_code=400, detail="Invalid folder path")

    loop = asyncio.get_event_loop()

    async def process_file(full_path: str, filename: str):
        logger.debug(f"Processing file: {full_path}")
        try:
            file_type = await loop.run_in_executor(None, agent_file_processor, full_path)
            markdown_text = ""
            rag_warning = ""
            compatibility_details = {}

            if file_type == '.pdf':
                pdf_type = await loop.run_in_executor(None, check_pdf_type, full_path)
                if pdf_type in ["scanned", "hybrid"]:
                    markdown_text = await loop.run_in_executor(None, extract_text_to_markdown, full_path, "eng", 300)
                else:
                    markdown_text = await loop.run_in_executor(None, normal_pdf_processor, full_path)


                markdown_text = await loop.run_in_executor(None, _format_text_to_markdown, markdown_text)
                markdown_text = await loop.run_in_executor(None, refine_markdown_structure, markdown_text)
                
                if isinstance(markdown_text, str) and "\\n" in markdown_text:
                    markdown_text = markdown_text.encode("utf-8").decode("unicode_escape")
                rag_compatible, details = await loop.run_in_executor(
                    None, lambda: is_rag_compatible(markdown_text, return_details=True)
                )
                compatibility_details = details
                if not rag_compatible:
                    rag_warning = format_compatibility_warning(details)
                    logger.warning(f"RAG compatibility warning for {filename}: {rag_warning}")

            elif file_type == '.docx':
                try:
                    pdf_path = await loop.run_in_executor(None, convert_docx_to_temp_pdf, full_path)
                    markdown_text = await loop.run_in_executor(None, extract_text_to_markdown, pdf_path, "eng", 300)
                    markdown_text = await loop.run_in_executor(None, _format_text_to_markdown, markdown_text)
                    markdown_text = await loop.run_in_executor(None, refine_markdown_structure, markdown_text)
                    os.unlink(pdf_path)
                    logger.debug(f"Deleted temporary PDF: {pdf_path}")
                    if isinstance(markdown_text, str) and "\\n" in markdown_text:
                        markdown_text = markdown_text.encode("utf-8").decode("unicode_escape")
                    rag_compatible, details = await loop.run_in_executor(
                        None, lambda: is_rag_compatible(markdown_text, return_details=True)
                    )
                    compatibility_details = details
                    if not rag_compatible:
                        rag_warning = format_compatibility_warning(details)
                        logger.warning(f"RAG compatibility warning for {filename}: {rag_warning}")
                except Exception as e:
                    markdown_text = f"Failed to process DOCX: {str(e)}"
                    logger.error(f"DOCX processing failed for {filename}: {str(e)}")

            elif file_type == '.pptx':
                try:
                    ppt_pdf_path = await loop.run_in_executor(None, ppt_to_pdf_win32com, full_path)
                    markdown_text = await loop.run_in_executor(None, extract_text_to_markdown, ppt_pdf_path, "eng", 300)
                    markdown_text = await loop.run_in_executor(None, _format_text_to_markdown, markdown_text)
                    markdown_text = await loop.run_in_executor(None, refine_markdown_structure, markdown_text)
                    os.unlink(ppt_pdf_path)
                    logger.debug(f"Deleted temporary PDF: {ppt_pdf_path}")
                    if isinstance(markdown_text, str) and "\\n" in markdown_text:
                        markdown_text = markdown_text.encode("utf-8").decode("unicode_escape")
                    rag_compatible, details = await loop.run_in_executor(
                        None, lambda: is_rag_compatible(markdown_text, return_details=True)
                    )
                    compatibility_details = details
                    if not rag_compatible:
                        rag_warning = format_compatibility_warning(details)
                        logger.warning(f"RAG compatibility warning for {filename}: {rag_warning}")
                except Exception as e:
                    markdown_text = f"Failed to process PPTX: {str(e)}"
                    logger.error(f"PPTX processing failed for {filename}: {str(e)}")

            elif file_type in ['.xlsx', '.csv', '.txt']:
                markdown_text = await loop.run_in_executor(None, {
                    '.xlsx': xlsx_to_mrkdwn,
                    '.csv': csv_to_mrkdwn,
                    '.txt': txt_to_mrkdwn
                }[file_type], full_path)
                markdown_text = await loop.run_in_executor(None, _format_text_to_markdown, markdown_text)
                markdown_text = await loop.run_in_executor(None, refine_markdown_structure, markdown_text)
                if isinstance(markdown_text, str) and "\\n" in markdown_text:
                    markdown_text = markdown_text.encode("utf-8").decode("unicode_escape")
                rag_compatible, details = await loop.run_in_executor(
                    None, lambda: is_rag_compatible(markdown_text, return_details=True)
                )
                compatibility_details = details
                if not rag_compatible:
                    rag_warning = format_compatibility_warning(details)
                    logger.warning(f"RAG compatibility warning for {filename}: {rag_warning}")

            elif file_type == '.png':
                try:
                    temp_txt_path = await loop.run_in_executor(None, extract_text_to_tempfile, full_path)
                    markdown_text = await loop.run_in_executor(None, txt_to_mrkdwn, temp_txt_path)
                    markdown_text = await loop.run_in_executor(None, _format_text_to_markdown, markdown_text)
                    markdown_text = await loop.run_in_executor(None, refine_markdown_structure, markdown_text)
                    os.unlink(temp_txt_path)
                    logger.debug(f"Deleted temporary text file: {temp_txt_path}")
                    if isinstance(markdown_text, str) and "\\n" in markdown_text:
                        markdown_text = markdown_text.encode("utf-8").decode("unicode_escape")
                    rag_compatible, details = await loop.run_in_executor(
                        None, lambda: is_rag_compatible(markdown_text, return_details=True)
                    )
                    compatibility_details = details
                    if not rag_compatible:
                        rag_warning = format_compatibility_warning(details)
                        logger.warning(f"RAG compatibility warning for {filename}: {rag_warning}")
                except Exception as e:
                    markdown_text = f"Failed to process PNG: {str(e)}"
                    logger.error(f"PNG processing failed for {filename}: {str(e)}")

            else:
                markdown_text = "Unsupported file format"
                compatibility_details = {"reason": "Unsupported file type"}
                logger.warning(f"Unsupported file format: {filename}")

            if rag_compatible:
                logger.info("\n==================Developer is testing logs========================================\n")
                logger.info(f"✅ File {filename} is RAG compatible, preparing to chunk and store in Chroma.")
                docs = await loop.run_in_executor(None, chunk_by_markdown_headers, markdown_text)
                logger.info(f"📄 Chunked {len(docs)} documents. Now saving to Chroma.")
                await loop.run_in_executor(None, save_to_chroma, docs)
            else:
                if rag_warning:
                    markdown_text += f"\n\n{rag_warning}"

            logger.info(f"Completed processing for {filename}: rag_compatible={compatibility_details.get('is_compatible', False)}")
            return filename, {
                "file_type": file_type,
                "markdown_text": markdown_text,
                "rag_compatible": compatibility_details.get("is_compatible", False),
                "compatibility_details": compatibility_details
            }


        except Exception as e:
            logger.error(f"Unexpected error processing {filename}: {str(e)}")
            return filename, {
                "error": str(e),
                "file_type": file_type if 'file_type' in locals() else "unknown"
            }

    def format_compatibility_warning(details: dict) -> str:
        reasons = []
        if not details.get("length_ok", details.get("length", 0) >= 100):
            reasons.append(f"Content too short ({details.get('length')} chars, needs 100+)")
        if not details.get("structure_ok", details.get("structure_score", 0) >= 1):
            reasons.append(f"Needs more structure (current: {details.get('structure_score')}/2)")
        warning = "<!-- RAG_COMPATIBILITY_WARNING: Structure may need manual review -->\n" + \
                 "\n".join(f"- {reason}" for reason in reasons)
        logger.debug(f"Generated RAG compatibility warning: {warning}")
        return warning

    semaphore = asyncio.Semaphore(10)
    async def process_with_semaphore(full_path, filename):
        async with semaphore:
            return await process_file(full_path, filename)

    tasks = []
    for root, _, files in os.walk(folder_path):
        for filename in files:
            full_path = os.path.join(root, filename)
            tasks.append(process_with_semaphore(full_path, filename))

    results = dict(await asyncio.gather(*tasks))
    
    success_count = sum(1 for v in results.values() if "error" not in v)
    rag_compatible_count = sum(1 for v in results.values() if v.get("rag_compatible", False))
    warning_count = sum(1 for v in results.values() if "error" not in v and not v.get("rag_compatible", False))
    
    logger.info(f"Folder conversion summary: total={len(results)}, successful={success_count}, "
                f"rag_compatible={rag_compatible_count}, warnings={warning_count}, errors={len(results) - success_count}")

    return JSONResponse(content={
        "files": results,
        "summary": {
            "total_files": len(results),
            "successful_conversions": success_count,
            "rag_compatible_files": rag_compatible_count,
            "files_needing_review": warning_count,
            "error_count": len(results) - success_count,
            "compatibility_thresholds": {
                "min_length": 100,
                "min_structure": 1
            }
        }
    })
# Initialize once
vs = Chroma(persist_directory="chroma_db2", embedding_function=embeddings)
rag_chain, chat_history = create_rag_chain(vs)

@app.post("/ask/")
async def ask_question(query: str):
    try:
        response = rag_chain.invoke({
            "input": query,
            "chat_history": chat_history.messages
        })

        # Append to memory manually
        chat_history.add_user_message(query)
        chat_history.add_ai_message(response["answer"])

        return {
            "question": query,
            "answer": response["answer"]
        }
    except Exception as e:
        logger.error(f"Error in ask_question: {e}")
        return {"error": str(e)}


@app.get("/")
async def root():
    logger.debug("Accessed root endpoint")
    return {"message": "Welcome to Folder to Markdown Converter API"}