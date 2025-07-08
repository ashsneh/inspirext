
import os
import re
import fitz
import pytesseract
import tempfile
import logging
import pdfplumber
import pandas as pd
from PIL import Image, ImageEnhance, ImageFilter
from pdfminer.high_level import extract_text
# from fpdf import FPDF
from docx2pdf import convert
from pdf2image import convert_from_path
import win32com.client
from markdownify import markdownify

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



def agent_file_processor(file_path):
    logger.debug(f"Processing file: {file_path}")
    lower_path = file_path.lower()
    if lower_path.endswith('.pdf'):
        logger.info(f"Detected PDF file: {file_path}")
        return ".pdf"
    elif lower_path.endswith('.txt'):
        logger.info(f"Detected TXT file: {file_path}")
        return ".txt"
    elif lower_path.endswith('.docx'):
        logger.info(f"Detected DOCX file: {file_path}")
        return ".docx"
    elif lower_path.endswith('.csv'):
        logger.info(f"Detected CSV file: {file_path}")
        return ".csv"
    elif lower_path.endswith('.xlsx'):
        logger.info(f"Detected XLSX file: {file_path}")
        return ".xlsx"
    elif lower_path.endswith('.pptx') or lower_path.endswith('.ppt'):
        logger.info(f"Detected PPT/PPTX file: {file_path}")
        return ".pptx"
    elif lower_path.endswith('.jpg') or lower_path.endswith('.jpeg') or lower_path.endswith('.png'):
        logger.info(f"Detected image file: {file_path}")
        return ".png"
    else:
        logger.warning(f"Unsupported file type: {file_path}")
        return "Unsupported file type"    


def normal_pdf_processor(file_path: str) -> str:
    """Enhanced PDF processor for text-based PDFs"""
    logger.debug(f"Starting normal PDF processing: {file_path}")
    markdown = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    logger.debug(f"No text extracted from page in {file_path}")
                    continue
                lines = text.splitlines()
                current_section = []
                for line in lines:
                    line = line.strip()
                    if not line or line == "[OFFICIAL]":
                        continue
                    if (line.upper() == line and len(line) < 50 and not line.endswith(".")) or \
                       line.lower().startswith(("step", "procedure", "section", "chapter")):
                        if current_section:
                            markdown.append(" ".join(current_section))
                        markdown.append(f"## {line}")
                        logger.debug(f"Detected heading: {line}")
                        current_section = []
                    elif re.match(r"^(Click|Select|Enter|At the|For e\.g|\d+[\.\)]\s+|Step \d+[\.:]?\s+|-|\•)", line, re.IGNORECASE):
                        if current_section:
                            markdown.append(" ".join(current_section))
                        cleaned_line = re.sub(r"^\d+[\.\)]\s+|Step \d+[\.:]?\s+|-\s+|\•\s+", "", line)
                        markdown.append(f"- {cleaned_line}")
                        logger.debug(f"Detected step: {cleaned_line}")
                        current_section = []
                    else:
                        current_section.append(line)
                if current_section:
                    markdown.append(" ".join(current_section))
        logger.info(f"Completed normal PDF processing: {file_path}")
        return "\n\n".join(markdown)
    except Exception as e:
        logger.error(f"Failed to process PDF {file_path}: {str(e)}")
        raise

def refine_markdown_structure(md_text: str) -> str:
    """Refine Markdown to ensure proper structure for RAG"""
    logger.debug("Refining Markdown structure")
    lines = md_text.splitlines()
    refined = []
    current_step = []
    in_step = False

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("#"):
            if current_step:
                refined.append(" ".join(current_step))
                logger.debug(f"Grouped step: {' '.join(current_step)}")
                current_step = []
                in_step = False
            refined.append(line)
        elif line.startswith("-"):
            if current_step:
                refined.append(" ".join(current_step))
                logger.debug(f"Grouped step: {' '.join(current_step)}")
            current_step = [line]
            in_step = True
        elif in_step:
            current_step.append(line)
        else:
            if current_step:
                refined.append(" ".join(current_step))
                logger.debug(f"Grouped step: {' '.join(current_step)}")
                current_step = []
                in_step = False
            refined.append(line)
    if current_step:
        refined.append(" ".join(current_step))
        logger.debug(f"Grouped final step: {' '.join(current_step)}")
    logger.info("Completed Markdown refinement")
    return "\n\n".join(refined)

def _format_text_to_markdown(text: str) -> str:
    """Convert text to structured Markdown with robust pattern detection"""
    logger.debug("Formatting text to Markdown")
    markdown = []
    lines = text.splitlines()
    in_procedure = False
    current_section = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if re.match(r"^(To|Steps?|Procedure|Section|Chapter|[A-Z][A-Za-z\s]{5,50}:)", line, re.IGNORECASE):
            if current_section:
                markdown.append(" ".join(current_section))
                logger.debug(f"Grouped section: {' '.join(current_section)}")
            markdown.append(f"## {line}")
            logger.debug(f"Detected heading: {line}")
            in_procedure = True
            current_section = []
        elif re.match(r"^\d+[\.\)]\s+|Step \d+[\.:]?\s+|-\s+|\•\s+", line, re.IGNORECASE):
            if current_section:
                markdown.append(" ".join(current_section))
                logger.debug(f"Grouped section: {' '.join(current_section)}")
            cleaned_line = re.sub(r"^\d+[\.\)]\s+|Step \d+[\.:]?\s+|-\s+|\•\s+", "", line)
            markdown.append(f"- {cleaned_line}")
            logger.debug(f"Detected step: {cleaned_line}")
            in_procedure = True
            current_section = []
        else:
            current_section.append(line)
            in_procedure = False
    if current_section:
        markdown.append(" ".join(current_section))
        logger.debug(f"Grouped final section: {' '.join(current_section)}")
    logger.info("Completed text to Markdown formatting")
    return "\n\n".join(markdown)


def check_pdf_type(pdf_path: str) -> str:
    """Determine PDF type with robust checks"""
    logger.debug(f"Checking PDF type: {pdf_path}")
    try:
        text = extract_text(pdf_path)
        doc = fitz.open(pdf_path)
        has_images = any(page.get_image_info() for page in doc)
        doc.close()
        text_length = len(text.strip())
        if text_length > 100:
            logger.info(f"PDF classified as text: {pdf_path}")
            return "text"
        elif has_images and text_length < 20:
            logger.info(f"PDF classified as scanned: {pdf_path}")
            return "scanned"
        elif has_images and text_length >= 20:
            logger.info(f"PDF classified as hybrid: {pdf_path}")
            return "hybrid"
        else:
            logger.info(f"PDF classified as scanned (default): {pdf_path}")
            return "scanned"
    except Exception as e:
        logger.error(f"PDF type analysis error for {pdf_path}: {str(e)}")
        return "scanned"


def extract_text_to_markdown(pdf_path: str, lang: str = "eng", dpi: int = 300) -> str:
    """
    Extracts text from a PDF (scanned or text-based) using PyMuPDF, pytesseract, and pdfplumber,
    then converts to Markdown using markdownify for robust structure.
    """
    logger.debug(f"Starting text extraction for PDF: {pdf_path}")
    try:
        fitz.TOOLS.mupdf_warnings(reset=False)
        doc = fitz.open(pdf_path)
        extracted_text = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text and len(text.strip()) > 20:
                        extracted_text.append(text.strip())
                        logger.debug(f"Extracted text with pdfplumber from page in {pdf_path}")
        except Exception as e:
            logger.warning(f"pdfplumber failed for {pdf_path}: {str(e)}")

        if not extracted_text or sum(len(t) for t in extracted_text) < 100:
            logger.info(f"Falling back to OCR for {pdf_path}")
            extracted_text = []
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img = img.convert('L')
                img = ImageEnhance.Contrast(img).enhance(1.5)
                img = img.filter(ImageFilter.MedianFilter())
                img = img.point(lambda x: 0 if x < 150 else 255)
                img = img.filter(ImageFilter.SHARPEN)
                text = pytesseract.image_to_string(img, lang=lang, config='--psm 6')
                extracted_text.append(text.strip())
                logger.debug(f"Page {page_num + 1} OCR output: {text}")

        doc.close()
        combined_text = "\n\n".join(extracted_text)
        logger.debug(f"Combined text length: {len(combined_text)} characters")

        from markdownify import markdownify
        markdown_text = markdownify(combined_text, heading_style="ATX")
        logger.info(f"Converted to Markdown: {pdf_path}")
        return markdown_text

    except FileNotFoundError:
        logger.error(f"PDF file not found: {pdf_path}")
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        logger.error(f"Processing failed for {pdf_path}: {str(e)}")
        raise Exception(f"Processing failed: {str(e)}")
    
# import pythoncom
# pythoncom.CoInitialize()

# def convert_docx_to_temp_pdf(docx_path: str) -> str:
#     """Convert DOCX to PDF with error handling"""
#     logger.debug(f"Converting DOCX to PDF: {docx_path}")
#     try:
#         with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
#             temp_pdf_path = temp_pdf.name
#         convert(docx_path, temp_pdf_path)
#         logger.info(f"Converted DOCX to PDF: {temp_pdf_path}")
#         return temp_pdf_path
#     except Exception as e:
#         logger.error(f"DOCX conversion failed for {docx_path}: {str(e)}")
#         raise Exception(f"DOCX conversion failed: {str(e)}")
import pythoncom
import tempfile
from win32com import client
import os

def convert_docx_to_temp_pdf(docx_path: str) -> str:
    """Convert DOCX to PDF with proper COM initialization"""
    logger.debug(f"Converting DOCX to PDF: {docx_path}")
    pythoncom.CoInitialize()  # Initialize COM for this thread
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
            temp_pdf_path = temp_pdf.name

        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(temp_pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
        word.Quit()

        logger.info(f"Converted DOCX to PDF: {temp_pdf_path}")
        return temp_pdf_path
    except Exception as e:
        logger.error(f"DOCX conversion failed for {docx_path}: {str(e)}")
        raise Exception(f"DOCX conversion failed: {str(e)}")
    finally:
        pythoncom.CoUninitialize()  # Ensure COM is uninitialized


def ppt_to_pdf_win32com(ppt_path: str) -> str:
    """Convert PPT/PPTX to PDF with fallback for non-Windows systems"""
    logger.debug(f"Converting PPT/PPTX to PDF: {ppt_path}")
    try:
        import platform
        if platform.system() != "Windows":
            logger.error(f"PPT conversion requires Windows: {ppt_path}")
            raise Exception("PPT conversion requires Windows")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
                pdf_path = temp_pdf.name
            deck = powerpoint.Presentations.Open(os.path.abspath(ppt_path))
            deck.SaveAs(pdf_path, 32)
            deck.Close()
            logger.info(f"Converted PPT to PDF: {pdf_path}")
            return pdf_path
        finally:
            powerpoint.Quit()
    except Exception as e:
        logger.error(f"PPT conversion failed for {ppt_path}: {str(e)}")
        raise Exception(f"PPT conversion failed: {str(e)}")

def xlsx_to_mrkdwn(file_path: str) -> str:
    """Convert Excel to Markdown using pandas"""
    logger.debug(f"Converting XLSX to Markdown: {file_path}")
    try:
        df = pd.read_excel(file_path)
        markdown = f"## Excel Data\n\n{df.to_markdown(index=False)}\n"
        logger.info(f"Converted XLSX to Markdown: {file_path}")
        return markdown
    except Exception as e:
        logger.error(f"Excel conversion failed for {file_path}: {str(e)}")
        raise Exception(f"Excel conversion failed: {str(e)}")

def csv_to_mrkdwn(file_path: str) -> str:
    """Convert CSV to Markdown using pandas"""
    logger.debug(f"Converting CSV to Markdown: {file_path}")
    try:
        df = pd.read_csv(file_path)
        markdown = f"## CSV Data\n\n{df.to_markdown(index=False)}\n"
        logger.info(f"Converted CSV to Markdown: {file_path}")
        return markdown
    except Exception as e:
        logger.error(f"CSV conversion failed for {file_path}: {str(e)}")
        raise Exception(f"CSV conversion failed: {str(e)}")
    
def txt_to_mrkdwn(file_path: str) -> str:
    """Convert TXT to Markdown with structure detection"""
    logger.debug(f"Converting TXT to Markdown: {file_path}")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        markdown = _format_text_to_markdown(text)
        logger.info(f"Converted TXT to Markdown: {file_path}")
        return markdown
    except Exception as e:
        logger.error(f"TXT conversion failed for {file_path}: {str(e)}")
        raise Exception(f"TXT conversion failed: {str(e)}")

def extract_text_to_tempfile(image_path: str) -> str:
    """Extract text from images with advanced preprocessing"""
    logger.debug(f"Extracting text from image: {image_path}")
    try:
        with Image.open(image_path) as img:
            img = img.convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.0)
            img = img.point(lambda x: 0 if x < 150 else 255)
            img = img.filter(ImageFilter.SHARPEN)
            text = pytesseract.image_to_string(img, config='--psm 6').strip()
        with tempfile.NamedTemporaryFile(mode='w+', suffix='.txt', delete=False) as temp_file:
            temp_file.write(text)
            logger.debug(f"Extracted text from image: {text[:100]}...")
            logger.info(f"Created temporary text file: {temp_file.name}")
            return temp_file.name
    except Exception as e:
        logger.error(f"Image OCR failed for {image_path}: {str(e)}")
        raise Exception(f"Image OCR failed: {str(e)}")



def is_rag_compatible(md_text: str, return_details: bool = False) -> tuple[bool, dict] | bool:
    logger.debug("Checking RAG compatibility")
    lines = md_text.splitlines()
    stripped_lines = [l.strip() for l in lines]

    # Add debug logging for step detection
    for line in stripped_lines:
        if line:  # Only check non-empty lines
            step_match = re.match(r"^(Step\s*\d+[:.)]?|\d+[\.\)]|-|\•)", line, re.IGNORECASE)
            logger.debug(f"Line: '{line}' | Step match: {bool(step_match)}")
            if step_match:
                logger.debug(f"Matched as step: {step_match.group()}")
    
    heading_pattern = re.compile(r"^(#+|[A-Z][a-zA-Z ]{3,40}:)")
    has_heading = any(heading_pattern.match(l.strip()) for l in lines)
    # has_heading = any(l.startswith("#") for l in lines)
    has_steps = any(re.match(r"^(Step\s*\d+[:.)]?|\d+[\.\)]|-|\•)", l, re.IGNORECASE) for l in stripped_lines)
    long_enough = len(md_text.strip()) > 100
    step_count = sum(1 for l in stripped_lines if re.match(r"^(Step\s*\d+[:.)]?|\d+[\.\)]|-|\•)", l, re.IGNORECASE))
    heading_count = sum(1 for l in stripped_lines if l.startswith("#"))

    # Score-based structure
    score = int(has_heading) + int(has_steps) + int(long_enough)
    is_compatible = score >= 1

    details = {
        "is_compatible": is_compatible,
        "length": len(md_text.strip()),
        "char_count": len(md_text),
        "line_count": len(lines),
        "step_count": step_count,
        "heading_count": heading_count,
        "has_headings": has_heading,
        "has_steps": has_steps,
        "structure_score": score,
    }

    logger.info(f"RAG compatibility check details: {details}")
    return (is_compatible, details) if return_details else is_compatible






