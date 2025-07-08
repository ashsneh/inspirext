from langchain.text_splitter import MarkdownHeaderTextSplitter

def chunk_by_markdown_headers(markdown_text: str) -> list:
    splitter = MarkdownHeaderTextSplitter(
        headers_to_split_on=[
            ("#", "h1"),
            ("##", "h2"),
            ("###", "h3")
        ]
    )
    return splitter.split_text(markdown_text)



#alternative code 


# from langchain.text_splitter import MarkdownHeaderTextSplitter, RecursiveCharacterTextSplitter

# def chunk_by_markdown_headers(markdown_text: str, chunk_size: int = 1000, chunk_overlap: int = 100) -> list:
#     # Step 1: Split by markdown headers
#     header_splitter = MarkdownHeaderTextSplitter(
#         headers_to_split_on=[
#             ("#", "h1"),
#             ("##", "h2"),
#             ("###", "h3")
#         ]
#     )
#     sections = header_splitter.split_text(markdown_text)

#     # Step 2: For each section, apply recursive character splitter
#     final_chunks = []
#     recursive_splitter = RecursiveCharacterTextSplitter(
#         chunk_size=chunk_size,
#         chunk_overlap=chunk_overlap
#     )

#     for section in sections:
#         chunks = recursive_splitter.split_documents([section])
#         final_chunks.extend(chunks)

#     return final_chunks
