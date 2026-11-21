import re
import zlib

def extract_pdf_text_pure_python(pdf_bytes):
    text_chunks = []
    # Find all streams in PDF
    stream_matches = re.finditer(b'stream\r?\n(.*?)\r?\nendstream', pdf_bytes, re.DOTALL)
    for match in stream_matches:
        raw_stream = match.group(1)
        decompressed = None
        try:
            decompressed = zlib.decompress(raw_stream)
        except Exception:
            try:
                decompressed = zlib.decompress(raw_stream, -zlib.MAX_WBITS)
            except Exception:
                decompressed = raw_stream

        if decompressed:
            # Extract strings in parentheses (text) from PDF operators: [(...) ...] TJ or (...) Tj
            for m in re.finditer(rb'\((.*?)\)\s*T[jJ]', decompressed):
                try:
                    text_chunks.append(m.group(1).decode('latin1', errors='ignore'))
                except Exception:
                    pass
    return " ".join(text_chunks)

print("PDF extractor helper defined successfully")
