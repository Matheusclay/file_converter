#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from pathlib import Path

if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

import tkinter as tk
from tkinter import filedialog, messagebox

from pdf2docx import Converter

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None

try:
    import docx2txt
    from docx import Document
    import comtypes.client
    comtypes_available = True
except ImportError:
    docx2txt = None
    comtypes_available = False

def safe_print(text):
    try:
        print(text)
    except UnicodeEncodeError:
        clean_text = text.encode('ascii', 'ignore').decode('ascii')
        print(clean_text)

def safe_basename(path):
    try:
        return os.path.basename(path)
    except:
        return "file"

def select_input_file():
    root = tk.Tk()
    root.withdraw()
    
    file_types = [
        ("Word Documents", "*.doc *.docx"),
        ("PDF Files", "*.pdf"),
        ("DOC Files", "*.doc"),
        ("DOCX Files", "*.docx"),
        ("All Files", "*.*")
    ]
    
    selected_file = filedialog.askopenfilename(
        title="Select file to convert",
        filetypes=file_types,
        initialdir=os.path.expanduser("~/Documents")
    )
    
    root.destroy()
    return selected_file if selected_file else None

def select_output_file(input_file):
    root = tk.Tk()
    root.withdraw()
    
    input_ext = Path(input_file).suffix.lower()
    base_name = Path(input_file).stem
    
    if input_ext == '.pdf':
        output_ext = '.docx'
        file_types = [("DOCX Files", "*.docx"), ("All Files", "*.*")]
    elif input_ext in ['.doc', '.docx']:
        output_ext = '.pdf'
        file_types = [("PDF Files", "*.pdf"), ("All Files", "*.*")]
    else:
        output_ext = '.pdf'
        file_types = [("All Files", "*.*")]
    
    suggested_name = f"{base_name}{output_ext}"
    
    output_file = filedialog.asksaveasfilename(
        title="Save converted file as:",
        defaultextension=output_ext,
        filetypes=file_types,
        initialfile=suggested_name,
        initialdir=os.path.dirname(input_file)
    )
    
    root.destroy()
    return output_file if output_file else None

def convert_pdf_to_docx(pdf_file, docx_file=None):
    if not os.path.exists(pdf_file):
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")
    
    if docx_file is None:
        base_name = os.path.splitext(pdf_file)[0]
        docx_file = f"{base_name}.docx"
    
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()
    
    safe_print(f"PDF -> DOCX: {safe_basename(pdf_file)} -> {safe_basename(docx_file)}")
    return docx_file

def convert_docx_to_pdf(docx_file, pdf_file=None):
    if not os.path.exists(docx_file):
        raise FileNotFoundError(f"DOCX file not found: {docx_file}")
    
    if docx2pdf_convert is None:
        raise ImportError("docx2pdf library not available. Run: pip install docx2pdf")
    
    if pdf_file is None:
        base_name = os.path.splitext(docx_file)[0]
        pdf_file = f"{base_name}.pdf"
    
    docx2pdf_convert(docx_file, pdf_file)
    safe_print(f"DOCX -> PDF: {safe_basename(docx_file)} -> {safe_basename(pdf_file)}")
    return pdf_file

def convert_doc_to_docx_with_word(doc_file, docx_file=None):
    if not comtypes_available:
        raise ImportError("comtypes library not available. Run: pip install comtypes")
    
    if docx_file is None:
        base_name = os.path.splitext(doc_file)[0]
        docx_file = f"{base_name}.docx"
    
    doc_file = os.path.abspath(doc_file)
    docx_file = os.path.abspath(docx_file)
    
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    try:
        doc = word.Documents.Open(doc_file)
        doc.SaveAs2(docx_file, FileFormat=16)
        doc.Close()
        word.Quit()
        
        safe_print(f"DOC -> DOCX: {safe_basename(doc_file)} -> {safe_basename(docx_file)}")
        return docx_file
    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        raise e

def convert_doc_to_docx(doc_file, docx_file=None):
    if not os.path.exists(doc_file):
        raise FileNotFoundError(f"DOC file not found: {doc_file}")
    
    if docx_file is None:
        base_name = os.path.splitext(doc_file)[0]
        docx_file = f"{base_name}.docx"
    
    if comtypes_available and sys.platform.startswith('win'):
        try:
            return convert_doc_to_docx_with_word(doc_file, docx_file)
        except Exception as e:
            safe_print(f"Failed to use Microsoft Word: {str(e)}")
            safe_print("Trying alternative method...")
    
    if docx2txt is None:
        raise ImportError("Cannot convert .doc file. Install Microsoft Word or convert manually to .docx")
    
    safe_print("WARNING: DOC conversion may lose formatting. Microsoft Word recommended.")
    
    text = docx2txt.process(doc_file)
    doc = Document()
    
    for paragraph in text.split('\n'):
        if paragraph.strip():
            doc.add_paragraph(paragraph)
    
    doc.save(docx_file)
    safe_print(f"DOC -> DOCX: {safe_basename(doc_file)} -> {safe_basename(docx_file)}")
    return docx_file

def convert_doc_to_pdf(doc_file, pdf_file=None):
    if pdf_file is None:
        base_name = os.path.splitext(doc_file)[0]
        pdf_file = f"{base_name}.pdf"
    
    temp_docx = doc_file.replace('.doc', '_temp.docx')
    convert_doc_to_docx(doc_file, temp_docx)
    result = convert_docx_to_pdf(temp_docx, pdf_file)
    
    try:
        os.remove(temp_docx)
    except:
        pass
    
    safe_print(f"DOC -> PDF: {safe_basename(doc_file)} -> {safe_basename(pdf_file)}")
    return result

def detect_conversion_type(input_file):
    extension = Path(input_file).suffix.lower()
    
    if extension == '.pdf':
        return 'pdf_to_docx'
    elif extension == '.docx':
        return 'docx_to_pdf'
    elif extension == '.doc':
        return 'doc_to_pdf'
    else:
        raise ValueError(f"Unsupported extension: {extension}. Use .pdf, .docx or .doc")

def auto_convert(input_file, output_file=None):
    conversion_type = detect_conversion_type(input_file)
    
    if conversion_type == 'pdf_to_docx':
        return convert_pdf_to_docx(input_file, output_file)
    elif conversion_type == 'docx_to_pdf':
        return convert_docx_to_pdf(input_file, output_file)
    elif conversion_type == 'doc_to_pdf':
        return convert_doc_to_pdf(input_file, output_file)

def gui_mode():
    safe_print("File Converter - GUI Mode")
    safe_print("=" * 30)
    
    safe_print("Opening file selector...")
    input_file = select_input_file()
    
    if not input_file:
        safe_print("No file selected. Operation cancelled.")
        return
    
    safe_print(f"File selected: {safe_basename(input_file)}")
    
    try:
        conversion_type = detect_conversion_type(input_file)
        input_ext = Path(input_file).suffix.lower()
        
        if input_ext == '.pdf':
            conversion_text = "PDF -> DOCX"
        elif input_ext == '.docx':
            conversion_text = "DOCX -> PDF"
        elif input_ext == '.doc':
            conversion_text = "DOC -> PDF"
        
        safe_print(f"Conversion type: {conversion_text}")
        
    except Exception as e:
        safe_print(f"Error: {str(e)}")
        return
    
    safe_print("Select where to save the converted file...")
    output_file = select_output_file(input_file)
    
    if not output_file:
        safe_print("Output location not selected. Operation cancelled.")
        return
    
    safe_print(f"Save to: {safe_basename(output_file)}")
    
    try:
        safe_print("Converting file...")
        converted_file = auto_convert(input_file, output_file)
        
        safe_print("Conversion completed successfully!")
        safe_print(f"File saved: {converted_file}")
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Conversion Completed", 
            f"File converted successfully!\n\nSaved to:\n{converted_file}"
        )
        root.destroy()
        
    except Exception as e:
        safe_print(f"Error during conversion: {str(e)}")
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Conversion Error", 
            f"An error occurred during conversion:\n\n{str(e)}"
        )
        root.destroy()

def main():
    if len(sys.argv) < 2:
        safe_print("File Converter")
        safe_print("")
        safe_print("Usage modes:")
        safe_print("  python main.py                    # GUI mode")
        safe_print("  python main.py <file> [output]    # Command line mode")
        safe_print("")
        safe_print("Command line examples:")
        safe_print("  python main.py document.pdf       # PDF > DOCX")
        safe_print("  python main.py document.docx      # DOCX > PDF")
        safe_print("  python main.py document.doc       # DOC > PDF")
        safe_print("")
        
        response = input("Use GUI mode? (y/n): ").lower().strip()
        if response in ['y', 'yes', '']:
            gui_mode()
        else:
            safe_print("Use: python main.py <file>")
        
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        conversion_type = detect_conversion_type(input_file)
        safe_print(f"Input file: {input_file}")
        
        converted_file = auto_convert(input_file, output_file)
        
        safe_print("Conversion completed!") 
        safe_print(f"File saved to: {converted_file}")
    
    except Exception as e:
        safe_print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 