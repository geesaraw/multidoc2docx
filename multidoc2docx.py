import os
import win32com.client as win32

def from_doc_to_docx(path_to_doc, path_to_docx):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path_to_doc)
    doc.SaveAs(path_to_docx, 12)
    doc.Close()
    word.Quit()

def convert_multiple_docs_to_docx(input_folder, output_folder):
    for root, _, files in os.walk(input_folder):
        for filename in files:
            if filename.lower().endswith('.doc'):
                path_to_doc = os.path.join(root, filename)
                path_to_docx = os.path.join(output_folder, os.path.relpath(path_to_doc, input_folder))
                path_to_docx = os.path.splitext(path_to_docx)[0] + ".docx"
                if not filename.startswith('~$'):
                    if not os.path.exists(os.path.dirname(path_to_docx)):
                        os.makedirs(os.path.dirname(path_to_docx))
                    from_doc_to_docx(path_to_doc, path_to_docx)
                    print(f"Converted: {path_to_doc} to {path_to_docx}")
                else:
                    print(f"Skipped: {path_to_doc}")

if __name__ == "__main__":
    input_folder = r'C:\path\to\doc\folder'
    output_folder = r'C:\path\to\output\folder'

    convert_multiple_docs_to_docx(input_folder, output_folder)
