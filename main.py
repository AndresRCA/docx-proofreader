import os
import zipfile
from xml.etree import ElementTree as ET

def extract_paragraphs(docx_path):
    paragraphs = []  # List to store paragraph data

    # Unzip the .docx file to access its contents
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # Check if both comments.xml and document.xml exist in the .docx archive
        if comments_xml_path in docx_zip.namelist() and document_xml_path in docx_zip.namelist():
            # Parse comments.xml
            with docx_zip.open(comments_xml_path) as comments_file:
                comments_tree = ET.parse(comments_file)
                comments_root = comments_tree.getroot()
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                comments = {
                    comment.attrib[f'{{{ns["w"]}}}id']: {
                        'content': ''.join(comment.itertext()).strip(),
                        'replies': []  # Placeholder for replies if needed
                    }
                    for comment in comments_root.findall('w:comment', ns)
                }

            # Parse document.xml
            with docx_zip.open(document_xml_path) as document_file:
                document_tree = ET.parse(document_file)
                document_root = document_tree.getroot()

                # Iterate through paragraphs and collect data
                for paragraph in document_root.findall('.//w:p', ns):
                    paragraph_text = ''.join(paragraph.itertext()).strip()
                    paragraph_data = {
                        'content': paragraph_text,
                        'comments': [],
                        'suggestions': []
                    }

                    # Handle comments
                    comment_references = paragraph.findall('.//w:commentRangeStart', ns)
                    if comment_references:
                        for comment_ref in comment_references:
                            comment_id = comment_ref.attrib[f'{{{ns["w"]}}}id']
                            if comment_id in comments:
                                paragraph_data['comments'].append({
                                    'content': comments[comment_id]['content'],
                                    'replies': comments[comment_id]['replies']
                                })

                    # Handle insertions
                    insertions = paragraph.findall('.//w:ins', ns)
                    for insertion in insertions:
                        # author = insertion.attrib.get(f'{{{ns["w"]}}}author', 'Unknown')
                        # date = insertion.attrib.get(f'{{{ns["w"]}}}date', 'Unknown')
                        inserted_text = ''.join(insertion.itertext()).strip()
                        paragraph_data['suggestions'].append({
                            'insertion': inserted_text,
                            'deletion': None,
                            'replies': []  # Placeholder for replies if needed
                        })

                    # Handle deletions
                    deletions = paragraph.findall('.//w:del', ns)
                    for deletion in deletions:
                        # author = deletion.attrib.get(f'{{{ns["w"]}}}author', 'Unknown')
                        # date = deletion.attrib.get(f'{{{ns["w"]}}}date', 'Unknown')
                        deleted_text = ''.join(deletion.itertext()).strip()
                        paragraph_data['suggestions'].append({
                            'insertion': None,
                            'deletion': deleted_text,
                            'replies': []  # Placeholder for replies if needed
                        })

                    # Append the paragraph data to the list
                    paragraphs.append(paragraph_data)
        else:
            print("Required XML files (comments.xml or document.xml) not found in the document.")

    return paragraphs

# Define the paths to the .docx file and relevant XML files
# Locate the first .docx file in the input folder
input_folder = os.path.join(os.getcwd(), "input")
docx_file = next((f for f in os.listdir(input_folder) if f.endswith(".docx")), None)

if not docx_file:
    raise FileNotFoundError("No .docx file found in the input folder.")

docx_path = os.path.join(input_folder, docx_file)
# Relevant XML files that hold the information for the docx file after being zipped
comments_xml_path = "word/comments.xml"
document_xml_path = "word/document.xml"

paragraphs = extract_paragraphs(docx_path)

# Print the structured data
for paragraph in paragraphs:
    print(paragraph)