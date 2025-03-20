import os
import zipfile
from xml.etree import ElementTree as ET

# Define XML namespaces for WordprocessingML
NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def parse_comments(comments_file):
    """Parse comments.xml and return a dictionary of comments."""
    comments_tree = ET.parse(comments_file)
    comments_root = comments_tree.getroot()
    return {
        comment.attrib[f'{{{NAMESPACES["w"]}}}id']: {
            'content': ''.join(comment.itertext()).strip(),
            'replies': [],  # Placeholder for replies if needed
            'selected_text': ''  # Initialize selected_text
        }
        for comment in comments_root.findall('w:comment', NAMESPACES)
    }

def extract_selected_text(document_root: ET.Element, comments):
    """Extract selected text for comments by tracking comment ranges."""
    current_comment_ids = []  # Track multiple active comment IDs
    current_text = []  # Collect text within a comment range
    for elem in document_root.iter():
        if elem.tag.endswith("commentRangeStart"):
            # Start tracking a new comment range
            comment_id = elem.attrib.get(f'{{{NAMESPACES["w"]}}}id')
            if comment_id:
                current_comment_ids.append(comment_id)
            current_text = []  # Reset text collection
        elif elem.tag.endswith("commentRangeEnd"):
            # End tracking a comment range and assign selected text
            comment_id = elem.attrib.get(f'{{{NAMESPACES["w"]}}}id')
            if comment_id in current_comment_ids:
                for active_comment_id in current_comment_ids:
                    if active_comment_id in comments:
                        comments[active_comment_id]['selected_text'] = " ".join(current_text).strip()
                current_comment_ids.remove(comment_id)
        elif elem.tag.endswith("t"):  # Regular text
            # Collect text within the current comment range
            current_text.append(elem.text or "")

def extract_paragraph_data(paragraph: ET.Element, comments):
    """Extract data for a single paragraph, including comments and suggestions."""
    paragraph_text = ''.join(paragraph.itertext()).strip()
    paragraph_data = {
        'content': paragraph_text,  # Full text of the paragraph
        'comments': [],  # Comments associated with the paragraph
        'suggestions': []  # Suggestions (insertions/deletions) in the paragraph
    }

    # Handle comments associated with the paragraph
    comment_references = paragraph.findall('.//w:commentRangeStart', NAMESPACES)
    if comment_references:
        for comment_ref in comment_references:
            comment_id = comment_ref.attrib[f'{{{NAMESPACES["w"]}}}id']
            if comment_id in comments:
                paragraph_data['comments'].append({
                    'content': comments[comment_id]['content'],  # Comment content
                    'selected_text': comments[comment_id].get('selected_text', ''),  # Text the comment refers to
                    'replies': comments[comment_id]['replies']  # Replies to the comment
                })

    # Handle insertions (suggested additions) in the paragraph
    insertions = paragraph.findall('.//w:ins', NAMESPACES)
    for insertion in insertions:
        inserted_text = ''.join(insertion.itertext()).strip()
        paragraph_data['suggestions'].append({
            'insertion': inserted_text,  # Text that was inserted
            'deletion': None,  # No deletion for this suggestion
            'replies': []  # Placeholder for replies if needed
        })

    # Handle deletions (suggested removals) in the paragraph
    deletions = paragraph.findall('.//w:del', NAMESPACES)
    for deletion in deletions:
        deleted_text = ''.join(deletion.itertext()).strip()
        paragraph_data['suggestions'].append({
            'insertion': None,  # No insertion for this suggestion
            'deletion': deleted_text,  # Text that was deleted
            'replies': []  # Placeholder for replies if needed
        })

    return paragraph_data

def extract_paragraphs(docx_path):
    """Extract paragraphs and their associated data from a .docx file."""
    paragraphs = []  # List to store paragraph data

    # Unzip the .docx file to access its contents
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # Check if both comments.xml and document.xml exist in the .docx archive
        if comments_xml_path in docx_zip.namelist() and document_xml_path in docx_zip.namelist():
            # Parse comments.xml to extract comment data
            with docx_zip.open(comments_xml_path) as comments_file:
                comments = parse_comments(comments_file)
            
            # Parse document.xml to extract paragraph and comment references
            with docx_zip.open(document_xml_path) as document_file:
                document_tree = ET.parse(document_file)
                document_root = document_tree.getroot()

                # Extract selected_text for comments
                extract_selected_text(document_root, comments)

                # Iterate through paragraphs and collect structured data
                for paragraph in document_root.findall('.//w:p', NAMESPACES):
                    paragraph_data = extract_paragraph_data(paragraph, comments)
                    # Skip empty paragraphs
                    if paragraph_data['content']:
                        paragraphs.append(paragraph_data)
        else:
            # Print an error message if required XML files are missing
            print("Required XML files (comments.xml or document.xml) not found in the document.")

    return paragraphs

# Define the paths to the .docx file and relevant XML files
input_folder = os.path.join(os.getcwd(), "input")
docx_file = next((f for f in os.listdir(input_folder) if f.endswith(".docx")), None)

if not docx_file:
    # Raise an error if no .docx file is found in the input folder
    raise FileNotFoundError("No .docx file found in the input folder.")

docx_path = os.path.join(input_folder, docx_file)
comments_xml_path = "word/comments.xml"
document_xml_path = "word/document.xml"

# Extract paragraphs and their associated data from the .docx file
paragraphs = extract_paragraphs(docx_path)

# Print the structured data for each paragraph
for paragraph in paragraphs:
    print(paragraph)