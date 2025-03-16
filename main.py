import zipfile
import xml.etree.ElementTree as ET

# Extract comments from the .docx file (which is a zipped XML file)
docx_zip = zipfile.ZipFile("./novel.docx", "r")

# Comments are stored in word/comments.xml
if "word/comments.xml" in docx_zip.namelist():
    comments_xml = docx_zip.read("word/comments.xml")
    root = ET.fromstring(comments_xml)

    # Extract comment text
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    extracted_comments = []
    for comment in root.findall("w:comment", namespaces):
        author = comment.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author", "Unknown")
        text = "".join(comment.itertext()).strip()
        extracted_comments.append(f"Comment by {author}: {text}")

    docx_zip.close()
else:
    extracted_comments = ["No comments found in the document."]

# Show extracted comments (preview)
print(extracted_comments[:10])