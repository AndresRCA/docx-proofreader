import zipfile
import xml.etree.ElementTree as ET

# Example <w:p> element with nested comments
xml_data = """
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:r>
        <w:t>Some text </w:t>
    </w:r>
    <w:commentRangeStart w:id="1"/>
    <w:sdt>
        <w:sdtContent>
            <w:commentRangeStart w:id="2"/>
            <w:r>
               <w:t>Nested comment text </w:t>
            </w:r>
        </w:sdtContent>
    </w:sdt>
    <w:r>
        <w:t>More text </w:t>
    </w:r>
    <w:ins>
        <w:r><w:t>insertion</w:t></w:r>
    </w:ins>  
    <w:del>
        <w:r><w:delText>deletion</w:delText></w:r>
    </w:del>    
    <w:commentRangeEnd w:id="1"/>
    <w:commentRangeEnd w:id="2"/>
</w:p>
"""
# XML namespace for WordprocessingML
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml"
}

# Formatting functions
def format_insertion_text(text):
    return f"**{text}**"  # Bold formatting for insertions

def format_deletion_text(text):
    return f"--{text}--"  # Double dash formatting for deletions

def extract_paragraphs(docx_path):
  with zipfile.ZipFile(docx_path, "r") as docx:
    with docx.open("word/document.xml") as xml_file:
      tree = ET.parse(xml_file)
      root = tree.getroot()

      paragraphs = []
      for p in root.findall(".//w:p", NAMESPACES):
        paragraph_id = p.attrib[f'{{{NAMESPACES["w14"]}}}paraId']
        paragraph_text = get_paragraph_text(p)
        if (paragraph_text):
          paragraphs.append({"id": paragraph_id, "content": paragraph_text})
      return paragraphs

def get_plain_text(element: ET.Element):
    """
    Recursively extracts all text from <w:t> and <w:delText> nodes in the given element.
    """
    parts = []
    # We use .iter() to walk the entire subtree.
    for child in element.iter():
        tag = child.tag.split("}")[-1]
        if tag in ("t", "delText"):
            parts.append(child.text or "")
    return "".join(parts)

# Function to recursively extract text and apply formatting
def get_paragraph_text(element):
    """
    Recursively process the element, but when encountering an <w:ins> or <w:del>
    element, first collect all its text (using get_plain_text) then apply the formatting
    only once.
    """
    text_parts = []
    
    for child in element:
        tag = child.tag.split("}")[-1]
        
        if tag == "ins":
            # For an insertion block, get all text inside it and apply formatting once.
            ins_text = get_plain_text(child)
            if ins_text:  # Only add formatted text if there's something to format
                text_parts.append(format_insertion_text(ins_text))

        elif tag == "del":
            # For a deletion block, do the same.
            del_text = get_plain_text(child)
            if del_text:
                text_parts.append(format_deletion_text(del_text))
        
        # For a run, we want to get its text.
        elif tag == "r":
            # There might be multiple text parts inside a run.
            run_text = get_plain_text(child)
            text_parts.append(run_text)

        # Otherwise, recursively process the child element
        else:
            text_parts.append(get_paragraph_text(child))

    full_paragraph = "".join(text_parts)
    return full_paragraph

# Parse the XML
root = ET.fromstring(xml_data)

# Extract the text
full_paragraph = get_paragraph_text(root)
print(full_paragraph)

paragraphs = extract_paragraphs("./input/test1.docx")
print("\n".join(map(str, paragraphs)))