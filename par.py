import zipfile
from collections import defaultdict
import xml.etree.ElementTree as ET

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

def extract_paragraphs(root: ET.Element):
  paragraphs = []
  for p in root.findall(".//w:p", NAMESPACES):
    paragraph_id = p.attrib[f'{{{NAMESPACES["w14"]}}}paraId']
    paragraph_text = get_paragraph_text(p)
    if (paragraph_text):
      paragraphs.append({"id": paragraph_id, "content": paragraph_text})
  return paragraphs

# Function to recursively extract text and apply formatting
def get_paragraph_text(element: ET.Element):
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

# Recursive function to extract text within comment ranges while handling nesting properly
def get_anchors(parent: ET.Element, active_comments: list, comments: dict, ancestors: list[ET.Element]):
  """
  Returns:
    dict[str, dict]: A dictionary where:
      - Keys (str) represent comment IDs.
      - Values (dict) contain:
        - anchor (str): The text associated with the comment.
  """
  for elem in parent:
    tag_name = elem.tag.split("}")[-1]  # Remove namespace

    # Push current element to ancestors stack
    ancestors.append(elem)

    # Start tracking a new comment range
    if tag_name == "commentRangeStart":
      comment_id = elem.attrib[f'{{{NAMESPACES["w"]}}}id']
      active_comments.append(comment_id)

    # Collect text for all active comments
    elif active_comments and tag_name in ["t", "delText"]:
      text = elem.text or ""  # Ensure text is not None

      # Check grandparent for <w:ins> or <w:del>
      grandparent = ancestors[-3] if len(ancestors) >= 3 else None
      if grandparent is not None:
        grandparent_tag = grandparent.tag.split("}")[-1]
        if grandparent_tag == "ins":
          text = format_insertion_text(text)
        elif grandparent_tag == "del":
          text = format_deletion_text(text)

      # Assign text to ALL active comments
      for comment_id in active_comments:
        comments[comment_id]["anchor"] += text

    # Recursively process child elements
    get_anchors(elem, active_comments, comments, ancestors)

    # Stop tracking when reaching a matching comment end
    if tag_name == "commentRangeEnd":
      comment_id = elem.attrib[f'{{{NAMESPACES["w"]}}}id']
      if comment_id in active_comments:
        active_comments.remove(comment_id)  # Remove the completed comment range
        
    # Pop current element from ancestors stack
    ancestors.pop()

  return comments  # Return updated comments dictionary

def sort_replies(comments):
  """
    Adds a reply list to the list of comments if there are comments that share the same anchor value.
    Returns:
      list[dict]: A list of paragraphs, each represented as:
        - id (str): Comment ID.
        - replies (list[dict]): Associated comments, each with:
          - id (str): Comment ID.
  """
  # Group comments by anchor
  grouped = defaultdict(list)
  for comment_id, comment in comments.items():
    grouped[comment['anchor']].append({"id": comment_id, **comment}) # grouped = { "{anchor_value}": [{ "id": str, "anchor": str }] }

  # Process grouped comments
  sorted_comments = []
  for group in grouped.values():
    main_comment = group[0].copy() # First comment is the main one
    main_comment['replies'] = []
    if len(group) > 1:
      main_comment['replies'] = [{'id': c['id']} for c in group[1:]] # Subsequent replies
    sorted_comments.append(main_comment)

  return sorted_comments

def get_comment_content(comments_root: ET.Element, comments: dict):
  # Update the content for each comment and its replies
  for comment in comments:
    # Update top-level comment content
    comment_id = comment["id"]
    comment_elem = comments_root.find(f"./w:comment[@w:id='{comment_id}']", NAMESPACES)
    if comment_elem is None:
      raise ValueError("Couldn't find comment with id=" + comment_id)
    comment["content"] = comment_elem.find(".//w:t", NAMESPACES).text

    # Update replies content
    for reply in comment["replies"]:
      reply_id = reply["id"]
      reply_elem = comments_root.find(f"./w:comment[@w:id='{reply_id}']", NAMESPACES)
      if reply_elem is None:
        raise ValueError("Couldn't find reply with id=" + reply_id)
      
      reply["content"] = reply_elem.find(".//w:t", NAMESPACES).text

  return comments

def extract_comments_anchor(document_root: ET.Element, comments_root: ET.Element, paragraph_id):
  """Extract the anchor for comments found in the paragraph identified with `paragraph_id`"""
  paragraph_root = document_root.find(f".//*[@w14:paraId='{paragraph_id}']", NAMESPACES)
  
  # Find all <w:commentRangeStart> and <w:commentRangeEnd> inside <w:p>
  comment_starts = paragraph_root.findall(".//w:commentRangeStart", NAMESPACES)
  comment_ends = paragraph_root.findall(".//w:commentRangeEnd", NAMESPACES)

  # Create a mapping of comment IDs found inside
  comments = {} # { '{id}': { 'anchor': str } }

  for start in comment_starts:
    comment_id = start.attrib[f'{{{NAMESPACES["w"]}}}id']
    comments[comment_id] = {"anchor": ""}

  # Find comment IDs that don't have a matching start
  unmatched_comment_ids = []

  for end in comment_ends:
    comment_id = end.attrib[f'{{{NAMESPACES["w"]}}}id']
    if comment_id not in comments:
      print(f"Couldn't find end of comment with ID={comment_id} in the same paragraph.")
      print("Ignoring comment from list...")
      unmatched_comment_ids.append(comment_id)
  # Remove unmatched comment IDs
  for comment_id in unmatched_comment_ids:
    comments.pop(comment_id, None)

  # Proceed with retrieving the commented sections of the paragraph
  comments = get_anchors(paragraph_root, [], comments, []) # { "{comment_id}": {anchor_text: str} }
  comments = sort_replies(comments) # { "id": str, "anchor": str, "replies": list[{ "id": str }] }
  comments = get_comment_content(comments_root, comments)
  print(comments)

  return comments

docx_path = "./input/test.docx"

# Parse the XML files we'll be using
document_root = None
comments_root = None
with zipfile.ZipFile(docx_path, "r") as docx:
  with docx.open("word/document.xml") as document_xml:
    document_tree = ET.parse(document_xml)
    document_root = document_tree.getroot()
  with docx.open("word/comments.xml") as comments_xml:
    comments_tree = ET.parse(comments_xml)
    comments_root = comments_tree.getroot()

paragraphs = extract_paragraphs(document_root)
print("\n".join(map(str, paragraphs)))
print("\n\n")

for paragraph in paragraphs:
  paragraph["comments"] = extract_comments_anchor(document_root, comments_root, paragraph["id"])