import os
import zipfile
from collections import defaultdict
import xml.etree.ElementTree as ET

# Formatting functions
def format_insertion_text(text):
  return f"**{text}**"

def format_deletion_text(text):
  return "".join(f"{char}̶" for char in text)  # Strikethrough formatting

def extract_paragraphs(docx_path):
  """
    Returns:
      list[dict]: A list of paragraphs, each represented as:
        - content (str): Paragraph text.
        - comments (list[dict]): Associated comments, each with:
            - id (str): Comment ID.
            - anchor (str): Highlighted text.
  """
  with zipfile.ZipFile(docx_path, "r") as docx:
    with docx.open("word/document.xml") as xml_file:
      tree = ET.parse(xml_file)
      root = tree.getroot()

      # Namespace handling
      ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
      paragraphs = []

      for paragraph in root.findall(".//w:p", ns):
        # Skip paragraphs without comments, insertions, or deletions
        if not any(paragraph.findall(f".//w:{tag}", ns) for tag in ["commentRangeStart", "ins", "del"]):
          continue

        paragraph_portions = []
        comments = []
        current_active_ids = []

        for elem in paragraph:
          text = ''
          
          # Regular text
          if elem.tag == f"{{{ns['w']}}}r":
            text_elem = elem.find("w:t", ns)
            if text_elem is None: continue

            text = text_elem.text

            for active_id in current_active_ids:
              comment = next(
                (c for c in comments if c["id"] == active_id), None
              )
              if comment:
                comment["text_portions"].append(text)

          # Insertions
          elif elem.tag == f"{{{ns['w']}}}ins":
            text_elem = elem.find(".//w:t", ns)
            if text_elem is None: continue

            text = format_insertion_text(text_elem.text)
            
            # Handle replies within insertions/deletions
            for child in elem.iter():
              if child.tag == f"{{{ns['w']}}}commentRangeStart":
                  comment_id = child.attrib[f'{{{ns["w"]}}}id']
                  if comment_id:
                    comments.append({"id": comment_id, "text_portions": []})
                    current_active_ids.append(comment_id)

              elif child.tag == f"{{{ns['w']}}}commentRangeEnd":
                comment_id = child.attrib[f'{{{ns["w"]}}}id']
                comment = next((c for c in comments if c["id"] == comment_id), None)
                if comment:
                  comment["anchor"] = "".join(comment["text_portions"])
                  del comment["text_portions"]
                if comment_id in current_active_ids:
                  current_active_ids.remove(comment_id)

            # check if there's an active id for a comment
            for active_id in current_active_ids:
              comment = next(
                (c for c in comments if c["id"] == active_id), None
              )
              if comment:
                comment["text_portions"].append(text)

          # Deletions
          elif elem.tag == f"{{{ns['w']}}}del":
            text_elem = elem.find(".//w:delText", ns)
            if text_elem is None: continue

            text = format_deletion_text(text_elem.text)
            
            # Handle replies within insertions/deletions
            for child in elem.iter():
              if child.tag == f"{{{ns['w']}}}commentRangeStart":
                comment_id = child.attrib[f'{{{ns["w"]}}}id']
                if comment_id:
                  comments.append({"id": comment_id, "text_portions": []})
                  current_active_ids.append(comment_id)

              elif child.tag == f"{{{ns['w']}}}commentRangeEnd":
                comment_id = child.attrib[f'{{{ns["w"]}}}id']
                comment = next((c for c in comments if c["id"] == comment_id), None)
                if comment:
                  comment["text_portions"].append(text)
                  comment["anchor"] = "".join(comment["text_portions"])
                  del comment["text_portions"]
                if comment_id in current_active_ids:
                  current_active_ids.remove(comment_id)

            # check if there's an active id for a comment
            for active_id in current_active_ids:
              comment = next(
                (c for c in comments if c["id"] == active_id), None
              )
              if comment:
                comment["text_portions"].append(text)

          # Comment starts
          elif elem.tag == f"{{{ns['w']}}}commentRangeStart":
            comment_id = elem.attrib[f'{{{ns["w"]}}}id']
            if comment_id:
              comments.append({"id": comment_id, "text_portions": []})
              current_active_ids.append(comment_id)

          # Comment endings
          elif elem.tag == f"{{{ns['w']}}}commentRangeEnd":
            comment_id = elem.attrib[f'{{{ns["w"]}}}id']
            comment = next(
              (c for c in comments if c["id"] == comment_id), None
            )
            if comment:
              comment["anchor"] = "".join(comment["text_portions"])
              del comment["text_portions"]
            if comment_id in current_active_ids:
              current_active_ids.remove(comment_id)
          
          paragraph_portions.append(text)

        full_paragraph_text = "".join(paragraph_portions)
        paragraph_dict = {"content": full_paragraph_text, "comments": comments}
        paragraphs.append(paragraph_dict)

      return paragraphs

def extract_comments(comments):
  """
    Extracts comment content from comments.xml and associates it with the respective comment IDs.
    Updates the content for top-level comments and their replies.
    Args:
      comments (list[dict]): A list of comments, each represented as:
        - id (str): Comment ID.
        - anchor (str): Highlighted text.
        - replies (list[dict]): Associated replies, each with:
            - id (str): Comment ID.
    Returns:
      list[dict]: Updated comments with content for each ID.
  """
  # Open the .docx file and access comments.xml
  with zipfile.ZipFile(docx_path, "r") as docx:
    with docx.open("word/comments.xml") as xml_file:
      tree = ET.parse(xml_file)
      root = tree.getroot()

      # Namespace handling
      ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

      # Update the content for each comment and its replies
      for comment in comments:
        # Update top-level comment content
        comment_id = comment["id"]
        comment_elem = root.find(f"./w:comment[@w:id='{comment_id}']", ns)
        if comment_elem is None:
          raise ValueError("Couldn't find comment with id=" + comment_id)
        comment["content"] = comment_elem.find("./w:p/w:r/w:t", ns).text

        # Update replies content
        for reply in comment["replies"]:
          reply_id = reply["id"]
          reply_elem = root.find(f"./w:comment[@w:id='{reply_id}']", ns)
          if reply_elem is None:
            raise ValueError("Couldn't find reply with id=" + reply_id)
          
          reply["content"] = reply_elem.find("./w:p/w:r/w:t", ns).text

  return comments

def sort_comments(comments):
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
  for comment in comments:
    grouped[comment['anchor']].append(comment)

  # Process grouped comments
  sorted_comments = []
  for group in grouped.values():
    main_comment = group[0].copy() # First comment is the main one
    main_comment['replies'] = []
    if len(group) > 1:
      main_comment['replies'] = [{'id': c['id']} for c in group[1:]] # Subsequent replies
    sorted_comments.append(main_comment)

  return sorted_comments

# Define the paths to the .docx file and relevant XML files
input_folder = os.path.join(os.getcwd(), "input")
docx_file = next((f for f in os.listdir(input_folder) if f.endswith(".docx")), None)

if not docx_file:
    # Raise an error if no .docx file is found in the input folder
    raise FileNotFoundError("No .docx file found in the input folder.")

docx_path = os.path.join(input_folder, docx_file)
paragraphs = extract_paragraphs(docx_path)
for paragraph in paragraphs:
    # print(paragraph)

    paragraph['comments'] = sort_comments(paragraph['comments'])
    # print(paragraph['comments'])

    paragraph['comments'] = extract_comments(paragraph['comments'])
    print(paragraph['comments'])
