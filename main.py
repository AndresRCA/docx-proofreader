import os
import zipfile
import xml.etree.ElementTree as ET

# Formatting functions
def format_insertion_text(text):
  return f"**{text}**"

def format_deletion_text(text):
  return "".join(f"{char}̶" for char in text)  # Strikethrough formatting

def extract_paragraphs(docx_path):
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

# Define the paths to the .docx file and relevant XML files
input_folder = os.path.join(os.getcwd(), "input")
docx_file = next((f for f in os.listdir(input_folder) if f.endswith(".docx")), None)

if not docx_file:
    # Raise an error if no .docx file is found in the input folder
    raise FileNotFoundError("No .docx file found in the input folder.")

docx_path = os.path.join(input_folder, docx_file)
paragraphs = extract_paragraphs(docx_path)
for paragraph in paragraphs:
    print(paragraph)