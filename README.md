# Proofreading Instructions Generation From DOCX Files

This script extracts paragraphs, insertions, deletions, and comments from a .docx file, formatting them into a structured .txt file that provides relevant context and comments to assist an AI agent in proofreading tasks. Great for working together with Google Docs.

## Features
- Extracts paragraphs from a `.docx` file.
- Detects and formats insertions (`**bold**`) and deletions (`--strikethrough--`).
- Captures comments and their replies, associating them with the relevant text.
- Includes contextual paragraphs for better understanding.
- Exports the extracted content into a readable `.txt` format.

## Usage
Run the script with the following arguments:
```sh
python main.py <docx_path> [-o <output_directory>] [-c <context_level>]
```

### Arguments:
- `<docx_path>`: Path to the input `.docx` file.
- `-o, --output_path` (optional): Directory to save the output file (default: current directory).
- `-c, --context_level` (optional): Number of surrounding paragraphs to include for context (default: `0`).

### Example:
```sh
python main.py document.docx -o output/ -c 1
```

This will extract and save the proofreading instructions in `output/proofread_instructions.txt` with one paragraph of context before and after each modified section.

## Output Format
```sh
python main.py document.docx -c 1
```
```
===
Current context:
{First paragraph. Single comment. In this section you can find some sample text and bd grmmar. End of paragraph.}
Second paragraph. Single suggestion. In this section **the things said in this text are**--the text is super-- casual and it doesn't read fluently.

Comment(s):
[In this section you can find some sample text and bd grmmar] -> check 1. double check. 
===
Current context:
First paragraph. Single comment. In this section you can find some sample text and bd grmmar. End of paragraph.
{Second paragraph. Single suggestion. In this section **the things said in this text are**--the text is super-- casual and it doesn't read fluently.}
Third paragraph. Comment+suggestion. In this section **the things said in this text are**--the text is super-- casual and it doesn't read fluently. End of paragraph.

Comment(s):
!NONE!
===
...
```

## Precautions
There are some things to take into account when using this tool:
* This scripts assumes a paragraph based scope, meaning any relevant edits and comments should be contained within their own paragraph. A comment that spans more than one paragraph will be ignored and not taken into account.