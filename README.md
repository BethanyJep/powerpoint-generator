# PowerPoint Generator 

A simple tool to convert Word documents into PowerPoint presentations using only Python without AI.

## Features

- Converts Word documents (`.docx`) to PowerPoint decks (`.pptx`)
- Extracts headings and paragraphs to slide titles and content
- No external AI or cloud dependencies

## Prerequisites

- Python 3.8 or higher
- python-docx
- python-pptx

## Installation

1. Clone this repository
2. Install dependencies:

```bash
pip install python-docx python-pptx
```

## Usage

### Command Line

Convert a Word document to PowerPoint:

```bash
python word_to_ppt.py input_document.docx --output output_presentation.pptx
```

Or use the demo script:

```bash
python demo.py input_document.docx --output output_presentation.pptx
```

- `input_document.docx`: Path to your source Word file
- `output_presentation.pptx`: Optional, path for the generated PowerPoint

## How It Works

1. Reads the Word file and parses headings and paragraphs.
2. Each heading creates a new slide with the heading as the title.
3. Paragraphs under each heading become bullet points on the slide.
4. Saves the resulting PowerPoint file.

## Example

```bash
python word_to_ppt.py develop-ai-agent-with-semantic-kernel.docx
```

This produces `develop-ai-agent-with-semantic-kernel_presentation.pptx` in the same folder.