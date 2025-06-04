# aimath2docx

A converter from Markdown files with LaTeX formulas to Word (.docx) with full OMML formula support.

## About

This tool was specifically designed for converting GPT (ChatGPT or other LLM) answers containing mathematical LaTeX formulas, as well as scientific Markdown documents, into properly formatted `.docx` files. It simplifies the workflow from AI-generated content or technical notes to clean, ready-to-use Word documents with full equation support.

## Installation

For the converter to work correctly, before installation you need to remove mathml2omml if it is installed:

    pip uninstall mathml2omml

To install the converter and GUI directly from GitHub:

    pip install git+https://github.com/AlloteSoftware/aimath2docx.git

## Usage: Command-line tool (aimath2docx.py)

This script converts a Markdown file with LaTeX formulas and tables into a Word (.docx) document using Python.

Features:
- Converts LaTeX formulas (inline, block, and math code blocks) to Word OMML format
- Supports Markdown formatting: bold, italic, strikeout, headers, lists
- Supports Markdown tables with formulas
- Automatically adjusts and fixes LaTeX for compatibility

Run from command line:

    aimath2docx input.md output.docx

## Usage: GUI application (aimath2docx-gui)

This is a graphical user interface for the converter built with customtkinter.

Features:
- File picker for Markdown input
- Output path selection with default filename
- Clean centered layout with status message

Run the GUI:

    aimath2docx-gui

## License

This project is licensed under the Apache License 2.0.

You are free to use, modify, and distribute this software, including for commercial purposes, as long as you retain the copyright:

Copyright 2025 AlloteSoftware
