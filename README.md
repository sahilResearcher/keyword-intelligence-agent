# keyword-intelligence-agent
Search PDFs with keyword logic.  Auto-generates Excel + PowerPoint
A Python automation tool for researchers and analysts.

Search through multiple PDF and Word documents simultaneously
using a custom AND + OR keyword logic — then automatically
generate a formatted Excel report and PowerPoint deck
from every sentence that matches.

Built by a secondary research professional with no coding
background, using Python and Claude (Anthropic's AI).

---

## Real Output — What This Actually Produces

Searched across Google and Meta 10-K reports (2025)
using these keywords:

Must-have:  capital expenditure
Any-of:     AI, artificial intelligence

7 matching sentences found across 2 files.

Example match pulled automatically:

> "We continue to invest in capital expenditures as we
>  scale our technical infrastructure, in particular for
>  AI, to meet the demand of our users and enterprise
>  customers."
>
> Source: GOOG-10-K-2025.pdf — Page 29

Excel output: one row per match with File Name,
Page Number, and Matched Sentence columns.

PowerPoint output: one slide per match with dark blue
header showing source file, full sentence as body,
and page number in bottom right corner.

---

## Who This Is For

- Secondary researchers reading large volumes of reports
- Analysts extracting specific insights from PDFs
- Business professionals building evidence based slide decks
- Anyone spending hours on manual copy paste research work

---

## How the Keyword Logic Works

A sentence is only saved when BOTH conditions are true:

Must-have  →  ALL of these keywords must appear
Any-of     →  AT LEAST ONE of these must appear

Example:

Must-have:  capital expenditure
Any-of:     AI, artificial intelligence

Catches:  "Capital expenditure on AI rose 40% this year."
Skips:    "Capital expenditure on cloud infrastructure rose."
Skips:    "AI adoption is accelerating across the industry."

---

## Files in This Repository

keyword_agent_dynamic.py   — main script (run this)
updated_Keyword_Boolean.docx — plain English explanation
                               of every line of code
requirements.txt           — libraries you need to install

---

## Installation

Make sure Python is installed on your computer.
Then install the required libraries:

pip install pymupdf pandas openpyxl python-pptx python-docx

---

## How to Run

1. Download keyword_agent_dynamic.py
2. Open terminal or command prompt
3. Navigate to where you saved the file
4. Run:

python keyword_agent_dynamic.py

5. The script will ask you:

Enter folder path where your PDFs are:
> C:\Users\YourName\Desktop\Research

Enter MUST-HAVE keywords:
> capital expenditure

Enter ANY-OF keywords:
> AI, artificial intelligence

Start scanning? (yes / no):
> yes

6. Both output files appear in your folder automatically.

---

## What Makes This Different From Ctrl+F

Ctrl+F searches one document at a time
This tool searches every document simultaneously

Ctrl+F finds exact word matches only
This tool uses AND + OR logic across keyword groups

Ctrl+F gives you a list of page numbers
This tool gives you a formatted Excel + PowerPoint
with every matched sentence already pulled out

---

## The Code Guide

The file updated_Keyword_Boolean.docx explains every
single line of the script in plain English.
No programming experience needed to understand it.

---

## Built With

PyMuPDF      — reads PDF files page by page
python-docx  — reads Word documents
pandas       — builds and saves the Excel file
python-pptx  — creates the PowerPoint slides
os and re    — handle files and text splitting

---

## About

Built by Sahil — secondary research professional.

This is the first in a series of research automation
tools. Follow for upcoming tools including competitor
mention tracker, statistic extractor, and briefing
deck builder.
