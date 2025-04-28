# ResumeParser

## What it does

This resume parser can extract data from .doc, .docx, and .pdf files. Using spire.doc, word document resumes are converted into pdf files. Text is extracted with pdfplumber and parts of speech are determined with beautiful soup. From there, any line that does not start with a common symbol for bullet points is extracted and placed into a word document.

## To run

Modify lines 74 and 75 depending on the filenames of the resume. Currently this is implemented for handshake student resumes to remove the text _resume_.pdf from the filename.

Modify lines 13 and 14 to point to a different excel file or resume folder.

Download the dependencies from requirements.txt. Create a folder in the working directory called "Resumes" and add all resumes you wish to parse to that file. Run python ResumeParser.py.

Note that as of 3/25/2025 spaCy is incompatable with python version 3.13. Downgrade to Python 3.12 to use this program.
