# ResumeParser

This program was created for Oakland University’s Career and Life Design Center when I was the Software Development Intern. It extracts work experience from student resumes made public on the university’s job board called Handshake. Any resume of type .doc, .docx (word documents), and .pdf files can be parsed. The program takes a folder of resumes as input. First, if a resume is a word document it is converted into a pdf file with spire.doc. Then the text of the resume is extracted with pdfplumber and parts of speech are analyzed with beautiful soup to detect the trigger word “experience”, find where new lines occur in the resume, and detect common bullet point characters. Once the keyword is found, the program will return that line and any following lines that start with a capital letter and not a bullet point. From there, the original file name, file name altered to extract the student name, original string after the keyword, any hyperlinks that the parser found, a value for number of lines added, and the altered string of experiences are exported to an excel document. 

Testing on student resumes, this program ran with a 92% success rate. Certain resume templates and resumes with images will not correctly parse. 


## To run

Clone this repository. Download all dependencies from requirements.txt. Add all resumes you wish to parse to the resumes folder (3 sample resumes are provided). Run python ResumeParser.py 

Notes: 
consider modifying lines 74 and 75 depending on your use case to prevent truncating file names. 
As of 3/25/2025 spaCy is incompatible with python version 3.13. Downgrade to Python 3.12 to use this program.

