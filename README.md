Reistee Extracts IServ - Teachers' Easy Extractor
=======

### Practical tool for teachers receiving student solutions on IServ.

When downloading student solutions from IServ you get a zip-file containing a folder for each student. These folders may contain many files of different formats. Extracting and merging them every week for correction is a tedious task. reistee aims to make this easier for teachers. Just drag and drop the zip-file onto the reistee.exe or reistee.py and all recognised files will be merged into one single pdf for every student. The pdf will be named "lastname, fistname.pdf" for better sorting in the file explorer.

### Download

Download the latest [release for windows here](https://github.com/MarcFranke/reistee/releases/download/v0.00.787-alpha/reistee.exe). LibreOffice is needed for converting odt, doc and docx files. Be sure to either have it installed in the standard directory (C:\Program Files\LibreOffice), or that the directory libreoffice\program is in your PATH environment variable.


### Usage

Just drag and drop a zip-file or folder onto the reistee.exe. Or use from the command line:
```powershell
.\reistee.exe solutions.zip
```
A folder called `solutions - reistee` will be created, with one pdf for every student.

### FAQ
Q: How do I use this?

A: Look above.

Q: Does OpenOffice also work?

A: Sadly no, as OpenOffice does not provide a way to convert files to pdf via commandline.

Q: Does MS Word also work?

A: Maybe? The code for converting files with MS Word was added in Version 0.00.787. But since I currently do not have access to a PC with MS Word, I cannot test this. I am happy for feedback from people who can test this.

Q: odt, doc and docx Documents are not included in the pdf!?

A: reistee can't find your LibreOffice installation. Is LibreOffice installed?

Q: LibreOffice is installed, but odt, doc and docx Documents are still not included.

A: Do you have it installed in a different directory than "C:\Program Files\LibreOffice"? If so, please add the folder Path\to\libreoffice\program to your PATH environment variable.

Q: reistee crashes!

A: Do you maybe have LibreOffice in PATH, but LibreOffice is no longer installed there? If not, write me an E-Mail at reistee@marcfranke.de, ideally include the zip and a screenshot or a copy of the error message.
