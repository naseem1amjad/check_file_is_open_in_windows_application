# Checks if a file is opened by a Windows Software
A python script that checks whether a file (doc, docx) is open in a windows application (MS Word, WordPad)
<br>
## Running of the Python script
### In Case Doc file is not opened
monitoring ['C:\\Docs', 'C:\\Users\\Naseem\\Desktop'] for ['*.doc', '*.docx'] <br>
No DOC or DOCX file is open in another application.<br>
<br>

### In Case Doc file is opened by MS Word
monitoring ['C:\\Docs', 'C:\\Users\\Naseem\\Desktop'] for ['*.doc', '*.docx'] <br>
The following DOC or DOCX files are open in another application: <br>
File: C:\Docs\MyDoc.docx <br>
Opened by: C:\Program Files\Microsoft Office\Office\WINWORD.EXE <br>
<br>

<br>
<br>
Script developed by Naseem Amjad (urdujini@gmail.com)