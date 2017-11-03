# RTF-Cleaner
RTF de-obfuscator for CVE-2017-0199 documents to find URLs statically.

Use this tool to statically find URLs in obfuscated RTF documents. 
Usage: .\rtfCleanerUniversal1.1.ps1 path-to-fil
Example: .\rtfCleanerUniversal1.1.ps1 'C:\Users\yams\Downloads\invoiceForYams.doc'
If the built in regexers don't find anything just enter what you find in the rtf file, IE pntxtb, lchars, etc... 
Looking to build in functionality to find these and do it automatically. Its manual for now though.
This works for any file type that contains {\parameters} to clean.
Use at your own risk!
