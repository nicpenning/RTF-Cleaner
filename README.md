# RTF-Cleaner
RTF de-obfuscator for CVE-2017-0199 documents to find URLs statically.

Use this tool to statically find URLs in obfuscated RTF documents. 

Usage: .\rtfCleanerUniversal.ps1 path-to-file

Example: .\rtfCleanerUniversal.ps1 'C:\Users\yams\Downloads\invoiceForYams.doc'

If the built in regexers don't find anything just enter what you find in the rtf file, IE pntxtb, lchars, etc... 

Looking to build in functionality to find these and do it automatically. Its manual for now though.

This works for any file type that contains {\parameters} to clean.

Use at your own risk! Tested on Windows and PowerShell for MacOS.

How it works:

1. Removes all parameters that break up hex code (built in and custom ones)

2. Converts relevant hex to readable strings

3. Removes Null Characters

4. Finds and provides URL to HTA

Output if a URL is found:

PS /Users/yammer/Documents/PowerShell/RTF-Cleaner> ./rtfCleanerUniversal.ps1 ../RTFs/invoice.doc  

Analyzing: ../RTFs/invoice.doc

180 of \{\\lchars\s*([^}]*?)\s*} Found!

URL Found!

https://malicious.site/fake/directory/payload.hta
