# VaocabularyExtractor
 
Extracts words and descriptions out of an PDF file into an Excel sheet. Works for the vocabulary lists of Helbling.

### How to use:
Insert the file path of the PDF with the vocabulary into the `file` variable as string. (line 8)
Write the side you want to start and end the extraction into `sides` variable. (line 9).

> [!warning]
> Sometimes the extraction stops some colums before the end of the page. Probably a bug of PDFPlumber.
