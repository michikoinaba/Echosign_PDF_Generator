# Echosign_PDF_Generator

 * this tool is used to retreive multiple files with homeowner's email addresses from the spreadsheet
 * 1. get the excel sheet with the list of homeowners' email addresses.
 * 2. read the spreadsheet and generate an array with all email addresses.
 * 3. loop through the email array and get an echsign doc_key from the documents table
 * 4. get a pdf form content with the doc_key from echosign API.
 * 5. create a new pdf file with the content from echosign API and save it locally.
 * 6. put all pdf files into a zip file.
