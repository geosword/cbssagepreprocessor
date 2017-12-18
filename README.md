# cbssagepreprocessor
A vbscript preprocessor which consolidates invoice lines for importing from a very niche service management tool, into Sage Line 50
This script is *highly* unlikely to be useful to *anyone* but me
Run in windows with 
cscript sagepreprocess.vbs, choose a suitable .csv to process, and file.csv.sov will be created. 
Lines are merged (added) based on the 5th field, which is the invoice number