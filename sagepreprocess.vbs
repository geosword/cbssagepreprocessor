rem VB script to consolidate transactions lines commonsense exported csv files, so that the standing charge and click charge appear in Sage as 
rem one single "plan printer service" line, rather than a seperate standing charge and click charge
Option Explicit

Const ForReading = 1 

Const NEWDESCRIPTION="""PLAN PRINTER SERVICE"""
Const TRANSTYPE=0
Const ACCOUNTID=1
Const NOMCODE=2
Const SOMETHING=3
Const TRANSDATE=4
Const DOCID=5
Const DESCRIPTION=6
Const AMOUNT=7
Const TAXCODE=8
Const VAT=9

rem a place to store records that already exist in the dictionary
dim existingRecord
rem to write to the file to import into sage
Dim record
rem the name of the file to write to
Dim sovFile
Dim strNextLine
dim records()
dim objTextFile
dim i
dim wShell
dim oExec
dim exportFile
dim sFileSelected
dim objFSO


Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
rem sFileSelected is the commonsense file
rem now to load it and process it.

Set objFSO = CreateObject("Scripting.FileSystemObject") 

exportFile=sFileSelected + ".sov.csv"

i=0
if objFSO.FileExists(sFileSelected) then
	Set objTextFile = objFSO.OpenTextFile _ 
		(sFileSelected, ForReading) 
	
	rem create a dictionary object that will serve as an associative array to hold our records (indexed on DOCID)
	dim writeRecords
	Set writeRecords=CreateObject("Scripting.Dictionary")
	writeRecords.CompareMode = vbTextCompare
	
	Do Until objTextFile.AtEndOfStream 
		strNextLine = objTextFile.Readline 
		record = Split(strNextLine , ",") 
		
		rem do some EXTREMELY RUDIMENTARY TESTING THAT IT IS A COMMONSENSE EXPORT FILE
		if record(TRANSTYPE)<>"SI" then
			MsgBox("This doesnt look like a commonsense exported file")
			objTextFile.Close
			wscript.quit(29001)
		end if
		
		if writeRecords.exists(record(DOCID)) then
			rem it already exists, so take AMOUNT of this record, and add it to the amount already in the dictionary.
			rem first we need to get the existing record.
			existingRecord = writeRecords(record(DOCID))
			existingRecord(AMOUNT)=csng(existingRecord(AMOUNT)) + csng(record(AMOUNT))
			existingRecord(VAT)=csng(existingRecord(VAT)) + csng(record(VAT))
			rem then write it back to the dictionary
			writeRecords(record(DOCID))=existingRecord
		 else
			rem it doesnt exist so add it.
			rem adjusting the description accordingly
			record(DESCRIPTION)= NEWDESCRIPTION
			writeRecords.Add record(DOCID), record
		end if
		rem Redim Preserve records(i)
		rem records(i)=thisRecord
		rem i=i+1
	Loop 
	objTextFile.Close

	Set sovFile = objFSO.CreateTextFile(exportFile,True)
	dim dictItems
	dictItems = writeRecords.Items
	for Each record in dictItems
		rem Wscript.Echo record(TRANSTYPE) & "," & record(DOCID) & " " & record(AMOUNT) & " " & record(DESCRIPTION)
		sovFile.Write Join(record,",") & vbCrLf
	next
end if
MsgBox(sFileSelected & " processed. Now import " & exportFile & " into sage")
sovFile.Close()
