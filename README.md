# Excel-VBA-Make-KML-Example
# This repository contains the files necessary to follow the Youtube tutorial on making a Google Earth .kml file from addresses stored in a .csv file.
The data file is stored as .csv (comma separated values), then manually converted into .xlsm (Excel macro file format)
when opened and saved on the local system. Similarly, all of  the VBA code is stored in a .txt file
then manually renamed to .bas when following the tutorial. The .bas file is then imported
into .xlsm file and run through the VBA code  environment.
# Prerequisites:
1. Google Earth must be installed on the local machine.<br>
  Google Earth can be downloaded here. Google: (https://www.google.com/earth/versions/#download-pro)
2. The Developer tab must be visible in Excel. <br>
  Expose Excel Developer tab on Ribbon. Youtube: (https://www.youtube.com/watch?v=JLQ8OuW0FlY)
3. A trusted folder where VBA code can be run in Excel.<br>
  Add the folder where you want to build and run this exercise to Excel's Trusted Location list. Youtube: (https://www.youtube.com/watch?v=t5OcD1bk7Ek)
# Steps for Building the project
1.	Download “makeKMLAddress.zip” from “<>Code” tab on the Github. Repository: https://github.com/jcurrie2021/Excel-VBA-Make-KML-Example
2.	Open “makeKMLAddress.zip” and save the files to your local pc (place in Excel Trusted folder).
3.	Rename “makeKMLAddress.txt” to “makeKMLAddress.bas”
4.	Open “SanJoseDelicatessens4_26_2021.csv” in Excel and save the file (“SAVE AS”) “SanJoseDelicatessens4_26_2021.xlsm” (save as type “Excel Macro-Enabled Workbook”).
5.	Click on the “Developer” tab to access the “Visual Basic” code window. Click on the “Visual Basic” icon.
(the “Microsoft Visual Basic for Applications” window appears). 
6.	Right click on “Microsoft Excel Objects”, followed by clicking “Import File” from the menu. Select “makeKMLAddress.bas” and click the “Open” button (this adds the makeKMLAddress module to the project). 
7.	 Toggle to the Excel workbook. From the “Developer” tab click “Macros” (the Macros dialog box appears). Click on the macro “makeKMLAddress” followed by clicking the “Run” button. This will read all of the addresses on the current tab and create “SanJoseDelicatessens4_26_2021.kml” in your project folder.
8.	You can now double-click on the “SanJoseDelicatessens4_26_2021.kml” from the Windows “File Explorer” to view your .kml file in Google Earth.  
# The macro code explained "makeKMLAddress.bas"<br> 
```diff
Attribute VB_Name = "makeKMLAddress"
Sub makeKMLAddress() 'subroutine name. Not necessarily the same as the VB_Name

'Variables are declared 
Dim shead As String 'xml heading (type: string)
Dim sfoot As String 'xml footer (type: string) 
Dim lRow As Long 'last row in the active sheet (type: long)
Dim lFile As Long 'file handle (type: long)
Dim sFile As String 'file name (type: string) 
Dim sPath As String 'file path (type: string) 
Dim snl As String 'new line and line feed (type: string)
Dim sSht As String 'worksheet name (type: string)
```
```diff
'***********************************
'Populate local variables
'(note: the apostrophe represents a comment in Visual Basic).
```
```diff
'Sheet name containing addresses
sSht = ActiveSheet.Name
``` 
```diff
'Get return string for adding line feed in .kml output file 
snl = vbCr & vbLf 
``` 
```diff
'Get last row 
lRow = Sheets(sSht).UsedRange.Rows.Count 
```
```diff
'Get path and file 
sPath = ThisWorkbook.Path & "\"
```
```diff
'sFile = this  workbook.name 
sFile = Replace(ThisWorkbook.Name, ".xlsm", ".kml") 
```
```diff
'define header for  .kml file
shead = "<?xml version='1.0' encoding='UTF-8'?>" & vbCr & vbLf
shead = shead & "<kml xmlns='http://www.opengis.net/kml/2.2'>" & vbCr & vbLf
shead = shead & "<Document>" & snl
```
```diff
'Define .kml footer
sfoot = "</Document>" & vbCr & vbLf & "</kml>"
```
```diff
'***********************************
'Open .kml file for writing 
lFile = FreeFile 'Get file handle
Open sPath & sFile For Output As lFile 'open file using sPath as folder, sFile as filename, lFile as file handle
```
```diff
'Write the header to disk
Print #lFile, shead
```
```diff
'Loop through data set Placemarks: name, description and address
For x = 2 To lRow 'start on row 2 of the active worksheets and continue reading down to the last row
Print #lFile, "<Placemark>" 'Start a new placemark record
Print #lFile, "<name>" & CStr(Trim(Sheets(sSht).Cells(x, 1).Value)) & "</name>" 'Enter name
Print #lFile, "<description>" & Sheets(sSht).Cells(x, 3).Value & "</description>" 'Enter description
Print #lFile, "<address>" & Sheets(sSht).Cells(x, 2).Value & "</address>" 'Enter address
Print #lFile, "</Placemark>" & snl 'Close placemark and add a line feed
Next x
```
```diff
'Print the footer to finish building the .kml file
Print #lFile, sfoot
Close lFile 'close the open file handle
```
```diff
MsgBox "Finished" 'Show message box that the process has finished.
```
```diff
'handle errors
On Error GoTo errmakeKMLAddress
 
errmakeKMLAddressExit:
Exit Sub

errmakeKMLAddress:
    MsgBox Err.Description
    Resume errmakeKMLAddressExit
End Sub
```
