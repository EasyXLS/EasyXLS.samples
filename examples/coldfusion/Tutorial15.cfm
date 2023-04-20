<!--
===================================================================================
Tutorial 15

This tutorial shows how to create an Excel file with hyperlinks in Java.

EasyXLS supports the following hyperlink types:
   (1) - hyperlink to URL
   (2) - hyperlink to file
   (3) - hyperlink to UNC
   (4) - hyperlink to cell in the same Excel file
   (5) - hyperlink to name

Every type of hyperlink accepts a tool tip description.

Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
===================================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.HyperlinkType" name="HyperlinkType" action="CREATE">


Tutorial 15<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<cfset xlsTab1 = workbook.easy_getSheetAt(0)>
<cfset xlsTab2 = workbook.easy_getSheetAt(1)>
	
<!-- Create hyperlink to URL -->
<cfset xlsTab1.easy_addHyperlink(HyperlinkType.URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2")>

<!-- Create hyperlink to file -->
<cfset xlsTab1.easy_addHyperlink(HyperlinkType.FILE, "c:\myfile.xls", "Link to file", "B3")>

<!-- Create hyperlink to UNC -->
<cfset xlsTab1.easy_addHyperlink(HyperlinkType.UNC, "\\computerName\Folder\file.txt", "Link to UNC", "B4:D4")>

<!-- Create hyperlink to cell on second sheet -->
<cfset xlsTab1.easy_addHyperlink(HyperlinkType.CELL, "'Second tab'!D3", "Link to CELL", "B5")>

<!-- Create a name on the second sheet -->
<cfset xlsTab2.easy_addName("Name", "=Second tab!$A$1:$A$4")>
	
<!-- Create hyperlink to name -->
<cfset xlsTab1.easy_addHyperlink(HyperlinkType.CELL, "Name", "Link to a name", "B6")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx")>

<!-- Confirm export of Excel file -->
<cfset sError = workbook.easy_getError()>
<CFIF (sError  IS "")>
  <cfoutput>
	File successfully created.
  </cfoutput>
<CFELSE>
  <cfoutput>
	Error encountered:  #sError#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>