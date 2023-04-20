<!--
=================================================================== 
Tutorial 34

This tutorial shows how to import Excel to ResultSet in ColdFusion.
is imported from the active sheet of the Excel file (the Excel file 
generated in Tutorial 09).
===================================================================
-->

	
Tutorial 34<br>
----------<br>


<!-- Create an instance of the class that imports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Import Excel to ResultSet -->
Reading file C:\Samples\Tutorial09.xlsx<br><br>
<cfset rs = workbook.easy_ReadXLSXActiveSheet_AsResultSet("C:\Samples\Tutorial09.xlsx")>

<!-- Confirm import of Excel file -->
<cfset sError = workbook.easy_getError()>
<CFIF (sError  IS "")>
  <cfoutput>
	<!-- Display imported ResultSet values -->
	<cfset row = 0>
	<cfloop condition="#rs.next()#">
		<cfset columnCount = rs.getMetaData().getColumnCount()>
		<cfloop from="1" to="#columnCount#" index="column">
			<cfoutput>
				At row #evaluate(row + 1)#, column #evaluate(column)# the value is '#rs.getString(JavaCast("int",column))#'<br>
			</cfoutput>
		</cfloop>
		<cfset row = row + 1>
	</cfloop>
  </cfoutput>
<CFELSE>
  <cfoutput>
	Error encountered:  #sError#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>


