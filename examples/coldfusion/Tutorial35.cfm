<!--
=======================================================================
Tutorial 35

This tutorial shows how to import Excel sheet to DataSet in ColdFusion.
The data is imported from a specific Excel sheet (For this example 
we use the Excel file generated in Tutorial 09).
=======================================================================
-->

	
Tutorial 35<br>
----------<br>


<!-- Create an instance of the class that imports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Import Excel sheet to ResultSet -->
Reading file C:\Samples\Tutorial09.xlsx<br><br>
<cfset rs = workbook.easy_ReadXLSXSheet_AsResultSet("C:\Samples\Tutorial09.xlsx", "First tab")>

<!-- Confirm import of Excl file -->
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


