<cfoutput>

	<!--- Create an instance of the POIUtility.cfc. --->
	<cfset objPOI = new lib.POIUtility() />


	<!---
		Read in the Exercises excel sheet. This has Push, Pull,
		and Leg exercises split up on to three different sheets.
		By default, the POI Utilty will read in all three sheets
		from the workbook. Since our excel sheet has a header
		row, we want to strip it out of our returned queries.

		Try playing around with the following

		 ColumnsToRead=2, ColumnStart = 1
	--->
	<cfset arrSheets = objPOI.ReadExcel(
		FilePath = ExpandPath( "./exercises.xls" ),
		HasHeaderRow = true
		) />


	<!--- CW  -- I've made the query standard by having propery column names, you can
	still do it the old way as below, or you can just get the columnlist from the query
	--->

	<!---
		The ReadExcel() has returned an array of sheet object.
		Let's loop over sheets and output the data. NOTE: This
		could be also done to insert into a DATABASE!
	--->
	<cfloop
		index="intSheet"
		from="1"
		to="#ArrayLen( arrSheets )#"
		step="1">

		<!--- Get a short hand to the current sheet. --->
		<cfset objSheet = arrSheets[ intSheet ] />


		<!---
			Output the name of the sheet. This is taken from
			the Tabs at the bottom of the workbook.
		--->
		<h3>
			#objSheet.Name#
		</h3>

		<!---
			Output the data from the Excel sheet in a table.
			We know the structrure of the Excel, so we can
			use the auto-named columns. Also, since we flagged
			the workbook as using column headers, the first
			row of the excel was stripped out and put into an
			array of column names.
		--->

		<table border="1">

			<cfset colList = ArrayToList(objSheet.Query.getColumnList())>
			<tr>
			<cfloop list="#colList#" index="col">
			<th><cfoutput>#col#</cfoutput></th>
			</cfloop>
			</tr>

		<!--- Loop over the data query. --->
		<cfloop query="objSheet.Query">
			<tr>
			<cfloop list="#colList#" index="colName">
				<td>#objSheet.Query[colName][objSheet.Query.currentRow]#</td>
			</cfloop>
			</tr>
		</cfloop>
		</table>

	</cfloop>

</cfoutput>
