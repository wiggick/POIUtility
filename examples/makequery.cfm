<!--- Small update to change this to a markdown table input 
||header1||header2||
|col1|col2|
--->
<!--- Kill extra space. --->
<cfsilent>

	<!--- We only care about this tag on the close. --->
	<cfif (THISTAG.ExecutionMode EQ "End")>
	
		<!--- 
			This is the variable name into which the query is 
			going to be stored.
		--->
		<cfparam
			name="ATTRIBUTES.Name"
			type="string"
			/>
	
		
		<!--- 
			Grab the generated content. This is the definition
			of the query columns and data.
		--->
		<cfset strSetup = Trim( THISTAG.GeneratedContent ) />
		
		<!--- Clear the generated content. --->
		<cfset THISTAG.GeneratedContent = "" />
		
		<!--- Clean up the generated content. --->
		<cfset strSetup = strSetup.ReplaceAll(
			"(?m)^[\t ]+|[\t ]+$",
			""
			) />
		
		<!--- Get the rows of data. --->
		<cfset arrRows = strSetup.Split( "[\r\n]+" ) />
		
		<!--- 
			Define the query using the first row of data.
			By default, all of these values are going to 
			be strings.
		--->
		<cfset qData = QueryNew( "" ) />
		<!--- Assertion that query will Always have a header row using || like markdown before and end 
		of row --->
		<cfset arrHeader=  ListtoArray(arrRows[1],"||",false)>
		<cfloop index="idxHeader" from="1" to="#ArrayLen(arrHeader)#">
			<cfset QueryAddColumn(
				qData, 
				arrHeader[idxHeader],
				"CF_SQL_VARCHAR",
				ArrayNew( 1 )
				) />
		</cfloop>
		
		<!--- Loop over the rest of the rows to add data. --->
		
		<cfloop
			index="intRow"
			from="2"
			to="#ArrayLen( arrRows )#"
			step="1">
			
			<!--- Add a row to the query. --->
			<cfset QueryAddRow( qData ) />
			<!--- we consider || null data at this point, and with markdown tables, each
			row begins and ends with a |, so for the listtoArray to work we need to rip of the first and last one--->
			<cfset cleanedRow = Mid(arrRows[intRow],2,len(arrRows[intRow]) - 2)>
			<cfset arrRow=  ListtoArray(cleanedRow,"|",true)>
			<!---<cfdump var="#qData.RecordCount#">
			<cfloop from="1" to="#ArrayLen(arrHeader)#" index="idx">
				<cfdump var="#arrHeader[idx]#">
			</cfloop>
			<cfdump var="#arrHeader#">
			<cfdump var="#arrRow#">
			<cfabort>--->
		
			<cfloop index="idxColumn" from="1" to="#ArrayLen(arrHeader)#">
				<cfset castType = (arrRow[idxColumn] eq "" ? "null" : "string" )>
				<cfset cellValue =(arrRow[idxColumn] eq "" ? "0" : "#arrRow[ idxColumn ]#" )>
				<!---<cfset castType ="string">--->
				<cfset qData[ arrHeader[idxColumn] ][ qData.RecordCount ] = JavaCast(
					castType,
					cellValue
					) />
			</cfloop>
								
		</cfloop>	
	
		
		<!--- Store the query into the caller. --->
		<cfset "CALLER.#ATTRIBUTES.Name#" = qData />
		
	</cfif>	
	
</cfsilent>