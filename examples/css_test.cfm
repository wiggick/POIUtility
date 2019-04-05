
<!--- Import the POI tag library. --->
<cfimport taglib="../lib/tags/poi/" prefix="poi" />


<!---
	Create an excel document and store binary data into REQUEST variable
	(if we also/or wanted to save this to disk, we could have supplied a
	"file" attribute). We are going to supply our branded template for
	this report by using the "template" attribute in the document tag.
--->


<poi:document
	name="REQUEST.ExcelData"
	createXLSX="true"
	EvaluateFormulas="false"
	style="font-family: verdana ; font-size: 10pt color:black; white-space: nowrap ;">

	<poi:css method="getPOIColors" result="myPOIColors" />
	<cfset colorList = StructKeyList(myPOIColors) />

	<poi:css method="getFillPatterns" result="myFillPatterns" />
	<cfset fillPatternList = StructKeyList(myFillPatterns) />

	<!--- Define style classes.--->
	<poi:classes>

		<poi:class
			name="title"
			style="font-family: arial ; color: white ; background-color: black ; font-size: 18pt ; text-align: left ;"
			/>

		<poi:class
			name="header"
			style="font-family: arial ; background-color: blue ; color: white ; font-size: 14pt ; border-bottom: solid 3px green ; border-top: 2px solid white ;"
			/>

	</poi:classes> 

	<!--- Define Sheets. --->
	<poi:sheets>
		<poi:sheet
			name="Color - Pattern - Region Test">		
		<cfset rowIndex = 1 />
		
		<poi:row class="Header" index="#rowIndex#">
			<poi:cell index="1" ColSpan="10" value="Cell Solid Background Color with Random Text Color Test" />
		</poi:row>
		<cfset cellindex = 0 />
		
			<cfloop from="1" to="#int(StructCount(myPOIColors)/10)#" index="colorRow">
				<cfset rowIndex += 1 />
				<poi:row index="#rowIndex#">
					<cfloop from="1" to="10" index="colorColumn">
						<cfset cellIndex += 1>
						<poi:cell index="#colorColumn#" style="border-bottom: solid 2px black; color: #ListGetAt(colorList, RandRange(1,StructCount(myPOIColors)))#; background-color: #ListGetAt(colorList,cellIndex)#;" type="string" value="#ListGetAt(colorList,cellIndex)#" />
					</cfloop>
				</poi:row>
			</cfloop>
			<!--- show the remainder --->
			
			<poi:row>
				<cfloop from="1" to="#StructCount(myPOIColors) mod 10#" index="colorColumn">
					<cfset cellIndex += 1>
					<poi:cell index="#colorColumn#" style="border-bottom: solid 2px black; background-color: #ListGetAt(colorList,cellIndex)#;" type="string" value="#ListGetAt(colorList,cellIndex)#" />
				</cfloop>
			</poi:row> 		
			<poi:row >
				<poi:cell index="1" ColSpan="10" class="header" value="Cell Pattern Test" />		
			</poi:row>
			<poi:row>
				<cfloop from="1" to="10" index="patternColumn">
					<poi:cell index="#patternColumn#" style="background-pattern: #ListGetAt(fillPatternList,patternColumn)#;" type="string" value="#ListGetAt(fillPatternList,patternColumn)#" />
				</cfloop>
			</poi:row>
				<poi:row>
				<cfloop from="1" to="9" index="patternColumn">
					<poi:cell index="#patternColumn#" style="background-pattern: #ListGetAt(fillPatternList,patternColumn + 10)#;" type="string" value="#ListGetAt(fillPatternList,patternColumn)#" />
				</cfloop>
			</poi:row>
			<poi:row>
				<poi:cell class="header" colspan="10" value="Colspan AND RowSpan Test" />
			</poi:row>
			<poi:row>
					<poi:cell value="Fantastic It supports rowspan" style="vertical-align: top; border-style: thick; border-style: thick; border-color: orange; border-top: thick;" ColSpan="5" RowSpan="3" />
					<poi:cell index="6" value="after the colspan" />
			</poi:row>
			<poi:row index="13" update="true">
				<poi:cell index="6" value="can I fill this in" />
			</poi:row>
			<poi:row index="14">
					<poi:cell value="after the rowspan" style="border-style:thick; border-color: red; color: blue;" />
			</poi:row>
			<poi:row>
				<poi:cell class="header" colspan="10" value="Color by HEX (closest match fo xls)" />
			</poi:row>
			<poi:row>
				<poi:cell value="##0022FF"/><poi:cell value="" style="background-color:##0022FF;"/>
				<poi:cell value="##ffee00"/><poi:cell value="" style="background-color:##ffee00;"/>
				<poi:cell value="##ffee66"/><poi:cell value="" style="background-color:##ffee66;"/>
			</poi:row>
	
		</poi:sheet>
	</poi:sheets>

</poi:document>


<!--- Tell the browser to expect an Excel file attachment. --->
<cfheader
	name="content-disposition"
	value="attachment; filename=poi_css_test_#DateFormat(now(),"YYYYMMDD")#.xlsx"
	/>

<!---
	Tell browser the length of the byte array output stream.
	This will help the browser provide download duration to
	the user.
--->
<cfheader
	name="content-length"
	value="#REQUEST.ExcelData.Size()#"
	/>

<!--- Stream the binary data to the user. --->
<cfcontent
	type="application/excel"
	variable="#REQUEST.ExcelData.ToByteArray()#"
	/>

