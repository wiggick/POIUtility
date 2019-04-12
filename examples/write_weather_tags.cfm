    <cf_makequery name="qSensor">
||Comment||SensorName||SensorType||Latitude||Longitude||Location||
|Beach season begins the Friday before Memorial weekend and ends on Labor Day. Beach hours during the season are from 11am - 7pm|Calumet Beach|Water|41.714739|-87.527356|(41.714739000000002, -87.527355999999997)|
||63rd Street Weather Station|Weather|41.780992|-87.572619|(41.780991999999998, -87.572619000000003)|
||63rd Street Beach|Water|41.784561|-87.571453|(41.784560999999997, -87.571453000000005)|
||Oak Street Weather Station|Weather|41.901997|-87.622817|(41.901997000000001, -87.622816999999998)|
||Foster Weather Station|Weather|41.976464|-87.647525|(41.976464, -87.647525000000002)|
|Montrose Beach is located at 4400 N. Lake Shore Drive|Montrose Beach|Water|41.969094|-87.638003|(41.969093999999998, -87.638002999999998)|
||Osterman Beach|Water|41.987675|-87.651008|(41.987675000000003, -87.651008000000004)|
||Ohio Street Beach|Water|41.894328|-87.613083|(41.894328000000002, -87.613083000000003)|
||Rainbow Beach|Water|41.760147|-87.550081|(41.760147000000003, -87.550081000000006)|
    </cf_makequery>


<cfsetting showdebugoutput="true" />


<!--- Import the POI tag library. --->
<cfimport taglib="../lib/tags/poi/" prefix="poi" />
    
    
<!--- 
    Create an excel document and store binary data into 
    REQUEST variable. 
--->
<poi:document 
    name="REQUEST.ExcelData"
    createXLSX=false
    file="#ExpandPath( './weatherSensors.xls' )#"
    style="font-family: verdana ; font-size: 10pt ; color: black ; white-space: nowrap ;">
    
    <!--- Define style classes. --->
    <poi:classes>
        
        <poi:class
            name="title"
            style="font-family: arial ; color: white ; background-color: green ; font-size: 18pt ; text-align: left ;"
            />
        
        <poi:class 
            name="header" 
            style="font-family: arial ; background-color: lime ; color: white ; font-size: 14pt ; border-bottom: solid 3px green ; border-top: 2px solid white ;" 
            />
            
    </poi:classes>
        
    <!--- Define Sheets. --->
    <poi:sheets>
    
        <poi:sheet 
            name="Chicago Weather Sensors"
            freezerow="2"
            orientation="landscape"
            zoom="130%">
        
            <!--- Define global column styles. --->
            <poi:columns>
                <poi:column style="width: 250px ; text-align: left ;" />
                <poi:column style="width: 130px ;" />
                <poi:column style="width: 130px ;" />
                <poi:column style="width: 100px ; text-align: left ;" />
                <poi:column style="width: 350px ; text-align: left ;" />
            </poi:columns>
            
            <!--- Title row. --->
            <poi:row class="title">
                <poi:cell value="Chicago Weather Sensor Locations" colspan="5" />
            </poi:row>
            
            <!--- Header row. --->
            <poi:row class="header">
                <poi:cell value="Sensor Name" />
                <poi:cell value="Sensor Type" />
                <poi:cell value="Latitude" />
                <poi:cell value="Longitude" />
                <poi:cell value="Location" />
            </poi:row>
            <!--- Output the sensor locations. --->
            <cfloop query="qSensor">
            
                <poi:row>
                    <poi:cell value="#qSensor.SensorName#" Comment="#qSensor.Comment#" />
                    <poi:cell value="#qSensor.SensorType#" />
                    <poi:cell value="#qSensor.Latitude#" />
                    <poi:cell value="#qSensor.Longitude#" />
                    <poi:cell value="#qSensor.Location#" />
                </poi:row>
            
            </cfloop>
                
        </poi:sheet>
        
    </poi:sheets>
        
</poi:document>



<!--- Tell the browser to expect an Excel file attachment. --->
<cfheader
    name="content-disposition"
    value="attachment; filename=weatherSensors.xls"
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
