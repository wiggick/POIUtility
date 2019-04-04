<cfswitch expression="#THISTAG.ExecutionMode#">

    <cfcase value="Start">

        <!--- Get a reference to the document tag context. --->
        <cfset VARIABLES.DocumentTag = GetBaseTagData( "cf_document" ) />

    
        <!--- Param tag attributes. --->

        <cfparam
            name="ATTRIBUTES.method"
            type="string"
            default=""
            />

        <cfparam
            name="ATTRIBUTES.result"
            type="any"
            />

            <cfswitch expression="#ATTRIBUTES.method#">
                <cfcase value="getPOIColors,getFillPatterns">
                    <cfinvoke component="#VARIABLES.DocumentTag.CSSRule#" method="#Attributes.method#" returnVariable="VARIABLES.result"> 
                </cfcase>
                <cfdefaultcase>
                     <cfset caller[Attributes.result] = "">
                    <cfexit>
                </cfdefaultcase>
            </cfswitch>

         

        <cfset caller[Attributes.result] = VARIABLES.result>
    </cfcase>
        
    <cfcase value="End">
    </cfcase>
</cfswitch>
    