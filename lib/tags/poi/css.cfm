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
            name="ATTRIBUTES.var"
            type="any"
            default=""
            />


        <cfparam
            name="ATTRIBUTES.result"
            type="any"
            default=""
            />

            <cfswitch expression="#ATTRIBUTES.method#">
                <cfcase value="getPOIColors">
                    <cfset VARIABLES.result = VARIABLES.DocumentTag.cssRule.getPOIColors()>
                </cfcase>
                 <cfcase value="getFillPatterns">
                    <cfset VARIABLES.result = VARIABLES.DocumentTag.cssRule.getFillPatterns()>
                </cfcase>
                <cfcase value="getBorderStyles">
                    <cfset VARIABLES.result = VARIABLES.DocumentTag.cssRule.getBorderStyles()>
                </cfcase>
            </cfswitch>

        <cfif ATTRIBUTES.result neq "" >
            <cfif StructKeyExists(VARIABLES,"result")>
                <cfset caller[Attributes.result] = VARIABLES.result>
            <cfelse>
                <cfset caller[Attributes.result]>
            </cfif>
        </cfif>
    </cfcase>
        
    <cfcase value="End">
    </cfcase>
</cfswitch>
    