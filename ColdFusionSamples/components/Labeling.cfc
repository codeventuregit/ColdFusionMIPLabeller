<cfcomponent displayname="Labeling" hint="CFC wrapper for MIP labeling operations">
  
  <cffunction name="init" access="public" returntype="Labeling" hint="Initialize the labeling component">
    <cfset variables.labeler = CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").Instance>
    <cfreturn this>
  </cffunction>
  
  <cffunction name="labelFile" access="public" returntype="boolean" hint="Apply sensitivity label to file">
    <cfargument name="filePath" type="string" required="true" hint="Absolute path to file">
    <cfargument name="labelId" type="string" required="false" default="" hint="Label GUID or empty for default">
    <cfargument name="justification" type="string" required="false" default="Applied by TRIS" hint="Justification text">
    
    <cftry>
      <cfreturn variables.labeler.ApplyLabelToFile(arguments.filePath, arguments.labelId, arguments.justification)>
      <cfcatch>
        <cflog file="mip-labeling" text="Error labeling #arguments.filePath#: #cfcatch.message#">
        <cfreturn false>
      </cfcatch>
    </cftry>
  </cffunction>
  
  <cffunction name="getLabelId" access="public" returntype="string" hint="Get applied label GUID from file">
    <cfargument name="filePath" type="string" required="true" hint="Absolute path to file">
    
    <cftry>
      <cfreturn variables.labeler.GetAppliedLabelId(arguments.filePath)>
      <cfcatch>
        <cflog file="mip-labeling" text="Error reading label from #arguments.filePath#: #cfcatch.message#">
        <cfreturn "">
      </cfcatch>
    </cftry>
  </cffunction>
  
  <cffunction name="isLabeled" access="public" returntype="boolean" hint="Check if file has any label applied">
    <cfargument name="filePath" type="string" required="true" hint="Absolute path to file">
    
    <cfset var labelId = getLabelId(arguments.filePath)>
    <cfreturn len(labelId) GT 0>
  </cffunction>
  
</cfcomponent>