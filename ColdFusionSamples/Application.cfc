<cfcomponent>
  <cfset this.name = "TRISApp">
  <cfset this.applicationTimeout = CreateTimeSpan(1, 0, 0, 0)>
  
  <cffunction name="onApplicationStart" returntype="boolean">
    <!--- Initialize MIP Labeler as application-scoped singleton --->
    <cfset application.mip = CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").Instance>
    
    <cflog file="application" text="MIP Labeler initialized successfully">
    <cfreturn true>
  </cffunction>
  
  <cffunction name="onError" returntype="void">
    <cfargument name="exception" required="true">
    <cfargument name="eventname" required="false" default="">
    
    <cflog file="application" text="Error in #arguments.eventname#: #arguments.exception.message#">
  </cffunction>
</cfcomponent>