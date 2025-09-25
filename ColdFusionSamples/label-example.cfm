<!--- Example: Label Office files after creation --->
<cfscript>
  // Example file paths (adjust for your environment)
  docxFile = "C:\drops\report.docx";
  xlsxFile = "C:\drops\data.xlsx";
  
  // Apply default OFFICIAL label to DOCX
  try {
    success = application.mip.ApplyLabelToFile(docxFile, "", "Applied by TRIS at creation");
    writeOutput("DOCX labeling: " & (success ? "SUCCESS" : "FAILED") & "<br>");
  } catch (any e) {
    writeOutput("DOCX labeling ERROR: " & e.message & "<br>");
  }
  
  // Apply specific label to XLSX (example: OFFICIAL: Sensitive)
  sensitiveLabel = "other-label-guid-here"; // Replace with actual GUID
  try {
    success = application.mip.ApplyLabelToFile(xlsxFile, sensitiveLabel, "Sensitive data classification");
    writeOutput("XLSX labeling: " & (success ? "SUCCESS" : "FAILED") & "<br>");
  } catch (any e) {
    writeOutput("XLSX labeling ERROR: " & e.message & "<br>");
  }
  
  // Verify applied labels
  try {
    docxLabel = application.mip.GetAppliedLabelId(docxFile);
    xlsxLabel = application.mip.GetAppliedLabelId(xlsxFile);
    
    writeOutput("DOCX label: " & (len(docxLabel) ? docxLabel : "NONE") & "<br>");
    writeOutput("XLSX label: " & (len(xlsxLabel) ? xlsxLabel : "NONE") & "<br>");
  } catch (any e) {
    writeOutput("Label verification ERROR: " & e.message & "<br>");
  }
</cfscript>