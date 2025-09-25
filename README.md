# ColdFusion MIP Labeller

A .NET library that enables ColdFusion applications to apply Microsoft Purview sensitivity labels to Office documents using the Microsoft Information Protection (MIP) SDK.

## Features

- Apply sensitivity labels to Word (.docx) and Excel (.xlsx) files
- Retrieve existing labels from Office documents
- Thread-safe singleton implementation
- Comprehensive error handling and debugging methods
- Support for proxy configurations
- Headless authentication using Azure AD client credentials

## Prerequisites

- .NET Framework 4.8 or later
- ColdFusion 2021+ with .NET Integration Services
- Azure AD application with appropriate MIP permissions
- Microsoft Information Protection SDK native libraries

## Installation

### 1. NuGet Package (Recommended)
```
Install-Package ColdFusionMIPLabeller
```

### 2. Manual Installation
1. Download the latest release from GitHub
2. Copy `ColdFusionMIPLabeller.dll` to your ColdFusion lib directory
3. Ensure MIP SDK native libraries are accessible (see Configuration section)

## Configuration

### Azure AD Application Setup

1. Register an application in Azure AD
2. Grant the following API permissions:
   - `InformationProtectionPolicy.Read.All`
   - `Content.SuperUser` (for Microsoft Purview)
3. Create a client secret
4. Note the Tenant ID, Client ID, and Client Secret

### MIP Native Libraries

The library requires Microsoft Information Protection SDK native libraries. Set one of:

1. **Environment Variable**: Set `MIP_NATIVE_PATH` to the directory containing MIP native DLLs
2. **Standard Locations**: The library automatically checks:
   - `C:\ColdFusion2023\cfusion\runtime\lib\MIP\x64`
   - `C:\ColdFusion2021\cfusion\runtime\lib\MIP\x64`
   - `C:\Program Files\Microsoft Information Protection SDK\Bin\x64`

## Usage

### Basic Setup

```coldfusion
<!--- Configure the labeler (typically in Application.cfc) --->
<cfscript>
// Configure MIP Labeler
CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").Configure(
    tenantId = "your-tenant-id",
    clientId = "your-client-id", 
    clientSecret = "your-client-secret",
    defaultLabelId = "your-default-label-guid",
    userIdentity = "service@yourcompany.com"  // Optional
);

// Get singleton instance
application.mipLabeler = CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").Instance;
</cfscript>
```

### Apply Labels to Files

```coldfusion
<cfscript>
// Apply default label to a Word document
success = application.mipLabeler.ApplyLabelToWordFile(
    filePath = "C:\documents\report.docx",
    labelId = "",  // Empty = use default
    justification = "Applied by automated system"
);

// Apply specific label to Excel file
success = application.mipLabeler.ApplyLabelToExcelFile(
    filePath = "C:\documents\data.xlsx", 
    labelId = "specific-label-guid",
    justification = "Sensitive financial data"
);

// Apply label to any supported file type
success = application.mipLabeler.ApplyLabelToFile(
    filePath = "C:\documents\presentation.pptx",
    labelId = "",
    justification = "Standard classification"
);
</cfscript>
```

### Read Existing Labels

```coldfusion
<cfscript>
// Get the label ID from a file
labelId = application.mipLabeler.GetAppliedLabelId("C:\documents\report.docx");

if (len(labelId)) {
    writeOutput("File has label: " & labelId);
} else {
    writeOutput("File has no label applied");
}
</cfscript>
```

### Using the CFC Wrapper

```coldfusion
<!--- Create component instance --->
<cfset labeling = CreateObject("component", "components.Labeling").init()>

<!--- Apply label using CFC --->
<cfset success = labeling.labelFile(
    filePath = "C:\documents\report.docx",
    labelId = "",
    justification = "Applied via CFC wrapper"
)>

<!--- Check if file is labeled --->
<cfif labeling.isLabeled("C:\documents\report.docx")>
    <cfoutput>File is labeled with: #labeling.getLabelId("C:\documents\report.docx")#</cfoutput>
</cfif>
```

## API Reference

### Static Methods

#### `Configure(tenantId, clientId, clientSecret, defaultLabelId, userIdentity?)`
Configures the MIP labeler with Azure AD credentials.

#### `Create(tenantId, clientId, clientSecret, defaultLabelId, userIdentity?)`
Configures and returns the singleton instance.

#### `WarmUp()`
Initializes the MIP SDK without performing labeling operations.

### Instance Methods

#### `ApplyLabelToFile(filePath, labelId?, justification?)`
Applies a sensitivity label to any supported Office file.
- **Returns**: `boolean` - Success status
- **Parameters**:
  - `filePath`: Absolute path to the file
  - `labelId`: Label GUID (optional, uses default if empty)
  - `justification`: Reason for labeling (optional)

#### `ApplyLabelToWordFile(filePath, labelId?, justification?)`
Specifically for Word documents (.docx files).

#### `ApplyLabelToExcelFile(filePath, labelId?, justification?)`
Specifically for Excel spreadsheets (.xlsx files).

#### `GetAppliedLabelId(filePath)`
Retrieves the current label GUID from a file.
- **Returns**: `string` - Label GUID or empty string if no label

### Diagnostic Methods

#### `GetConfigurationStatus()`
Returns current configuration status for debugging.

#### `ValidateConfiguration()`
Validates the current configuration without initializing MIP SDK.

#### `TestInitialization()`
Tests if MIP SDK can be initialized with current configuration.

## Error Handling

The library throws detailed exceptions with context information:

```coldfusion
<cftry>
    <cfset success = application.mipLabeler.ApplyLabelToFile(filePath, "", "")>
    <cfcatch type="any">
        <cflog file="mip-errors" text="Labeling failed: #cfcatch.message#">
        <!--- Handle error appropriately --->
    </cfcatch>
</cftry>
```

## Troubleshooting

### Common Issues

1. **"Native MIP folder not found"**
   - Ensure MIP SDK native libraries are installed
   - Set `MIP_NATIVE_PATH` environment variable
   - Check that the path contains required DLL files

2. **Authentication failures**
   - Verify Azure AD application permissions
   - Check tenant ID, client ID, and client secret
   - Ensure the service account has access to sensitivity labels

3. **File access errors**
   - Ensure files are not open in other applications
   - Verify ColdFusion service has write permissions
   - Check that file paths are absolute

### Debug Methods

Use the debug methods for detailed troubleshooting:

```coldfusion
<cfscript>
// Check configuration
writeOutput(CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").ValidateConfiguration());

// Test initialization
writeOutput(CreateObject(".NET", "ColdFusionMIPLabeller.Labeler").TestInitialization());

// Debug specific file labeling
result = application.mipLabeler.DebugApplyLabel("C:\test\document.docx");
writeOutput(result);
</cfscript>
```

## Supported File Types

- Microsoft Word (.docx)
- Microsoft Excel (.xlsx)
- Microsoft PowerPoint (.pptx) - via `ApplyLabelToFile()`

## Security Considerations

- Store Azure AD credentials securely (use ColdFusion's secure profile or environment variables)
- Use least-privilege principle for Azure AD application permissions
- Regularly rotate client secrets
- Monitor labeling activities through Azure AD audit logs

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- Create an issue on GitHub for bug reports
- Check existing issues for known problems
- Provide detailed error messages and configuration when reporting issues

## Changelog

### v1.0.0
- Initial release
- Support for Word and Excel labeling
- ColdFusion integration
- Comprehensive error handling and debugging