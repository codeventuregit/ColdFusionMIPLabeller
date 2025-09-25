using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;

namespace ColdFusionMIPLabeller
{
    /// <summary>
    /// Main labeling service for applying Microsoft Purview sensitivity labels to Office files.
    /// Thread-safe singleton implementation for ColdFusion .NET Integration Service.
    /// </summary>
    public class Labeler
    {
        // Configuration provided by ColdFusion via Configure() method
        public static string TenantId = "";
        public static string ClientId = "";
        public static string ClientSecret = "";
        public static string DefaultLabelId = "";
        public static string UserIdentity = "";

        private static readonly Lazy<Labeler> _instance = new Lazy<Labeler>(() => new Labeler());
        private static readonly object _lock = new object();
        private IFileEngine? _fileEngine;
        private bool _initialized = false;

        public static Labeler Instance => _instance.Value;

        private Labeler() { }

        /// <summary>
        /// Gets the singleton instance without reflection.
        /// </summary>
        /// <returns>The singleton Labeler instance</returns>
        public static Labeler GetInstance() => Instance;

        /// <summary>
        /// Configures the static fields with provided values.
        /// </summary>
        /// <param name="tenantId">Azure AD tenant ID</param>
        /// <param name="clientId">Application client ID</param>
        /// <param name="clientSecret">Application client secret</param>
        /// <param name="defaultLabelId">Default sensitivity label GUID</param>
        /// <param name="userIdentity">User email for MIP identity (optional, defaults to service account)</param>
        public static void Configure(string tenantId, string clientId, string clientSecret, string defaultLabelId, string? userIdentity = null)
        {
            if (string.IsNullOrEmpty(tenantId)) throw new ArgumentException("Tenant ID cannot be null or empty", nameof(tenantId));
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentException("Client ID cannot be null or empty", nameof(clientId));
            if (string.IsNullOrEmpty(clientSecret)) throw new ArgumentException("Client secret cannot be null or empty", nameof(clientSecret));
            if (string.IsNullOrEmpty(defaultLabelId)) throw new ArgumentException("Default label ID cannot be null or empty", nameof(defaultLabelId));

            TenantId = tenantId;
            ClientId = clientId;
            ClientSecret = clientSecret;
            DefaultLabelId = defaultLabelId;
            UserIdentity = userIdentity;
        }

        /// <summary>
        /// Configures and returns the singleton instance.
        /// </summary>
        /// <param name="tenantId">Azure AD tenant ID</param>
        /// <param name="clientId">Application client ID</param>
        /// <param name="clientSecret">Application client secret</param>
        /// <param name="defaultLabelId">Default sensitivity label GUID</param>
        /// <param name="userIdentity">User email for MIP identity (optional)</param>
        /// <returns>The configured singleton instance</returns>
        public static Labeler Create(string tenantId, string clientId, string clientSecret, string defaultLabelId, string? userIdentity = null)
        {
            Configure(tenantId, clientId, clientSecret, defaultLabelId, userIdentity);
            return Instance;
        }



        /// <summary>
        /// Initializes the MIP SDK without performing any labeling operations.
        /// </summary>
        public static void WarmUp() => Instance.EnsureInitialized();

        /// <summary>
        /// Resets initialization state for testing purposes.
        /// </summary>
        public static void ResetForTesting()
        {
            lock (_lock)
            {
                Instance._initialized = false;
                Instance._fileEngine = null;
            }
        }

        /// <summary>
        /// Apply sensitivity label to Word document.
        /// </summary>
        /// <param name="filePath">Absolute path to DOCX file</param>
        /// <param name="labelId">Label GUID or null for default</param>
        /// <param name="justification">Justification text or null for default</param>
        /// <returns>True on success</returns>
        public bool ApplyLabelToWordFile(string filePath, string? labelId, string? justification)
        {
            try
            {
                ValidateFilePath(filePath);
                if (!filePath.ToLowerInvariant().EndsWith(".docx"))
                    throw new ArgumentException("File must be a Word document (.docx)", nameof(filePath));
                    
                EnsureInitialized();

                var effectiveLabelId = string.IsNullOrEmpty(labelId) ? DefaultLabelId : labelId;
                var effectiveJustification = string.IsNullOrEmpty(justification) ? "Applied by TRIS to Word document" : justification;

                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    throw new InvalidOperationException($"Label not found: {effectiveLabelId}");

                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = effectiveJustification
                };

                handler.SetLabel(label, options, new ProtectionSettings());
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                // Replace original file with labeled version
                if (result && File.Exists(outputPath))
                {
                    File.Delete(filePath);
                    File.Move(outputPath, filePath);
                }
                
                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to apply label to Word file {filePath}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Apply sensitivity label to Excel spreadsheet.
        /// </summary>
        /// <param name="filePath">Absolute path to XLSX file</param>
        /// <param name="labelId">Label GUID or null for default</param>
        /// <param name="justification">Justification text or null for default</param>
        /// <returns>True on success</returns>
        public bool ApplyLabelToExcelFile(string filePath, string? labelId, string? justification)
        {
            try
            {
                ValidateFilePath(filePath);
                if (!filePath.ToLowerInvariant().EndsWith(".xlsx"))
                    throw new ArgumentException("File must be an Excel spreadsheet (.xlsx)", nameof(filePath));
                    
                EnsureInitialized();

                var effectiveLabelId = string.IsNullOrEmpty(labelId) ? DefaultLabelId : labelId;
                var effectiveJustification = string.IsNullOrEmpty(justification) ? "Applied by TRIS to Excel spreadsheet" : justification;

                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    throw new InvalidOperationException($"Label not found: {effectiveLabelId}");

                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = effectiveJustification
                };

                handler.SetLabel(label, options, new ProtectionSettings());
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                if (result && File.Exists(outputPath))
                {
                    File.Delete(filePath);
                    File.Move(outputPath, filePath);
                }
                
                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to apply label to Excel file {filePath}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Apply sensitivity label to file. If labelId is null/empty, uses DefaultLabelId.
        /// </summary>
        /// <param name="filePath">Absolute path to file (must exist and be closed)</param>
        /// <param name="labelId">Label GUID or null for default</param>
        /// <param name="justification">Justification text or null for default</param>
        /// <returns>True on success</returns>
        public bool ApplyLabelToFile(string filePath, string? labelId, string? justification)
        {
            try
            {
                ValidateFilePath(filePath);
                EnsureInitialized();

                var effectiveLabelId = string.IsNullOrEmpty(labelId) ? DefaultLabelId : labelId;
                var effectiveJustification = string.IsNullOrEmpty(justification) ? "Applied by TRIS at creation" : justification;

                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    throw new InvalidOperationException($"Label not found: {effectiveLabelId}");

                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = effectiveJustification
                };

                handler.SetLabel(label, options, new ProtectionSettings());
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                // Replace original file with labeled version
                if (result && File.Exists(outputPath))
                {
                    try
                    {
                        File.Delete(filePath);
                        File.Move(outputPath, filePath);
                        return true;
                    }
                    catch (Exception moveEx)
                    {
                        // Clean up temp file if move failed
                        try { File.Delete(outputPath); } catch { }
                        throw new InvalidOperationException($"Failed to replace original file: {moveEx.Message}", moveEx);
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                var errorDetails = $"ApplyLabel Error: {ex.GetType().Name} - {ex.Message}";
                if (ex.InnerException != null)
                    errorDetails += $" | Inner: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}";
                
                // Log to Windows Event Log for debugging
                try
                {
                    System.Diagnostics.EventLog.WriteEntry("Application", errorDetails, System.Diagnostics.EventLogEntryType.Error);
                }
                catch { /* Ignore logging errors */ }
                
                throw new Exception(errorDetails, ex);
            }
        }

        /// <summary>
        /// Get currently applied label GUID from file.
        /// </summary>
        /// <param name="filePath">Absolute path to file</param>
        /// <returns>Label GUID or empty string if no label applied</returns>
        public string GetAppliedLabelId(string filePath)
        {
            try
            {
                ValidateFilePath(filePath);
                EnsureInitialized();

                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, filePath, true).GetAwaiter().GetResult();
                var label = handler.Label;
                
                return label?.Label?.Id ?? "";
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get label from {filePath}: {ex.Message}", ex);
            }
        }

        private void ValidateFilePath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
            
            if (!Path.IsPathRooted(filePath))
                throw new ArgumentException("File path must be absolute", nameof(filePath));
            
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");
        }

        private void EnsureInitialized()
        {
            if (_initialized) return;

            lock (_lock)
            {
                if (_initialized) return;

                try
                {
                    // Configure TLS for older .NET Framework
                    System.Net.ServicePointManager.SecurityProtocol = 
                        System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11;
                    
                    // Ensure MIP native DLLs are on PATH - check common locations
                    var possiblePaths = new[]
                    {
                        Environment.GetEnvironmentVariable("MIP_NATIVE_PATH"),
                        @"C:\ColdFusion2023\cfusion\runtime\lib\MIP\x64",
                        @"C:\ColdFusion2021\cfusion\runtime\lib\MIP\x64",
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Information Protection SDK", "Bin", "x64")
                    };
                    
                    string? nativeDir = null;
                    foreach (var path in possiblePaths)
                    {
                        if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                        {
                            nativeDir = path;
                            break;
                        }
                    }
                    
                    if (nativeDir != null)
                    {
                        var cur = Environment.GetEnvironmentVariable("PATH") ?? "";
                        if (cur.IndexOf(nativeDir, StringComparison.OrdinalIgnoreCase) == -1)
                            Environment.SetEnvironmentVariable("PATH", nativeDir + ";" + cur, EnvironmentVariableTarget.Process);
                    }
                    
                    // Initialize MIP SDK
                    MIP.Initialize(MipComponent.File);

                    if (!Guid.TryParse(ClientId, out _))
                        throw new InvalidOperationException("CLIENT_ID must be a GUID (Azure AD Application (client) ID).");

                    // Create file profile
                    var authDelegate = new ConfidentialAuthDelegate();
                    var consentDelegate = new ConsentDelegate();
                    var appInfo = new ApplicationInfo 
                    { 
                        ApplicationId = ClientId, 
                        ApplicationName = "TRIS", 
                        ApplicationVersion = "1.0" 
                    };
                    var mipContext = MIP.CreateMipContext(appInfo, Path.GetTempPath(), LogLevel.Error, null, null);
                    var profileSettings = new FileProfileSettings(mipContext, CacheStorageType.OnDisk, consentDelegate);

                    var profile = MIP.LoadFileProfileAsync(profileSettings).GetAwaiter().GetResult();

                    // Create file engine
                    var identity = new Identity(UserIdentity);
                    var engineSettings = new FileEngineSettings(
                        "TRIS-MIP-Engine",
                        authDelegate,
                        "",
                        "en-US")
                    {
                        Identity = identity,
                        Cloud = Cloud.Commercial
                    };

                    _fileEngine = profile.AddEngineAsync(engineSettings).GetAwaiter().GetResult();
                    _initialized = true;
                }
                catch (Exception ex)
                {
                    var details = $"TenantId: {TenantId}, ClientId: {ClientId}, Error: {ex.Message}";
                    throw new InvalidOperationException($"Failed to initialize MIP SDK: {details}", ex);
                }
            }
        }

        internal static string GetTenantIdInternal() => TenantId;
        internal static string GetClientIdInternal() => ClientId;
        internal static string GetClientSecretInternal() => ClientSecret;
        
        /// <summary>
        /// Diagnostic method to check configuration status
        /// </summary>
        public static string GetConfigurationStatus()
        {
            return $"TenantId: '{TenantId}' (Length: {TenantId?.Length ?? 0}), " +
                   $"ClientId: '{ClientId}' (Length: {ClientId?.Length ?? 0}), " +
                   $"ClientSecret: '{(string.IsNullOrEmpty(ClientSecret) ? "EMPTY" : "SET")}' (Length: {ClientSecret?.Length ?? 0}), " +
                   $"DefaultLabelId: '{DefaultLabelId}' (Length: {DefaultLabelId?.Length ?? 0}), " +
                   $"UserIdentity: '{UserIdentity}' (Length: {UserIdentity?.Length ?? 0})";
        }
        
        /// <summary>
        /// Test method to validate configuration before initialization
        /// </summary>
        public static string ValidateConfiguration()
        {
            var errors = new List<string>();
            
            if (string.IsNullOrEmpty(TenantId))
                errors.Add("TenantId is null or empty");
            else if (!Guid.TryParse(TenantId, out _))
                errors.Add($"TenantId '{TenantId}' is not a valid GUID");
                
            if (string.IsNullOrEmpty(ClientId))
                errors.Add("ClientId is null or empty");
            else if (!Guid.TryParse(ClientId, out _))
                errors.Add($"ClientId '{ClientId}' is not a valid GUID");
                
            if (string.IsNullOrEmpty(ClientSecret))
                errors.Add("ClientSecret is null or empty");
                
            if (string.IsNullOrEmpty(DefaultLabelId))
                errors.Add("DefaultLabelId is null or empty");
            else if (!Guid.TryParse(DefaultLabelId, out _))
                errors.Add($"DefaultLabelId '{DefaultLabelId}' is not a valid GUID");
                
            return errors.Count == 0 ? "Configuration is valid" : $"Configuration errors: {string.Join(", ", errors)}";
        }
        
        /// <summary>
        /// Test method to check if MIP SDK can be initialized with current configuration
        /// </summary>
        public static string TestInitialization()
        {
            try
            {
                var configStatus = ValidateConfiguration();
                if (!configStatus.StartsWith("Configuration is valid"))
                    return $"FAILED: {configStatus}";
                    
                Instance.EnsureInitialized();
                return "SUCCESS: MIP SDK initialized successfully";
            }
            catch (Exception ex)
            {
                return $"FAILED: {ex.GetType().Name} - {ex.Message}";
            }
        }
        
        /// <summary>
        /// Debug method to test Word file labeling step by step
        /// </summary>
        public string DebugApplyWordLabel(string filePath)
        {
            var step = "Unknown";
            try
            {
                step = "1. Validate Word file";
                ValidateFilePath(filePath);
                if (!filePath.ToLowerInvariant().EndsWith(".docx"))
                    return "FAILED: File must be a Word document (.docx)";
                
                step = "2. Ensure initialized";
                EnsureInitialized();
                
                step = "3. Get effective label ID";
                var effectiveLabelId = DefaultLabelId;
                
                step = "4. Create file handler";
                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                
                step = "5. Get label by ID";
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    return $"FAILED at step {step}: Label not found: {effectiveLabelId}";
                
                step = "6. Create Word labeling options";
                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = "Applied by TRIS to Word document"
                };
                
                step = "7. Set label on handler";
                handler.SetLabel(label, options, new ProtectionSettings());
                
                step = "8. Commit changes";
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                step = "9. Replace original Word file";
                if (result && File.Exists(outputPath))
                {
                    File.Delete(filePath);
                    File.Move(outputPath, filePath);
                }
                
                return $"SUCCESS: Word label applied. Result: {result}";
            }
            catch (Exception ex)
            {
                var errorMsg = $"FAILED at {step}: {ex.GetType().Name} - {ex.Message}";
                if (ex.InnerException != null)
                    errorMsg += $" | Inner: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}";
                return errorMsg;
            }
        }

        /// <summary>
        /// Debug method to test Excel file labeling step by step
        /// </summary>
        public string DebugApplyExcelLabel(string filePath)
        {
            var step = "Unknown";
            try
            {
                step = "1. Validate Excel file";
                ValidateFilePath(filePath);
                if (!filePath.ToLowerInvariant().EndsWith(".xlsx"))
                    return "FAILED: File must be an Excel spreadsheet (.xlsx)";
                
                step = "2. Ensure initialized";
                EnsureInitialized();
                
                step = "3. Get effective label ID";
                var effectiveLabelId = DefaultLabelId;
                
                step = "4. Create file handler";
                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                
                step = "5. Get label by ID";
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    return $"FAILED at step {step}: Label not found: {effectiveLabelId}";
                
                step = "6. Create Excel labeling options";
                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = "Applied by TRIS to Excel spreadsheet"
                };
                
                step = "7. Set label on handler";
                handler.SetLabel(label, options, new ProtectionSettings());
                
                step = "8. Commit changes";
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                step = "9. Replace original Excel file";
                if (result && File.Exists(outputPath))
                {
                    var originalSize = new FileInfo(filePath).Length;
                    var tempSize = new FileInfo(outputPath).Length;
                    
                    File.Delete(filePath);
                    File.Move(outputPath, filePath);
                    
                    var finalSize = new FileInfo(filePath).Length;
                    return $"SUCCESS: Excel label applied. Result: {result}, Original: {originalSize}b, Temp: {tempSize}b, Final: {finalSize}b";
                }
                else if (!result)
                {
                    return $"FAILED: Commit returned false. TempFile exists: {File.Exists(outputPath)}";
                }
                else
                {
                    return $"FAILED: Temp file not created. Result: {result}";
                }
                
                return $"SUCCESS: Excel label applied. Result: {result}";
            }
            catch (Exception ex)
            {
                var errorMsg = $"FAILED at {step}: {ex.GetType().Name} - {ex.Message}";
                if (ex.InnerException != null)
                    errorMsg += $" | Inner: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}";
                return errorMsg;
            }
        }

        /// <summary>
        /// Test method to verify boolean returns work with ColdFusion
        /// </summary>
        public bool TestBooleanReturn(bool value) => value;

        /// <summary>
        /// Test what DebugApplyExcelLabel actually returns
        /// </summary>
        public string TestDebugResult(string filePath)
        {
            var result = DebugApplyExcelLabel(filePath);
            return $"Result: '{result}' | StartsWith SUCCESS: {result.StartsWith("SUCCESS: ")} | Length: {result.Length}";
        }

        /// <summary>
        /// Debug version of ApplyLabelToExcelFile that returns detailed string results
        /// </summary>
        public string ApplyLabelToExcelFileDebug(string filePath, string? labelId, string? justification)
        {
            try
            {
                ValidateFilePath(filePath);
                if (!filePath.ToLowerInvariant().EndsWith(".xlsx"))
                    return "FALSE: File must be an Excel spreadsheet (.xlsx)";
                    
                EnsureInitialized();

                var effectiveLabelId = string.IsNullOrEmpty(labelId) ? DefaultLabelId : labelId;
                var effectiveJustification = string.IsNullOrEmpty(justification) ? "Applied by TRIS to Excel spreadsheet" : justification;

                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    return $"FALSE: Label not found: {effectiveLabelId}";

                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = effectiveJustification
                };

                handler.SetLabel(label, options, new ProtectionSettings());
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                if (result && File.Exists(outputPath))
                {
                    try
                    {
                        File.Delete(filePath);
                        File.Move(outputPath, filePath);
                        return "TRUE: Excel labeling completed successfully";
                    }
                    catch (Exception moveEx)
                    {
                        try { File.Delete(outputPath); } catch { }
                        return $"FALSE: File replacement failed: {moveEx.Message}";
                    }
                }
                
                return $"FALSE: Commit result: {result}, TempFile exists: {File.Exists(outputPath)}";
            }
            catch (Exception ex)
            {
                return $"FALSE: Exception: {ex.Message}";
            }
        }

        /// <summary>
        /// Debug method to test generic file labeling step by step
        /// </summary>
        public string DebugApplyLabel(string filePath)
        {
            var step = "Unknown";
            try
            {
                step = "1. Validate file";
                ValidateFilePath(filePath);
                
                step = "2. Ensure initialized";
                EnsureInitialized();
                
                step = "3. Get effective label ID";
                var effectiveLabelId = DefaultLabelId;
                
                step = "4. Create file handler";
                var outputPath = filePath + ".tmp";
                var handler = _fileEngine!.CreateFileHandlerAsync(filePath, outputPath, true).GetAwaiter().GetResult();
                
                step = "5. Get label by ID";
                var label = _fileEngine.GetLabelById(effectiveLabelId!);
                
                if (label == null)
                    return $"FAILED at step {step}: Label not found: {effectiveLabelId}";
                
                step = "6. Create labeling options";
                var options = new LabelingOptions
                {
                    AssignmentMethod = AssignmentMethod.Standard,
                    JustificationMessage = "Applied by TRIS at creation"
                };
                
                step = "7. Set label on handler";
                handler.SetLabel(label, options, new ProtectionSettings());
                
                step = "8. Commit changes";
                var result = handler.CommitAsync(outputPath).GetAwaiter().GetResult();
                
                step = "9. Replace original file";
                if (result && File.Exists(outputPath))
                {
                    File.Delete(filePath);
                    File.Move(outputPath, filePath);
                }
                
                return $"SUCCESS: Label applied. Result: {result}";
            }
            catch (Exception ex)
            {
                var errorMsg = $"FAILED at {step}: {ex.GetType().Name} - {ex.Message}";
                if (ex.InnerException != null)
                    errorMsg += $" | Inner: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}";
                return errorMsg;
            }
        }
    }
}