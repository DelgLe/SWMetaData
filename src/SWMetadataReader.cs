using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SWMetadataReader
{
    private static readonly string[] ValidExtensions = { ".sldprt", ".sldasm", ".slddrw" };

    public static Dictionary<string, string> ReadMetadata(SldWorks swApp, string filePath)
    {
        // Validate input
        ValidateFilePath(filePath);

        if (swApp == null)
            throw new InvalidOperationException("SolidWorks not initialized. Please restart the application.");

        ModelDoc2? swModel = null;
        var metadata = new Dictionary<string, string>();

        try
        {
            // Open document
            Logger.LogDocument($"Starting document read: {Path.GetFileName(filePath)}", filePath);
            swModel = OpenDocument(swApp, filePath);

            // Read all metadata
            ReadCustomProperties(swModel, metadata);
            ReadSummaryInfo(swModel, metadata);
            AddFileInfo(swModel, metadata, filePath);
            ReadConfigurations(swModel, metadata);
            ReadMaterialInfo(swModel, metadata);
            ReadComponentInfo(swModel, metadata);
            
            Logger.LogDocument($"Metadata read complete - {metadata.Count} properties", filePath);
        }
        catch (COMException comEx)
        {
            Logger.LogException(comEx, $"SolidWorks COM error - {Path.GetFileName(filePath)}");
            throw new Exception($"SolidWorks COM error: {comEx.Message} (HRESULT: 0x{comEx.HResult:X8})", comEx);
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, $"Error reading metadata - {Path.GetFileName(filePath)}");
            throw new Exception($"Error reading metadata: {ex.Message}", ex);
        }
        finally
        {
            // Close document but keep SolidWorks running for next file
            CleanupDocument(swApp, swModel);
        }

        return metadata;
    }
    private static void ReadConfigurations(ModelDoc2 swModel, Dictionary<string, string> metadata)
    {
        try
        {
            string[] configNames = (string[])swModel.GetConfigurationNames();
            if (configNames != null && configNames.Length > 0)
            {
                metadata["Configurations"] = string.Join(", ", configNames);
            }
        }
        catch (Exception)
        {
            // Ignore configuration errors
        }
    }

    private static void ReadMaterialInfo(ModelDoc2 swModel, Dictionary<string, string> metadata)
    {
        try
        {
            // Get material name for parts
            string materialName = swModel.MaterialIdName;
            if (!string.IsNullOrWhiteSpace(materialName))
            {
                metadata["Material"] = materialName;
            }
        }
        catch (Exception)
        {
            // Ignore material errors
        }
    }

    private static void ReadComponentInfo(ModelDoc2 swModel, Dictionary<string, string> metadata)
    {
        try
        {
            // Get component count for assemblies
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                AssemblyDoc swAssembly = (AssemblyDoc)swModel;
                int componentCount = swAssembly.GetComponentCount(false);
                metadata["ComponentCount"] = componentCount.ToString();
            }
        }
        catch (Exception)
        {
            // Ignore component count errors
        }
    }

    public static SldWorks CreateSolidWorksInstance()
    {
        try
        {
            var swAppInstance = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            if (swAppInstance == null)
                throw new InvalidOperationException("Failed to create SolidWorks instance");

            return swAppInstance;
        }
        catch (COMException)
        {
            throw new InvalidOperationException("Failed to connect to SolidWorks. Ensure SolidWorks is installed and properly registered.");
        }
    }

    public static void CleanupSolidWorks(SldWorks swApp)
    {
        if (swApp == null) return;

        try
        {
            swApp.ExitApp();
            Marshal.ReleaseComObject(swApp);
            // Force garbage collection to help with COM cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        catch (Exception)
        {
            // Ignore cleanup errors
        }
    }

    private static void ValidateFilePath(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty");

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (Array.IndexOf(ValidExtensions, extension) == -1)
            throw new ArgumentException($"Invalid file type. Supported types: {string.Join(", ", ValidExtensions)}");
    }

    private static ModelDoc2 OpenDocument(SldWorks swApp, string filePath)
    {
        swDocumentTypes_e docType = GetDocumentType(filePath);
        int errors = 0, warnings = 0;

        Logger.LogDocument($"Attempting to open {docType}: {Path.GetFileName(filePath)}", filePath);

        ModelDoc2 swModel = swApp.OpenDoc6(filePath, (int)docType,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

        if (swModel == null)
        {
            // For failed opens, especially assemblies, SolidWorks might still have 
            // partially loaded documents that need cleanup
            Logger.LogWarning($"Document failed to open: {Path.GetFileName(filePath)} (Errors: {errors}, Warnings: {warnings})", Path.GetFileName(filePath));
            
            // Log hanging documents before cleanup
            Logger.LogHangingDocuments(swApp, "Failed document open - ");
            
            // Force cleanup of any partially loaded documents
            try
            {
                Logger.LogInfo("Cleaning up any partially loaded documents", Path.GetFileName(filePath));
                swApp.CloseAllDocuments(true); // Force close any hanging documents
                
                // Additional cleanup - sometimes SolidWorks needs a moment
                System.Threading.Thread.Sleep(500);
                
                // Force garbage collection to help with COM cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                Logger.LogInfo("Cleanup completed for failed document open", Path.GetFileName(filePath));
            }
            catch (Exception cleanupEx)
            {
                Logger.LogWarning($"Cleanup after failed open had issues: {cleanupEx.Message}", Path.GetFileName(filePath));
            }

            string errorMsg = $"Failed to open document: {Path.GetFileName(filePath)}";
            if (errors != 0) errorMsg += $" (Errors: {errors})";
            if (warnings != 0) errorMsg += $" (Warnings: {warnings})";
            throw new InvalidOperationException(errorMsg);
        }

        Logger.LogDocument($"Successfully opened: {Path.GetFileName(filePath)}", filePath);
        return swModel;
    }

    private static void ReadCustomProperties(ModelDoc2 swModel, Dictionary<string, string> metadata)
    {
        try
        {
            CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
            if (propMgr == null) return;

            object propNamesObj = propMgr.GetNames();
            if (propNamesObj == null) return;

            object[] propNames = (object[])propNamesObj;
            foreach (string propName in propNames)
            {
                if (string.IsNullOrEmpty(propName)) continue;

                string propValue = "";
                string resolvedValue = "";

                propMgr.Get2(propName, out propValue, out resolvedValue);
                metadata[propName] = !string.IsNullOrEmpty(resolvedValue) ? resolvedValue : propValue ?? "";
            }
        }
        catch (Exception)
        {
            // Ignore custom property errors
        }
    }

    private static void ReadSummaryInfo(ModelDoc2 swModel, Dictionary<string, string> metadata)
    {
        try
        {
            var summaryFields = new[]
            {
                (swSummInfoField_e.swSumInfoTitle, "Title"),
                (swSummInfoField_e.swSumInfoAuthor, "Author"),
                (swSummInfoField_e.swSumInfoComment, "Comments"),
                (swSummInfoField_e.swSumInfoSubject, "Subject"),
                (swSummInfoField_e.swSumInfoKeywords, "Keywords")
            };

            foreach (var (field, name) in summaryFields)
            {
                try
                {
                    string value = swModel.SummaryInfo[(int)field];
                    if (!string.IsNullOrEmpty(value))
                        metadata[name] = value;
                }
                catch
                {
                    // Skip individual fields that can't be read
                }
            }
        }
        catch (Exception)
        {
            // Ignore summary info errors
        }
    }

    private static void AddFileInfo(ModelDoc2 swModel, Dictionary<string, string> metadata, string originalPath)
    {
        try
        {
            string modelPath = swModel.GetPathName();
            metadata["FilePath"] = !string.IsNullOrEmpty(modelPath) ? modelPath : originalPath;
            metadata["FileName"] = Path.GetFileName(metadata["FilePath"]);
            metadata["DocumentType"] = GetDocumentType(originalPath).ToString();
            metadata["FileSize"] = new FileInfo(originalPath).Length.ToString() + " bytes";
            metadata["LastModified"] = File.GetLastWriteTime(originalPath).ToString("yyyy-MM-dd HH:mm:ss");
        }
        catch (Exception)
        {
            // Ignore file info errors
        }
    }

    private static void CleanupDocument(SldWorks swApp, ModelDoc2? swModel)
    {
        if (swModel == null) return;

        try
        {
            // Get the full path name which is more reliable for closing documents
            string pathName = swModel.GetPathName();
            
            // Try multiple methods to close the document properly
            bool docClosed = false;
            
            if (!string.IsNullOrEmpty(pathName))
            {
                // Method 1: Close by path name (most reliable)
                try
                {
                    swApp.CloseDoc(pathName);
                    docClosed = true;
                    Logger.LogDocument($"Document closed successfully: {Path.GetFileName(pathName)}", pathName);
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to close document by path: {ex.Message}", Path.GetFileName(pathName));
                }
            }
            
            // Method 2: If path method failed, try by title
            if (!docClosed)
            {
                try
                {
                    string title = swModel.GetTitle();
                    if (!string.IsNullOrEmpty(title))
                    {
                        swApp.CloseDoc(title);
                        docClosed = true;
                        Logger.LogDocument($"Document closed by title: {title}", pathName);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to close document by title: {ex.Message}", pathName);
                }
            }
            
            // Method 3: Force close all documents if previous methods failed
            if (!docClosed)
            {
                try
                {
                    Logger.LogWarning("Attempting to close all documents as fallback", pathName);
                    swApp.CloseAllDocuments(true); // true = force close without saving
                    Logger.LogInfo("All documents closed forcefully", pathName);
                    docClosed = true;
                }
                catch (Exception ex)
                {
                    Logger.LogError($"Failed to force close documents: {ex.Message}", pathName);
                }
            }
            
            // Always release COM object regardless of close success
            Marshal.ReleaseComObject(swModel);
            
            // Additional cleanup - force garbage collection to help release COM resources
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
            if (!docClosed)
            {
                Logger.LogWarning("Document may not have been properly closed - check SolidWorks taskbar", pathName);
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error during document cleanup: {ex.Message}", swModel?.GetPathName());
            // Still try to release COM object
            try
            {
                Marshal.ReleaseComObject(swModel);
            }
            catch
            {
                // Final fallback - ignore any errors
            }
        }
    }

    public static swDocumentTypes_e GetDocumentType(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLowerInvariant();

        return extension switch
        {
            ".sldprt" => swDocumentTypes_e.swDocPART,
            ".sldasm" => swDocumentTypes_e.swDocASSEMBLY,
            ".slddrw" => swDocumentTypes_e.swDocDRAWING,
            _ => throw new ArgumentException($"Unsupported file extension: {extension}")
        };
    }
}