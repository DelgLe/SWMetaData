using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SolidWorksMetadataReader
{
    private static readonly string[] ValidExtensions = { ".sldprt", ".sldasm", ".slddrw" };

    public static Dictionary<string, string> ReadMetadata(SldWorks swApp, string filePath)
    {
        // Validate input
        ValidateFilePath(filePath);

        if (swApp == null)
            throw new InvalidOperationException("SolidWorks not initialized. Please restart the application.");

        ModelDoc2 swModel = null;
        var metadata = new Dictionary<string, string>();

        try
        {
            // Open document
            swModel = OpenDocument(swApp, filePath);

            // Read all metadata
            ReadCustomProperties(swModel, metadata);
            ReadSummaryInfo(swModel, metadata);
            AddFileInfo(swModel, metadata, filePath);
            ReadConfigurations(swModel, metadata);
        }
        catch (COMException comEx)
        {
            throw new Exception($"SolidWorks COM error: {comEx.Message} (HRESULT: 0x{comEx.HResult:X8})", comEx);
        }
        catch (Exception ex)
        {
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
            var configMgr = swModel.ConfigurationManager;
            if (configMgr == null) return;

            var configsObj = configMgr.GetType().InvokeMember("Configurations", System.Reflection.BindingFlags.GetProperty, null, configMgr, null);
            if (configsObj is System.Collections.IEnumerable configsEnum)
            {
                var configNames = new List<string>();
                foreach (object configObj in configsEnum)
                {
                    if (configObj is SolidWorks.Interop.sldworks.Configuration config)
                    {
                        string name = config.Name;
                        if (!string.IsNullOrEmpty(name))
                            configNames.Add(name);
                    }
                }
                if (configNames.Count > 0)
                {
                    metadata["Configurations"] = string.Join(", ", configNames);
                }
            }
        }
        catch (Exception)
        {
            // Ignore configuration errors
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

        ModelDoc2 swModel = swApp.OpenDoc6(filePath, (int)docType,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

        if (swModel == null)
        {
            string errorMsg = $"Failed to open document: {Path.GetFileName(filePath)}";
            if (errors != 0) errorMsg += $" (Errors: {errors})";
            if (warnings != 0) errorMsg += $" (Warnings: {warnings})";
            throw new InvalidOperationException(errorMsg);
        }

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

    private static void CleanupDocument(SldWorks swApp, ModelDoc2 swModel)
    {
        if (swModel == null) return;

        try
        {
            swApp?.CloseDoc(swModel.GetTitle());
            Marshal.ReleaseComObject(swModel);
        }
        catch (Exception)
        {
            // Ignore document cleanup errors
        }
    }

    private static swDocumentTypes_e GetDocumentType(string filePath)
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