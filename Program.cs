using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SolidWorksMetadataReader
{
    private static readonly string[] ValidExtensions = { ".sldprt", ".sldasm", ".slddrw" };
    private static SldWorks swApp = null;

    static void Main(string[] args)
    {
        Console.WriteLine("SolidWorks Metadata Reader");
        Console.WriteLine("==========================");
        Console.WriteLine("Enter file paths to read metadata, or type 'exit' to quit.\n");

        try
        {
            // Initialize SolidWorks once at startup
            InitializeSolidWorks();

            // Main interactive loop
            RunInteractiveLoop();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
        finally
        {
            // Ensure cleanup on exit
            CleanupSolidWorks();
            Console.WriteLine("\nApplication closed. Press any key to exit...");
            Console.ReadKey();
        }
    }

    private static void InitializeSolidWorks()
    {
        Console.WriteLine("Initializing SolidWorks connection...");
        try
        {
            swApp = CreateSolidWorksInstance();
            swApp.Visible = false; // Run in background
            Console.WriteLine("✓ SolidWorks connection established\n");
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to initialize SolidWorks: {ex.Message}", ex);
        }
    }

    private static void RunInteractiveLoop()
    {
        while (true)
        {
            Console.Write("Enter file path (or 'exit' to quit): ");
            string input = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter a file path or 'exit'.\n");
                continue;
            }

            if (input.Equals("exit", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Exiting application...");
                break;
            }

            try
            {
                var metadata = ReadMetadata(input);
                DisplayMetadata(input, metadata);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"Details: {ex.InnerException.Message}");
            }

            Console.WriteLine(); // Add spacing between operations
        }
    }

    private static void DisplayMetadata(string filePath, Dictionary<string, string> metadata)
    {
        Console.WriteLine($"\n--- Metadata for: {Path.GetFileName(filePath)} ---");

        if (metadata.Count == 0)
        {
            Console.WriteLine("No metadata found.");
            return;
        }

        // Group and display metadata in a organized way
        DisplayGroup("File Information", metadata, new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified" });
        DisplayGroup("Document Properties", metadata, new[] { "Title", "Author", "Subject", "Comments", "Keywords" });
        DisplayGroup("Custom Properties", metadata, key => !IsStandardProperty(key));

        Console.WriteLine($"Total properties found: {metadata.Count}");
    }

    private static void DisplayGroup(string groupName, Dictionary<string, string> metadata, string[] keys)
    {
        var found = new List<string>();
        foreach (string key in keys)
        {
            if (metadata.ContainsKey(key))
            {
                found.Add($"  {key}: {metadata[key]}");
            }
        }

        if (found.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            found.ForEach(Console.WriteLine);
        }
    }

    private static void DisplayGroup(string groupName, Dictionary<string, string> metadata, Func<string, bool> filter)
    {
        var customProps = new List<string>();
        foreach (var kvp in metadata)
        {
            if (filter(kvp.Key))
            {
                customProps.Add($"  {kvp.Key}: {kvp.Value}");
            }
        }

        if (customProps.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            customProps.ForEach(Console.WriteLine);
        }
    }

    private static bool IsStandardProperty(string key)
    {
        var standardProps = new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified",
                                   "Title", "Author", "Subject", "Comments", "Keywords" };
        return Array.IndexOf(standardProps, key) >= 0;
    }

    public static Dictionary<string, string> ReadMetadata(string filePath)
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
            CleanupDocument(swModel);
        }

        return metadata;
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

    private static SldWorks CreateSolidWorksInstance()
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
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not read custom properties - {ex.Message}");
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
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not read summary info - {ex.Message}");
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
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not read file info - {ex.Message}");
        }
    }

    private static void CleanupDocument(ModelDoc2 swModel)
    {
        if (swModel == null) return;

        try
        {
            swApp?.CloseDoc(swModel.GetTitle());
            Marshal.ReleaseComObject(swModel);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not close document - {ex.Message}");
        }
    }

    private static void CleanupSolidWorks()
    {
        if (swApp == null) return;

        try
        {
            Console.WriteLine("Closing SolidWorks connection...");
            swApp.ExitApp();
            Marshal.ReleaseComObject(swApp);
            swApp = null;

            // Force garbage collection to help with COM cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("✓ SolidWorks connection closed");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning during SolidWorks cleanup: {ex.Message}");
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