using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SolidWorksMetadataReader
{
    static void Main(string[] args)
    {
        string filePath = @"M:\CAD\Products\KCV\Parts\107926.SLDPRT";
        try
        {
            var metadata = ReadMetadata(filePath);
            foreach (var kvp in metadata)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        Console.ReadKey();
    }

    public static Dictionary<string, string> ReadMetadata(string filePath)
    {
        SldWorks swApp = null;
        ModelDoc2 swModel = null;
        var metadata = new Dictionary<string, string>();

        try
        {
            // Connect to SolidWorks
            swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            if (swApp == null)
            {
                throw new InvalidOperationException("Failed to connect to SolidWorks. Make sure SolidWorks is installed.");
            }

            swApp.Visible = false; // Run in background

            // Determine document type based on file extension
            swDocumentTypes_e docType = GetDocumentType(filePath);

            // Open document
            int errors = 0, warnings = 0;
            swModel = swApp.OpenDoc6(filePath, (int)docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            if (swModel == null)
            {
                throw new InvalidOperationException($"Failed to open document. Errors: {errors}, Warnings: {warnings}");
            }

            // Get custom properties
            CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
            if (propMgr != null)
            {
                object propNamesObj = propMgr.GetNames();

                if (propNamesObj != null)
                {
                    object[] propNames = (object[])propNamesObj;

                    if (propNames.Length > 0)
                    {
                        foreach (string propName in propNames)
                        {
                            string propValue = "";
                            string resolvedValue = "";

                            // Try to get the resolved value first
                            propMgr.Get2(propName, out propValue, out resolvedValue);

                            // Use resolved value if available, otherwise use the property value
                            metadata[propName] = !string.IsNullOrEmpty(resolvedValue) ? resolvedValue : propValue;
                        }
                    }
                }
            }

            // Get summary info (document properties)
            try
            {
                string title = swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoTitle];
                if (!string.IsNullOrEmpty(title)) metadata["Title"] = title;

                string author = swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoAuthor];
                if (!string.IsNullOrEmpty(author)) metadata["Author"] = author;

                string comments = swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoComment];
                if (!string.IsNullOrEmpty(comments)) metadata["Comments"] = comments;

                string subject = swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoSubject];
                if (!string.IsNullOrEmpty(subject)) metadata["Subject"] = subject;

                string keywords = swModel.SummaryInfo[(int)swSummInfoField_e.swSumInfoKeywords];
                if (!string.IsNullOrEmpty(keywords)) metadata["Keywords"] = keywords;
            }
            catch (Exception ex)
            {
                // Summary info might not be available for all documents
                Console.WriteLine($"Warning: Could not read summary info - {ex.Message}");
            }

            // Add file information
            metadata["FilePath"] = swModel.GetPathName();
            metadata["FileName"] = System.IO.Path.GetFileName(swModel.GetPathName());
            metadata["DocumentType"] = docType.ToString();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error reading SolidWorks metadata: {ex.Message}", ex);
        }
        finally
        {
            // Cleanup - Close document first, then release COM objects
            try
            {
                if (swModel != null)
                {
                    swApp?.CloseDoc(swModel.GetTitle());
                    Marshal.ReleaseComObject(swModel);
                }

                if (swApp != null)
                {
                    swApp.ExitApp();
                    Marshal.ReleaseComObject(swApp);
                }
            }
            catch (Exception cleanupEx)
            {
                Console.WriteLine($"Warning during cleanup: {cleanupEx.Message}");
            }
        }

        return metadata;
    }

    private static swDocumentTypes_e GetDocumentType(string filePath)
    {
        string extension = System.IO.Path.GetExtension(filePath).ToLower();

        switch (extension)
        {
            case ".sldprt":
                return swDocumentTypes_e.swDocPART;
            case ".sldasm":
                return swDocumentTypes_e.swDocASSEMBLY;
            case ".slddrw":
                return swDocumentTypes_e.swDocDRAWING;
            default:
                return swDocumentTypes_e.swDocPART; // Default fallback
        }
    }
}