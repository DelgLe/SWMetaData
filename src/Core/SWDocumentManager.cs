using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

/// <summary>
/// Utility class for managing SolidWorks document lifecycle and cleanup operations
/// </summary>
public static class SWDocumentManager
{
    /// <summary>
    /// Properly close a SolidWorks document with multiple fallback methods
    /// </summary>
    public static void CloseDocumentSafely(SldWorks swApp, ModelDoc2? swModel)
    {
        if (swModel == null) return;

        try
        {
            string pathName = swModel.GetPathName();
            if (!string.IsNullOrEmpty(pathName))
            {
                swApp.CloseDoc(pathName);
                Logger.LogInfo($"Document closed: {Path.GetFileName(pathName)}");
            }
            else
            {
                string title = swModel.GetTitle();
                swApp.CloseDoc(title);
                Logger.LogInfo($"Document closed by title: {title}");
            }
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Normal document close failed: {ex.Message}");
            try
            {
                swApp.CloseAllDocuments(true);
                Logger.LogInfo("Forced close of all documents");
            }
            catch (Exception forceEx)
            {
                Logger.LogError($"Force close failed: {forceEx.Message}");
            }
        }
        finally
        {
            try
            {
                Marshal.ReleaseComObject(swModel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception releaseEx)
            {
                Logger.LogError($"COM object release failed: {releaseEx.Message}");
            }
        }
    }

    /// <summary>
    /// Force cleanup of hanging SolidWorks documents with comprehensive logging
    /// </summary>
    public static void ForceCleanupHangingDocuments(SldWorks swApp, string context = "")
    {
        try
        {
            int docCount = swApp.GetDocumentCount();
            if (docCount > 0)
            {
                Logger.LogRuntime("ForceCleanupHangingDocuments started", context, $"Document count: {docCount}");
                
                // Log all hanging documents before cleanup
                Logger.LogHangingDocuments(swApp, context);
                
                swApp.CloseAllDocuments(true); // Force close without saving
                
                // Give SolidWorks a moment to clean up
                Thread.Sleep(500);
                
                // Force garbage collection to help with COM cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Log the cleanup result
                Logger.LogCloseAllDocumentsResult(swApp, docCount, context);
                
                // Verify cleanup
                int remainingDocs = swApp.GetDocumentCount();
                if (remainingDocs == 0)
                {
                    Logger.LogInfo($"{context}Successfully cleaned up hanging documents", context);
                }
                else
                {
                    Logger.LogWarning($"{context}Warning: {remainingDocs} document(s) still hanging after cleanup", context);
                }
            }
            else
            {
                Logger.LogInfo("No hanging documents found", context);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, $"ForceCleanupHangingDocuments - {context}");
        }
    }

    /// <summary>
    /// Get the current count of open documents in SolidWorks
    /// </summary>
    public static int GetOpenDocumentCount(SldWorks swApp)
    {
        try
        {
            return swApp.GetDocumentCount();
        }
        catch (Exception ex)
        {
            Logger.LogError($"Failed to get document count: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Check if there are hanging documents and optionally clean them up
    /// </summary>
    public static bool HasHangingDocuments(SldWorks swApp)
    {
        return GetOpenDocumentCount(swApp) > 0;
    }

    /// <summary>
    /// Create a new SolidWorks application instance
    /// </summary>
    public static SldWorks CreateSolidWorksInstance()
    {
        try
        {
            var swAppInstance = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            if (swAppInstance == null)
                throw new InvalidOperationException("Failed to create SolidWorks instance");

            return swAppInstance;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            throw new InvalidOperationException("Failed to connect to SolidWorks. Ensure SolidWorks is installed and properly registered.");
        }
    }

    /// <summary>
    /// Properly cleanup and exit SolidWorks application
    /// </summary>
    public static void CleanupSolidWorks(SldWorks swApp)
    {
        if (swApp == null) return;

        try
        {
            // Close all documents first
            ForceCleanupHangingDocuments(swApp, "Application shutdown - ");
            
            // Exit the application
            swApp.ExitApp();
            
            // Release COM object
            Marshal.ReleaseComObject(swApp);
            
            // Force garbage collection
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
            Logger.LogInfo("SolidWorks application cleaned up successfully");
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error during SolidWorks cleanup: {ex.Message}");
            // Continue with cleanup even if errors occur
        }
    }

    /// <summary>
    /// Get the document type from file extension
    /// </summary>
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