using System;
using System.IO;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

/// <summary>
/// Logger class for tracking application runtime events and SolidWorks document operations
/// </summary>
public static class Logger
{
    private static string _logFilePath = "C:\\Users\\alexanderd\\source\\repos\\DelgLe\\SWMetaData\\src\\sw.log";
    private static readonly object _lockObject = new object();
    private static bool _isInitialized = false;

    /// <summary>
    /// Initialize the logger with configuration or custom log directory
    /// </summary>
    /// <param name="config">Application configuration containing log directory</param>
    /// <param name="logDirectory">Custom log directory (overrides config if provided)</param>
    public static void Initialize(AppConfig? config = null, string? logDirectory = null)
    {
        try
        {
            // Determine log directory from config, parameter, or default
            if (!string.IsNullOrEmpty(logDirectory))
            {
                // Use provided parameter
            }
            else if (config?.LogDirectory != null)
            {
                logDirectory = config.LogDirectory;
            }
            else
            {
                logDirectory = Path.Combine(System.Environment.CurrentDirectory, "logs");
            }

            // Create logs directory if it doesn't exist
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            // Create log file with timestamp
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            _logFilePath = Path.Combine(logDirectory, $"swmetadata_{timestamp}.log");

            _isInitialized = true;

            // Write initial log entry
            LogRuntime("Logger initialized", $"Log file: {_logFilePath}");
            if (config?.LogDirectory != null)
            {
                LogRuntime("Log directory loaded from config", config.LogDirectory);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to initialize logger: {ex.Message}");
            _isInitialized = false;
        }
    }

    /// <summary>
    /// Log an information message
    /// </summary>
    public static void LogInfo(string message, string? details = null)
    {
        WriteLog("INFO", message, details);
    }

        /// <summary>
    /// Log an information message
    /// </summary>
    public static void WriteAndLogUserMessage(string message, string? details = null)
    {
        WriteLog("INFO", message, details, true, false);
    }

    /// <summary>
    /// Log a warning message
    /// </summary>
    public static void LogWarning(string message, string? details = null)
    {
        WriteLog("WARN", message, details);
    }

    /// <summary>
    /// Log an error message
    /// </summary>
    public static void LogError(string message, string? details = null)
    {
        WriteLog("ERROR", message, details);
    }

    /// <summary>
    /// Log application runtime events
    /// </summary>
    public static void LogRuntime(string operation, string? context = null, string? details = null)
    {
        string message = $"RUNTIME: {operation}";
        if (!string.IsNullOrEmpty(context))
        {
            message += $" [{context}]";
        }
        WriteLog("INFO", message, details, false); // Silent log (file only)
    }

    /// <summary>
    /// Log document operations with file information
    /// </summary>
    public static void LogDocument(string operation, string? filePath = null, string? documentType = null, string? details = null)
    {
        var sb = new StringBuilder();
        sb.Append($"DOC: {operation}");
        
        
        if (!string.IsNullOrEmpty(documentType))
        {
            sb.Append($" ({documentType})");
        }

        string? fullDetails = details;
        if (!string.IsNullOrEmpty(filePath))
        {
            fullDetails = string.IsNullOrEmpty(details) ? $"Path: {filePath}" : $"Path: {filePath} {details}";
        }

        WriteLog("INFO", sb.ToString(), fullDetails);
    }

    /// <summary>
    /// Log hanging documents before CloseAllDocuments is called
    /// </summary>
    public static void LogHangingDocuments(SldWorks swApp, string context = "")
    {
        try
        {
            int docCount = swApp.GetDocumentCount();
            
            if (docCount == 0)
            {
                LogInfo($"No hanging documents found", context);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Found {docCount} hanging document(s) - Context: {context}");
            sb.AppendLine("Hanging Documents List:");

            try
            {
                // Get documents array directly
                object docsObj = swApp.GetDocuments();
                
                if (docsObj is object[] docsArray && docsArray.Length > 0)
                {
                    for (int i = 0; i < docsArray.Length; i++)
                    {
                        if (docsArray[i] is ModelDoc2 doc)
                        {
                            try
                            {
                                string path = doc.GetPathName() ?? "Unknown Path";
                                string title = doc.GetTitle() ?? "Unknown Title";
                                swDocumentTypes_e type = (swDocumentTypes_e)doc.GetType();

                                sb.AppendLine($"  {i + 1}. [{type}] {title}");
                                sb.AppendLine($"      Path: {path}");

                                // Log individual document as well
                                LogDocument("HANGING", path, type.ToString(), $"Title: {title}, Context: {context}");
                            }
                            catch (Exception docEx)
                            {
                                sb.AppendLine($"  {i + 1}. [ERROR] Could not read document info: {docEx.Message}");
                            }
                        }
                        else
                        {
                            sb.AppendLine($"  {i + 1}. [ERROR] Document is not a valid ModelDoc2 object");
                        }
                    }
                }
                else
                {
                    sb.AppendLine("  No documents found in GetDocuments() array");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error enumerating documents: {ex.Message}");
            }

            LogWarning("Hanging documents detected before cleanup", sb.ToString());
        }
        catch (Exception ex)
        {
            LogError("Failed to log hanging documents", $"Context: {context}, Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Log the result of CloseAllDocuments operation
    /// </summary>
    public static void LogCloseAllDocumentsResult(SldWorks swApp, int documentCountBefore, string context = "")
    {
        try
        {
            int documentCountAfter = swApp.GetDocumentCount();
            int closedCount = documentCountBefore - documentCountAfter;

            string message = $"CloseAllDocuments executed - Context: {context}";
            string details = $"Documents before: {documentCountBefore}\n" +
                           $"Documents after: {documentCountAfter}\n" +
                           $"Documents closed: {closedCount}";

            if (documentCountAfter == 0)
            {
                LogInfo(message + " - SUCCESS: All documents closed", details);
            }
            else
            {
                LogWarning(message + $" - PARTIAL: {documentCountAfter} documents still hanging", details);
                
                // Log the remaining hanging documents
                LogHangingDocuments(swApp, $"After CloseAllDocuments - {context}");
            }
        }
        catch (Exception ex)
        {
            LogError("Failed to log CloseAllDocuments result", $"Context: {context}, Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Log exceptions with full details
    /// </summary>
    public static void LogException(Exception ex, string context = "")
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Exception in context: {context}");
        sb.AppendLine($"Type: {ex.GetType().Name}");
        sb.AppendLine($"Message: {ex.Message}");
        sb.AppendLine($"Stack Trace: {ex.StackTrace}");
        
        if (ex.InnerException != null)
        {
            sb.AppendLine("Inner Exception:");
            sb.AppendLine($"  Type: {ex.InnerException.GetType().Name}");
            sb.AppendLine($"  Message: {ex.InnerException.Message}");
        }

        LogError("Exception occurred", sb.ToString());
    }

    /// <summary>
    /// Write a log entry to file and console
    /// </summary>
    /// <param name="level">Log level (INFO, WARN, ERROR, etc.)</param>
    /// <param name="message">Main log message</param>
    /// <param name="details">Optional detailed information</param>
    /// <param name="includeConsole">If true, only writes to file (no console output)</param>
    private static void WriteLog(string level, string message, string? details = null, bool includeConsole = true, bool includeTimestamp = true)
    {
        try
        {
            lock (_lockObject)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var log_sb = new StringBuilder();
                var messsage_sb = new StringBuilder();

                log_sb.AppendLine($"[{timestamp}] [{level}] {message}");
                messsage_sb.AppendLine($"{message}");

                string formattedOutput = ""; 

                if (includeTimestamp)
                {
                    formattedOutput = log_sb.ToString();
                }
                else
                {
                    formattedOutput = messsage_sb.ToString();
                }

                // Write to console (unless silent logging is requested)
                if (includeConsole)
                {
                    Console.Write(formattedOutput);
                }

                // Write to file (only if initialized)
                if (_isInitialized)
                {
                    File.AppendAllText(_logFilePath, log_sb.ToString(), Encoding.UTF8);
                }
            }
        }
        catch (Exception ex)
        {
            // Fallback logging if the main logging fails
            string fallbackMessage = $"[LOGGING ERROR] {ex.Message}\n[{level}] {message}";
            if (!string.IsNullOrEmpty(details))
            {
                fallbackMessage += $"\nDetails: {details}";
            }
            Console.WriteLine(fallbackMessage);
        }
    }

    /// <summary>
    /// Get the current log file path
    /// </summary>
    public static string GetLogFilePath()
    {
        return _logFilePath;
    }

    /// <summary>
    /// Check if logger is initialized
    /// </summary>
    public static bool IsInitialized => _isInitialized;
}