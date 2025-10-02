using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class CommandInterface
{
    private static string? _databasePath = null;
    private static AppConfig? _config = null;

    /// <summary>
    /// Properly close a SolidWorks document with multiple fallback methods
    /// </summary>
    private static void CloseDocumentSafely(SldWorks swApp, ModelDoc2? swModel)
    {
        // Even if swModel is null, we might need to clean up hanging documents
        if (swModel == null) 
        {
            // Check if there are any open documents that might be hanging
            try
            {
                // Get count of currently open documents
                int docCount = swApp.GetDocumentCount();
                if (docCount > 0)
                {
                    Logger.LogWarning($"Found {docCount} potentially hanging documents, cleaning up", "InitialCleanup");
                    swApp.CloseAllDocuments(true);
                    System.Threading.Thread.Sleep(200);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"Warning during hanging document cleanup: {ex.Message}", "InitialCleanup");
            }
            return;
        }

        try
        {
            // Method 1: Try to close by path (most reliable)
            string pathName = swModel.GetPathName();
            if (!string.IsNullOrEmpty(pathName))
            {
                try
                {
                    swApp.CloseDoc(pathName);
                    Logger.LogDocument($"Document closed: {Path.GetFileName(pathName)}", pathName);
                    return;
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to close by path: {ex.Message}", Path.GetFileName(pathName));
                }
            }

            // Method 2: Try to close by title
                try
                {
                    string title = swModel.GetTitle();
                    if (!string.IsNullOrEmpty(title))
                    {
                        swApp.CloseDoc(title);
                        Logger.LogDocument($"Document closed by title: {title}", pathName);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to close by title: {ex.Message}", pathName);
                }            // Method 3: Force close all documents as last resort
            Logger.LogWarning("Forcing close of all documents", pathName);
            swApp.CloseAllDocuments(true);
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error closing document: {ex.Message}", swModel?.GetPathName());
        }
    }

    /// <summary>
    /// Force cleanup of hanging documents, especially useful for assemblies that fail to open
    /// </summary>
    private static void ForceCleanupHangingDocuments(SldWorks swApp, string context = "")
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
                System.Threading.Thread.Sleep(500);
                
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

    public static void RunInteractiveLoop(SldWorks swApp)
    {
        // Get current configuration (already loaded in Main)
        _config = ConfigManager.GetCurrentConfig();
        Logger.LogRuntime("Application started", "RunInteractiveLoop");
        
        if (!string.IsNullOrEmpty(_config.DatabasePath))
        {
            _databasePath = _config.DatabasePath;
            Logger.LogRuntime("Database path loaded from config", _databasePath);
        }

        // Show initial options

        while (true)
        {
            ShowMainMenu();
            Console.Write("\nEnter your choice: ");
            string input = Console.ReadLine()?.Trim() ?? "";

            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter a valid option.\n");
                ShowMainMenu();
                continue;
            }

            if (input.Equals("exit", StringComparison.OrdinalIgnoreCase) || input == "7")
            {
                Logger.LogRuntime("Application exit requested", "User input");
                break;
            }

            switch (input)
            {
                case "1":
                    ProcessSingleFile(swApp);
                    break;
                case "2":
                    SetupDatabase();
                    break;
                case "3":
                    ProcessFileWithDatabase(swApp);
                    break;
                case "4":
                    ManageTargetFiles();
                    break;
                case "5":
                    ProcessAllTargetFiles(swApp);
                    break;
                case "6":
                    ManageConfiguration();
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
                    break;
            }
        }
    }

    private static void ShowMainMenu()
    {
        Console.WriteLine("\n=== SolidWorks Metadata Reader ===");
        Console.WriteLine("1. Process single file (display only)");
        Console.WriteLine("2. Setup/Create database");
        Console.WriteLine("3. Process file and save to database");
        Console.WriteLine("4. Manage target files");
        Console.WriteLine("5. Process all target files (batch)");
        Console.WriteLine("6. Configuration settings");
        Console.WriteLine("7. Exit");
        Console.WriteLine($"Current database: {(_databasePath ?? "Not set")}");
        if (_config != null)
        {
            Console.WriteLine($"Config loaded: {(_config.DatabasePath != null ? "Yes" : "Default")}");
        }
    }

    private static void ProcessSingleFile(SldWorks swApp)
    {
        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
            DisplayMetadata(filePath, metadata);
            
            // Check if this is an assembly and offer BOM option
            if (metadata.TryGetValue("DocumentType", out var docType) && 
                docType.Contains("ASSEMBLY"))
            {
                OfferBomOption(swApp, filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"Details: {ex.InnerException.Message}");
        }
    }

    private static void SetupDatabase()
    {
        Console.WriteLine("\n=== Database Setup ===");
        
        // Check if config has a database path
        bool hasConfigDatabase = _config != null && !string.IsNullOrEmpty(_config.DatabasePath);
        
        if (hasConfigDatabase)
        {
            Console.WriteLine($"Configuration database path found: {_config!.DatabasePath}");
            Console.WriteLine("Choose an option:");
            Console.WriteLine("1. Use existing config database path (setup tables if needed)");
            Console.WriteLine("2. Create new database in current folder");
            Console.Write("Enter your choice (1-2): ");
            
            string choice = Console.ReadLine()?.Trim() ?? "";
            
            if (choice == "1")
            {
                SetupConfigDatabase();
                return;
            }
            else if (choice != "2")
            {
                Console.WriteLine("Invalid choice. Creating new database in current folder...");
            }
        }
        
        // Create new database in current folder
        SetupNewDatabase();
    }

    private static void SetupConfigDatabase()
    {
        if (_config == null || string.IsNullOrEmpty(_config.DatabasePath))
        {
            Console.WriteLine("No config database path available.");
            return;
        }

        try
        {
            string configDbPath = _config.DatabasePath;
            bool dbExists = System.IO.File.Exists(configDbPath);
            
            Console.WriteLine($"Setting up database at: {configDbPath}");
            
            if (dbExists)
            {
                Console.WriteLine("Database file already exists. Checking/creating tables...");
            }
            else
            {
                Console.WriteLine("Database file doesn't exist. Creating new database...");
                
                // Create directory if it doesn't exist
                string? directory = System.IO.Path.GetDirectoryName(configDbPath);
                if (!string.IsNullOrEmpty(directory) && !System.IO.Directory.Exists(directory))
                {
                    System.IO.Directory.CreateDirectory(directory);
                    Console.WriteLine($"Created directory: {directory}");
                }
            }
            
            // Create/initialize database and tables
            using (var dbManager = new SWDatabaseManager(configDbPath))
            {
                _databasePath = configDbPath;
                
                if (dbExists)
                {
                    Console.WriteLine("Database connected successfully!");
                    Console.WriteLine("Tables verified/created:");
                }
                else
                {
                    Console.WriteLine("Database created successfully!");
                    Console.WriteLine("Tables initialized:");
                }
                
                
                // Check if target_files has existing data
                int targetFileCount = dbManager.GetTargetFileCount();
                if (targetFileCount > 0)
                {
                    Logger.LogInfo($"Found {targetFileCount} existing target files in database.");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error setting up config database: {ex.Message}");
            _databasePath = null;
        }
    }

    private static void SetupNewDatabase()
    {
        Console.Write("Enter database file name (will be created in current folder): ");
        string dbName = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(dbName))
        {
            Logger.LogWarning("Database name cannot be empty.");
            return;
        }

        // Ensure .db extension
        if (!dbName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
        {
            dbName += ".db";
        }

        try
        {
            string newDbPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), dbName);
            
            if (System.IO.File.Exists(newDbPath))
            {
                Console.Write($"Database file '{dbName}' already exists. Overwrite? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() != "y")
                {
                    Console.WriteLine("Database setup cancelled.");
                    return;
                }
            }
            
            // Create database and initialize tables
            using (var dbManager = new SWDatabaseManager(newDbPath))
            {
                _databasePath = newDbPath;
                
                Logger.LogInfo($"Database created successfully: {newDbPath}");
                Logger.LogInfo("Tables initialized:");

            }
            
            // Ask if user wants to update config
            if (_config != null)
            {
                Console.Write("Update configuration to use this new database? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() == "y")
                {
                    _config.DatabasePath = newDbPath;
                    Console.WriteLine("Configuration updated. Use 'Configuration settings > Save current configuration' to persist changes.");
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error creating database: {ex.Message}");
            _databasePath = null;
        }
    }

    private static void ProcessFileWithDatabase(SldWorks swApp)
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Logger.LogWarning("No database configured. Please setup database first (option 2).");
            return;
        }

        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            Logger.LogInfo("Reading metadata from SolidWorks file...");
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
            
            // Display the metadata
            DisplayMetadata(filePath, metadata);

            // Save to database
            using var dbManager = new SWDatabaseManager(_databasePath);
            Console.WriteLine("Saving to database...");
            long documentId = dbManager.InsertDocumentMetadata(metadata);
            Console.WriteLine($"Document saved with ID: {documentId}");

            // Handle BOM for assemblies
            if (metadata.TryGetValue("DocumentType", out var docType) && 
                docType.Contains("ASSEMBLY"))
            {
                Console.WriteLine("\nAssembly detected. Processing BOM...");
                
                // Re-open document for BOM processing
                ModelDoc2? swModel = null;
                try
                {
                    int errors = 0, warnings = 0;
                    
                    swDocumentTypes_e documentType = SWMetadataReader.GetDocumentType(filePath);
                    Logger.LogInfo($"Opening {documentType} for BOM processing: {Path.GetFileName(filePath)}");
                    
                    swModel = swApp.OpenDoc6(filePath, (int)documentType,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                    if (swModel != null)
                    {
                        var bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                        if (bomItems.Count > 0)
                        {
                            dbManager.InsertBomItems(filePath, bomItems);
                            Logger.LogInfo($"Saved {bomItems.Count} BOM items to database");
                            
                            // Offer to display BOM
                            Console.Write("Display BOM? (y/n): ");
                            if (Console.ReadLine()?.Trim().ToLower() == "y")
                            {
                                DisplayBomFromList(bomItems, "Bill of Materials (from database)");
                            }
                        }
                        else
                        {
                            Logger.LogWarning("No BOM items found in assembly");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Could not re-open document for BOM processing (Errors: {errors}, Warnings: {warnings})");
                        
                        // Assembly files sometimes hang even when open fails - force cleanup
                        if (documentType == swDocumentTypes_e.swDocASSEMBLY)
                        {
                            ForceCleanupHangingDocuments(swApp, "Assembly failed to open - ");
                        }
                    }
                }
                finally
                {
                    CloseDocumentSafely(swApp, swModel);
                }
            }

            Console.WriteLine("File processing complete!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"Details: {ex.InnerException.Message}");
        }
    }

    public static void DisplayMetadata(string filePath, Dictionary<string, string> metadata)
    {
        Console.WriteLine($"\n--- Metadata for: {System.IO.Path.GetFileName(filePath)} ---");

        if (metadata.Count == 0)
        {
            Console.WriteLine("No metadata found.");
            return;
        }

        DisplayGroup("File Information", metadata, new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified" });
        DisplayGroup("Document Properties", metadata, new[] { "Title", "Author", "Subject", "Comments", "Keywords" });

        // Dedicated section for material and components
        if (metadata.ContainsKey("Material"))
        {
            Console.WriteLine($"\nMaterial: {metadata["Material"]}");
        }
        
        if (metadata.ContainsKey("ComponentCount"))
        {
            Console.WriteLine($"\nComponent Count: {metadata["ComponentCount"]}");
        }

        // Dedicated section for configurations (robust display)
        if (metadata.TryGetValue("Configurations", out var configs))
        {
            var configList = configs
                .Split(',')
                .Select(c => c.Trim())
                .Where(c => !string.IsNullOrEmpty(c))
                .ToList();
            Console.WriteLine("\nConfigurations:");
            if (configList.Count > 0)
            {
                foreach (var config in configList)
                {
                    Console.WriteLine($"  {config}");
                }
            }
            else
            {
                Console.WriteLine("  (none found)");
            }
        }

        DisplayGroup("Custom Properties", metadata, key => !IsStandardProperty(key) && key != "Configurations" && key != "Material" && key != "ComponentCount");

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

    public static void OfferBomOption(SldWorks swApp, string filePath)
    {
        try
        {
            Console.WriteLine("\n--- Assembly BOM Options ---");
            Console.WriteLine("Would you like to view the Bill of Materials (BOM) for this assembly?");
            Console.WriteLine("1. View unsuppressed components only (BOM)");
            Console.WriteLine("2. View all components (including suppressed)");
            Console.WriteLine("3. Skip BOM");
            Console.Write("Enter your choice (1-3): ");

            string choice = Console.ReadLine()?.Trim() ?? "";

            if (choice == "1" || choice == "2")
            {
                DisplayBom(swApp, filePath, choice == "2");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error offering BOM option: {ex.Message}");
        }
    }

    public static void DisplayBom(SldWorks swApp, string filePath, bool includeSupressed)
    {
        ModelDoc2? swModel = null;
        try
        {
            // Re-open the document (it should already be cached/quick to open)
            int errors = 0, warnings = 0;
            
            swDocumentTypes_e docType = SWMetadataReader.GetDocumentType(filePath);
            Logger.WriteAndLogUserMessage($"Opening {docType} for BOM display: {Path.GetFileName(filePath)}");
            
            swModel = swApp.OpenDoc6(filePath, (int)docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            if (swModel == null)
            {
                Console.WriteLine($"Error: Could not open assembly for BOM generation (Errors: {errors}, Warnings: {warnings}).");
                
                // Assembly files sometimes hang even when open fails - force cleanup
                if (docType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    ForceCleanupHangingDocuments(swApp, "Assembly failed to open - ");
                }
                return;
            }

            List<BomItem> bomItems;
            string title;

            if (includeSupressed)
            {
                bomItems = SWAssemblyTraverser.GetAllComponents(swModel, true);
                title = "Complete Component List (Including Suppressed)";
            }
            else
            {
                bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                title = "Bill of Materials (Unsuppressed Components Only)";
            }

            DisplayBomFromList(bomItems, title);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating BOM: {ex.Message}");
        }
        finally
        {
            // Always close the document, even if an exception occurred
            CloseDocumentSafely(swApp, swModel);
        }
    }

    private static void DisplayBomFromList(List<BomItem> bomItems, string title)
    {
        Console.WriteLine($"\n--- {title} ---");
        
        if (bomItems.Count == 0)
        {
            Console.WriteLine("No components found in the assembly.");
        }
        else
        {
            Console.WriteLine($"Total components: {bomItems.Count}");
            Console.WriteLine(new string('-', 80));

            foreach (var item in bomItems)
            {
                string indent = new string(' ', item.Level * 2);
                string quantity = item.Quantity > 1 ? $"({item.Quantity}x) " : "";
                string suppression = item.IsSuppressed ? " [SUPPRESSED]" : "";
                
                Console.WriteLine($"{indent}{quantity}{item.ComponentName}");
                Console.WriteLine($"{indent}  File: {item.FileName} | Config: {item.Configuration}{suppression}");
                
                if (item.Level == 0) // Add spacing between top-level components
                    Console.WriteLine();
            }
        }
    }

    private static void ManageTargetFiles()
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Console.WriteLine("No database configured. Please setup database first (option 2).");
            return;
        }

        while (true)
        {
            Console.WriteLine("\n=== Target Files Management ===");
            Console.WriteLine("1. View all target files");
            Console.WriteLine("2. Add target file");
            Console.WriteLine("3. Remove target file");
            Console.WriteLine("4. Validate target files (check if files exist)");
            Console.WriteLine("5. Back to main menu");
            Console.Write("Enter your choice: ");

            string choice = Console.ReadLine()?.Trim() ?? "";

            switch (choice)
            {
                case "1":
                    ViewTargetFiles();
                    break;
                case "2":
                    AddTargetFile();
                    break;
                case "3":
                    RemoveTargetFile();
                    break;
                case "4":
                    ValidateTargetFiles();
                    break;
                case "5":
                    //ShowMainMenu(); Should always show on return
                    return;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
                    break;
            }
        }
    }

    private static void ViewTargetFiles()
    {
        try
        {
            using var dbManager = new SWDatabaseManager(_databasePath!);
            var targetFiles = dbManager.GetTargetFiles();

            if (targetFiles.Count == 0)
            {
                Console.WriteLine("No target files found in database.");
                return;
            }

            Console.WriteLine($"\n=== Target Files ({targetFiles.Count} files) ===");
            Console.WriteLine(new string('-', 100));
            Console.WriteLine("{0,-3} {1,-15} {2,-30} {3,-40}", "ID", "EngID", "File Name", "File Path");
            Console.WriteLine(new string('-', 100));

            foreach (var file in targetFiles)
            {
                string engId = file.EngID ?? "N/A";
                string fileName = file.FileName ?? "N/A";
                string filePath = file.FilePath ?? "N/A";
                
                // Truncate long paths for display
                if (filePath.Length > 40)
                    filePath = "..." + filePath.Substring(filePath.Length - 37);

                Console.WriteLine("{0,-3} {1,-15} {2,-30} {3,-40}", 
                    file.TargetID, engId, fileName, filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error viewing target files: {ex.Message}");
        }
    }

    private static void AddTargetFile()
    {
        try
        {
            Console.Write("Enter file path: ");
            string filePath = Console.ReadLine()?.Trim() ?? "";

            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("File path cannot be empty.");
                return;
            }

            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine("Warning: File does not exist at the specified path.");
                Console.Write("Continue anyway? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() != "y")
                    return;
            }

            Console.Write("Enter Engineering ID (optional): ");
            string? engId = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(engId)) engId = null;

            Console.Write("Enter Drawing ID (optional): ");
            string? drawId = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(drawId)) drawId = null;

            Console.Write("Enter notes (optional): ");
            string? notes = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(notes)) notes = null;

            using var dbManager = new SWDatabaseManager(_databasePath!);
            
            var targetFile = new TargetFileInfo
            {
                EngID = engId,
                FileName = System.IO.Path.GetFileName(filePath),
                FilePath = filePath,
                DrawID = drawId,
                Notes = notes,
                SourceDirectory = System.IO.Path.GetDirectoryName(filePath),
                FolderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(filePath)),
                FileCount = 1
            };

            dbManager.AddTargetFile(targetFile);
            Console.WriteLine("Target file added successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding target file: {ex.Message}");
        }
    }

    private static void RemoveTargetFile()
    {
        try
        {
            using var dbManager = new SWDatabaseManager(_databasePath!);
            var targetFiles = dbManager.GetTargetFiles();

            if (targetFiles.Count == 0)
            {
                Console.WriteLine("No target files to remove.");
                return;
            }

            Console.WriteLine("\nTarget Files:");
            foreach (var file in targetFiles.Take(20)) // Show first 20
            {
                Console.WriteLine($"{file.TargetID}: {file.FileName} ({file.EngID ?? "No EngID"})");
            }

            if (targetFiles.Count > 20)
                Console.WriteLine($"... and {targetFiles.Count - 20} more files");

            Console.Write("Enter Target ID to remove: ");
            if (int.TryParse(Console.ReadLine(), out int targetId))
            {
                dbManager.RemoveTargetFile(targetId);
                Console.WriteLine("Target file removed successfully!");
            }
            else
            {
                Console.WriteLine("Invalid Target ID.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error removing target file: {ex.Message}");
        }
    }

    private static void ValidateTargetFiles()
    {
        try
        {
            using var dbManager = new SWDatabaseManager(_databasePath!);
            var allFiles = dbManager.GetTargetFiles();
            var validFiles = dbManager.GetValidTargetFiles();

            Logger.WriteAndLogUserMessage($"\nValidation Results:");
            Logger.WriteAndLogUserMessage($"Total target files: {allFiles.Count}");
            Logger.WriteAndLogUserMessage($"Valid files (exist on disk): {validFiles.Count}");
            Logger.WriteAndLogUserMessage($"Missing files: {allFiles.Count - validFiles.Count}");

            if (allFiles.Count != validFiles.Count)
            {
                Console.WriteLine("\nMissing files:");
                foreach (var file in allFiles)
                {
                    if (!string.IsNullOrEmpty(file.FilePath) && !System.IO.File.Exists(file.FilePath))
                    {
                        Console.WriteLine($"  ID {file.TargetID}: {file.FilePath}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error validating target files: {ex.Message}");
        }
    }

    private static void ProcessAllTargetFiles(SldWorks swApp)
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Logger.WriteAndLogUserMessage("No database configured. Please setup database first (option 2).");
            return;
        }

        try
        {
            Logger.LogRuntime("Batch processing started", "ProcessAllTargetFiles");
            using var dbManager = new SWDatabaseManager(_databasePath);
            var validFiles = dbManager.GetValidTargetFiles();

            if (validFiles.Count == 0)
            {
                Logger.LogInfo("No valid target files found to process", "ProcessAllTargetFiles");
                return;
            }

            Logger.LogInfo($"Found {validFiles.Count} valid target files to process", "ProcessAllTargetFiles");
            Console.Write("Continue with batch processing? (y/n): ");
            
            if (Console.ReadLine()?.Trim().ToLower() != "y")
            {
                Logger.LogInfo("Batch processing cancelled by user", "ProcessAllTargetFiles");
                return;
            }

            int processed = 0;
            int errors = 0;

            foreach (var targetFile in validFiles)
            {
                try
                {
                    Console.WriteLine($"\nProcessing {processed + 1}/{validFiles.Count}: {targetFile.FileName}");
                    
                    if (string.IsNullOrEmpty(targetFile.FilePath))
                    {
                        Console.WriteLine("Skipping - no file path");
                        errors++;
                        continue;
                    }
                    
                    // Check if file should be excluded (prevents hanging from feature templates, etc.)
                    if (SWMetadataReader.IsFileExcluded(targetFile.FilePath))
                    {
                        Logger.LogInfo($"Skipped excluded file: {targetFile.FileName}", Path.GetExtension(targetFile.FilePath));
                        errors++; // Count as error to track exclusions
                        continue;
                    }

                    var metadata = SWMetadataReader.ReadMetadata(swApp, targetFile.FilePath);
                    long documentId = dbManager.InsertDocumentMetadata(metadata);
                    
                    // Process BOM for assemblies
                    if (metadata.TryGetValue("DocumentType", out var docType) && 
                        docType.Contains("ASSEMBLY"))
                    {
                        ModelDoc2? swModel = null;
                        try
                        {
                            int swErrors = 0, warnings = 0;
                            var assemblyDocType = SWMetadataReader.GetDocumentType(targetFile.FilePath);
                            Console.WriteLine($"  Opening {assemblyDocType}: {Path.GetFileName(targetFile.FilePath)}");
                            
                            swModel = swApp.OpenDoc6(targetFile.FilePath, 
                                (int)assemblyDocType,
                                (int)SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent, 
                                "", ref swErrors, ref warnings);

                            if (swModel != null)
                            {
                                var bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                                if (bomItems.Count > 0)
                                {
                                    dbManager.InsertBomItems(targetFile.FilePath, bomItems);
                                    Console.WriteLine($"  Saved {bomItems.Count} BOM items");
                                }
                                else
                                {
                                    Console.WriteLine("  No BOM items found");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"  Could not open assembly (Errors: {swErrors}, Warnings: {warnings})");
                                
                                // Assembly files sometimes hang even when open fails - force cleanup
                                if (assemblyDocType == swDocumentTypes_e.swDocASSEMBLY)
                                {
                                    ForceCleanupHangingDocuments(swApp, "  Assembly failed to open - ");
                                }
                            }
                        }
                        catch (Exception bomEx)
                        {
                            Console.WriteLine($"  BOM processing error: {bomEx.Message}");
                        }
                        finally
                        {
                            CloseDocumentSafely(swApp, swModel);
                        }
                    }

                    processed++;
                    Console.WriteLine($"  Success! Document ID: {documentId}");
                    
                    // Periodic cleanup check every 10 files to prevent accumulation
                    if (processed % 10 == 0)
                    {
                        ForceCleanupHangingDocuments(swApp, $"  Periodic cleanup after {processed} files - ");
                        SWMetadataReader.PeriodicCleanup(processed, "BatchProcessing");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error: {ex.Message}");
                    errors++;
                    
                    // Force cleanup after errors to prevent hanging documents
                    ForceCleanupHangingDocuments(swApp, "  Error cleanup - ");
                    
                    // Also perform periodic cleanup on errors to prevent accumulation
                    SWMetadataReader.PeriodicCleanup(processed + errors, "ErrorHandling");
                }
            }

            Logger.WriteAndLogUserMessage($"\n=== Batch Processing Complete ===");
            Logger.WriteAndLogUserMessage($"Successfully processed: {processed}");
            Logger.WriteAndLogUserMessage($"Errors: {errors}");
            Logger.WriteAndLogUserMessage($"Total: {validFiles.Count}");

            Logger.LogInfo($"Batch processing complete - Processed: {processed}, Errors: {errors}, Total: {validFiles.Count}", "ProcessAllTargetFiles");
            
            // Log excluded file types for reference
            var excludedTypes = SWMetadataReader.GetExcludedExtensions();
            Logger.LogInfo($"Excluded file types: {string.Join(", ", excludedTypes)}", "ProcessAllTargetFiles");
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "ProcessAllTargetFiles - Batch processing failed");
        }
    }

    private static void ManageConfiguration()
    {
        while (true)
        {
            Console.WriteLine("\n=== Configuration Management ===");
            Console.WriteLine("1. View current configuration");
            Console.WriteLine("2. Set database path");
            Console.WriteLine("3. Create example config file");
            Console.WriteLine("4. Reload configuration from file");
            Console.WriteLine("5. Save current configuration");
            Console.WriteLine("6. Back to main menu");
            Console.Write("Enter your choice: ");

            string choice = Console.ReadLine()?.Trim() ?? "";

            switch (choice)
            {
                case "1":
                    ViewCurrentConfiguration();
                    break;
                case "2":
                    SetDatabasePath();
                    break;
                case "3":
                    CreateExampleConfigFile();
                    break;
                case "4":
                    ReloadConfiguration();
                    break;
                case "5":
                    SaveCurrentConfiguration();
                    break;
                case "6":
                    ///ShowMainMenu();Should always show on return
                    return;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
                    break;
            }
        }
    }

    private static void ViewCurrentConfiguration()
    {
        if (_config == null)
        {
            Console.WriteLine("No configuration loaded.");
            return;
        }

        Console.WriteLine("\n=== Current Configuration ===");
        Console.WriteLine($"Database Path: {_config.DatabasePath ?? "Not set"}");
        Console.WriteLine($"Auto Create Database: {_config.AutoCreateDatabase}");
        Console.WriteLine($"Default Target Files Path: {_config.DefaultTargetFilesPath ?? "Not set"}");
        Console.WriteLine("\nProcessing Settings:");
        Console.WriteLine($"  Process BOM for Assemblies: {_config.Processing.ProcessBomForAssemblies}");
        Console.WriteLine($"  Include Custom Properties: {_config.Processing.IncludeCustomProperties}");
        Console.WriteLine($"  Validate Files Exist: {_config.Processing.ValidateFilesExist}");
        Console.WriteLine($"  Batch Processing Timeout: {_config.Processing.BatchProcessingTimeout}ms");
    }

    private static void SetDatabasePath()
    {
        Console.Write("Enter database file path: ");
        string dbPath = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(dbPath))
        {
            Console.WriteLine("Database path cannot be empty.");
            return;
        }

        // Expand relative paths to full paths
        dbPath = Path.GetFullPath(dbPath);

        if (!dbPath.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Adding .db extension...");
            dbPath += ".db";
        }

        if (_config == null)
            _config = ConfigManager.GetCurrentConfig();

        _config.DatabasePath = dbPath;
        _databasePath = dbPath;

        Console.WriteLine($"Database path set to: {dbPath}");

        if (!File.Exists(dbPath))
        {
            Console.Write("Database file doesn't exist. Create it now? (y/n): ");
            if (Console.ReadLine()?.Trim().ToLower() == "y")
            {
                try
                {
                    using var dbManager = new SWDatabaseManager(dbPath);
                    Console.WriteLine("Database created successfully!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating database: {ex.Message}");
                }
            }
        }
    }

    private static void CreateExampleConfigFile()
    {
        Console.Write("Enter path for example config file (or press Enter for 'example-config.json'): ");
        string configPath = Console.ReadLine()?.Trim() ?? "";
        
        if (string.IsNullOrEmpty(configPath))
            configPath = "example-config.json";

        try
        {
            ConfigManager.CreateExampleConfig(configPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating example config: {ex.Message}");
        }
    }

    private static void ReloadConfiguration()
    {
        Console.Write("Enter config file path (or press Enter for default): ");
        string? configPath = Console.ReadLine()?.Trim();
        
        if (string.IsNullOrEmpty(configPath))
            configPath = null; // Use default

        try
        {
            _config = ConfigManager.LoadConfigFromJSON(configPath);
            if (!string.IsNullOrEmpty(_config.DatabasePath))
            {
                _databasePath = _config.DatabasePath;
            }
            Console.WriteLine("Configuration reloaded successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reloading configuration: {ex.Message}");
        }
    }

    private static void SaveCurrentConfiguration()
    {
        if (_config == null)
        {
            Console.WriteLine("No configuration to save.");
            return;
        }

        Console.Write("Enter config file path (or press Enter for default): ");
        string? configPath = Console.ReadLine()?.Trim();
        
        if (string.IsNullOrEmpty(configPath))
            configPath = null; // Use default

        try
        {
            ConfigManager.SaveConfig(_config, configPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving configuration: {ex.Message}");
        }
    }
}
