using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class CommandInterface
{
    private static string? _databasePath = null;
    private static AppConfig? _config = null;
    private static TargetFileProcessor? _targetFileProcessor = null;
    private static InteractiveFileProcessor? _interactiveFileProcessor = null;

    /// <summary>
    /// Properly close a SolidWorks document with multiple fallback methods
    /// </summary>

    public static void RunInteractiveLoop(SldWorks swApp)
    {
        _config = ConfigManager.GetCurrentConfig();
        RefreshInteractiveFileProcessor();
        Logger.LogRuntime("Application started", "RunInteractiveLoop");
        if (!string.IsNullOrEmpty(_config.DatabasePath))
        {
            _databasePath = _config.DatabasePath;
            Logger.LogRuntime("Database path loaded from config", _databasePath);
            RefreshTargetFileProcessor();
            RefreshInteractiveFileProcessor();
        }

        var mainMenu = new MenuFactory("SolidWorks Metadata Reader")
            .AddOption("1", "Process single file (display only)", () => RequireInteractiveFileProcessor().ProcessSingleFile(swApp))
            .AddOption("2", "Setup/Create database", SetupDatabase)
            .AddOption("3", "Process file and save to database", () => ProcessFileWithDatabase(swApp))
            .AddOption("4", "Manage target files", ManageTargetFiles)
            .AddOption("5", "Process all target files (batch)", () => ProcessAllTargetFiles(swApp))
            .AddOption("6", "Configuration settings", ManageConfiguration)
            .AddOption("7", "Exit", () => {
                Logger.LogRuntime("Application exit requested", "User input");
                return false; // Exit the menu loop
            });

        // Add status information to the menu display
        AddMainMenuStatusInfo();
        
        mainMenu.RunMenu();
    }

    private static void ProcessFileWithDatabase(SldWorks swApp)
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Logger.LogWarning("No database configured. Please setup database first (option 2).");
            return;
        }

        RequireInteractiveFileProcessor().ProcessFileWithDatabase(swApp, _databasePath!);
    }

    private static void AddMainMenuStatusInfo()
    {
        Console.WriteLine($"Current database: {(_databasePath ?? "Not set")}");
        if (_config != null)
        {
            Console.WriteLine($"Config loaded: {(_config.DatabasePath != null ? "Yes" : "Default")}");
        }
    }


    private static void SetupDatabase()
    {
        // Check if config has a database path
        bool hasConfigDatabase = _config != null && !string.IsNullOrEmpty(_config.DatabasePath);
        
        if (hasConfigDatabase)
        {
            Console.WriteLine($"Configuration database path found: {_config!.DatabasePath}");
            
            var menu = MenuFactoryExtensions.CreateStandardMenu("Database Setup")
                .AddOption("1", "Use existing config database path (setup tables if needed)", () => {
                    SetupConfigDatabase();
                    return false; // Exit menu after selection
                })
                .AddOption("2", "Create new database in current folder", () => {
                    SetupNewDatabase();
                    return false; // Exit menu after selection
                });
                
            menu.RunMenu();
        }
        else
        {
            // No config database, go straight to creating new database
            SetupNewDatabase();
        }
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
            using (var dbManager = new DatabaseGateway(configDbPath))
            {
                _databasePath = configDbPath;
                RefreshTargetFileProcessor();
                RefreshInteractiveFileProcessor();
                
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
            using (var dbManager = new DatabaseGateway(newDbPath))
            {
                _databasePath = newDbPath;
                RefreshTargetFileProcessor();
                RefreshInteractiveFileProcessor();
                
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
            
            swDocumentTypes_e docType = SWDocumentManager.GetDocumentType(filePath);
            Logger.WriteAndLogUserMessage($"Opening {docType} for BOM display: {Path.GetFileName(filePath)}");
            
            swModel = swApp.OpenDoc6(filePath, (int)docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            if (swModel == null)
            {
                Console.WriteLine($"Error: Could not open assembly for BOM generation (Errors: {errors}, Warnings: {warnings}).");
                
                // Assembly files sometimes hang even when open fails - force cleanup
                if (docType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    SWDocumentManager.ForceCleanupHangingDocuments(swApp, "Assembly failed to open - ");
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
            SWDocumentManager.CloseDocumentSafely(swApp, swModel);
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

    private static void RefreshTargetFileProcessor()
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            _targetFileProcessor = null;
            return;
        }

        _targetFileProcessor = new TargetFileProcessor(_databasePath, _config);
    }

    private static TargetFileProcessor? RequireTargetFileProcessor()
    {
        RefreshTargetFileProcessor();

        if (_targetFileProcessor == null)
        {
            Console.WriteLine("Target file processor is unavailable. Please configure the database first.");
        }

        return _targetFileProcessor;
    }

    private static void RefreshInteractiveFileProcessor()
    {
        _interactiveFileProcessor = new InteractiveFileProcessor(_config);
    }

    private static InteractiveFileProcessor RequireInteractiveFileProcessor()
    {
        if (_interactiveFileProcessor == null)
        {
            RefreshInteractiveFileProcessor();
        }

        return _interactiveFileProcessor!;
    }

    private static void ManageTargetFiles()
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Console.WriteLine("No database configured. Please setup database first (option 2).");
            return;
        }

        if (RequireTargetFileProcessor() == null)
        {
            return;
        }

        var menu = MenuFactoryExtensions.CreateStandardMenu("Target Files Management")
            .AddOption("1", "View all target files", ViewTargetFiles)
            .AddOption("2", "Add target file", AddTargetFile)
            .AddOption("3", "Remove target file", RemoveTargetFile)
            .AddOption("4", "Validate target files (check if files exist)", ValidateTargetFiles)
            .AddBackOption("5", "Back to main menu");

        menu.RunMenu();
    }

    private static void ViewTargetFiles()
    {
        var processor = RequireTargetFileProcessor();
        processor?.ViewTargetFilesInteractive();
    }

    private static void AddTargetFile()
    {
        var processor = RequireTargetFileProcessor();
        processor?.AddTargetFileInteractive();
    }

    private static void RemoveTargetFile()
    {
        var processor = RequireTargetFileProcessor();
        processor?.RemoveTargetFileInteractive();
    }

    private static void ValidateTargetFiles()
    {
        var processor = RequireTargetFileProcessor();
        processor?.ValidateTargetFilesInteractive();
    }

    private static void ProcessAllTargetFiles(SldWorks swApp)
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Logger.WriteAndLogUserMessage("No database configured. Please setup database first (option 2).");
            return;
        }

        var processor = RequireTargetFileProcessor();
        if (processor == null) return;

        try
        {
            Logger.LogRuntime("Batch processing started", "ProcessAllTargetFiles");
            var validFiles = processor.GetValidTargetFiles();

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

            var result = processor.ProcessFiles(swApp, validFiles, message => Console.WriteLine(message));

            Logger.WriteAndLogUserMessage($"\n=== Batch Processing Complete ===");
            Logger.WriteAndLogUserMessage($"Successfully processed: {result.ProcessedFiles}");
            Logger.WriteAndLogUserMessage($"Errors (including exclusions): {result.ErrorCount}");
            if (result.ExcludedCount > 0)
            {
                Logger.WriteAndLogUserMessage($"Excluded files: {result.ExcludedCount}");
            }
            Logger.WriteAndLogUserMessage($"Total: {result.TotalFiles}");

            Logger.LogInfo(
                $"Batch processing complete - Processed: {result.ProcessedFiles}, Errors: {result.ErrorCount}, Excluded: {result.ExcludedCount}, Total: {result.TotalFiles}",
                "ProcessAllTargetFiles");
            
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
        var menu = MenuFactoryExtensions.CreateStandardMenu("Configuration Management")
            .AddOption("1", "View current configuration", ViewCurrentConfiguration)
            .AddOption("2", "Set database path", SetDatabasePath)
            .AddOption("3", "Create example config file", CreateExampleConfigFile)
            .AddOption("4", "Reload configuration from file", ReloadConfiguration)
            .AddOption("5", "Save current configuration", SaveCurrentConfiguration)
            .AddBackOption("6", "Back to main menu");

        menu.RunMenu();
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
        Console.WriteLine($"  Process BOM for Assemblies: {_config.ProcessBomForAssemblies}");
        Console.WriteLine($"  Include Custom Properties: {_config.IncludeCustomProperties}");
        Console.WriteLine($"  Validate Files Exist: {_config.ValidateFilesExist}");
        Console.WriteLine($"  Batch Processing Timeout: {_config.BatchProcessingTimeout}ms");
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
        RefreshTargetFileProcessor();
        RefreshInteractiveFileProcessor();

        Console.WriteLine($"Database path set to: {dbPath}");

        if (!File.Exists(dbPath))
        {
            Console.Write("Database file doesn't exist. Create it now? (y/n): ");
            if (Console.ReadLine()?.Trim().ToLower() == "y")
            {
                try
                {
                    using var dbManager = new DatabaseGateway(dbPath);
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
            _databasePath = _config.DatabasePath;
            RefreshTargetFileProcessor();
            RefreshInteractiveFileProcessor();
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
