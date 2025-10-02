using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class CommandInterface
{
    private static readonly ProcessorContext _processors = new();
    private static SWDisplayProcessor InteractiveFiles => _processors.InteractiveFileProcessor;
    private static DatabaseSetupProcessor DatabaseSetup => _processors.DatabaseSetupProcessor;
    private static ConfigurationSetupProcessor ConfigurationSetup => _processors.ConfigurationProcessor;
    private static TargetFileProcessor TargetFiles => _processors.TargetFileProcessor;

    /// <summary>
    /// Entry point for the interactive command loop.
    /// </summary>
    public static void RunInteractiveLoop(SldWorks swApp)
    {
        var config = ConfigManager.GetCurrentConfig();
        _processors.Initialize(config);

        Logger.LogRuntime("Application started", "RunInteractiveLoop");
        if (_processors.HasDatabase)
        {
            Logger.LogRuntime("Database path loaded from config", _processors.DatabasePath);
        }

        var mainMenu = new MenuFactory("SolidWorks Metadata Reader")
            .AddOption("1", "Process single file (display only)", () => InteractiveFiles.ProcessSingleFile(swApp))
            .AddOption("2", "Setup/Create database", DatabaseSetup.SetupDatabase)
            .AddOption("3", "Process file and save to database", () => InteractiveFiles.ProcessFileWithDatabase(swApp))
            .AddOption("4", "Manage target files", ManageTargetFiles)
            .AddOption("5", "Process all target files (batch)", () => TargetFiles.ProcessAllTargetFiles(swApp))
            .AddOption("6", "Configuration settings", ManageConfiguration)
            .AddOption("7", "Exit", () =>
            {
                Logger.LogRuntime("Application exit requested", "User input");
                return false;
            });

        mainMenu.RunMenu();
    }


    public static void OfferBomOption(SldWorks swApp, string filePath)
    {
        try
        {
            var menu = MenuFactoryExtensions.CreateStandardMenu("Assembly BOM Options")
                .AddOption("1", "View unsuppressed components only (BOM)", () =>
                {
                    DisplayBom(swApp, filePath, false);
                    return false;
                })
                .AddOption("2", "View all components (including suppressed)", () =>
                {
                    DisplayBom(swApp, filePath, true);
                    return false;
                })
                .AddOption("3", "Skip BOM", () => false);

            Console.WriteLine("Would you like to view the Bill of Materials (BOM) for this assembly?");
            menu.RunMenu();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error offering BOM option: {ex.Message}");
        }
    }

    public static void DisplayBom(SldWorks swApp, string filePath, bool includeSupressed)
    {
        try
        {
            var bomItems = SWMetadataReader.GetBomItems(swApp, filePath, includeSupressed, message =>
            {
                Logger.WriteAndLogUserMessage(message.TrimStart());
            });

            string title = includeSupressed
                ? "Complete Component List (Including Suppressed)"
                : "Bill of Materials (Unsuppressed Components Only)";

            DisplayBomFromList(bomItems, title);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating BOM: {ex.Message}");
        }
    }

    private static void DisplayBomFromList(List<BomItem> bomItems, string title)
    {
        Console.WriteLine($"\n--- {title} ---");

        if (bomItems.Count == 0)
        {
            Console.WriteLine("No components found in the assembly.");
            return;
        }

        Console.WriteLine($"Total components: {bomItems.Count}");
        Console.WriteLine(new string('-', 80));

        foreach (var item in bomItems)
        {
            string indent = new string(' ', item.Level * 2);
            string quantity = item.Quantity > 1 ? $"({item.Quantity}x) " : string.Empty;
            string suppression = item.IsSuppressed ? " [SUPPRESSED]" : string.Empty;

            Console.WriteLine($"{indent}{quantity}{item.ComponentName}");
            Console.WriteLine($"{indent}  File: {item.FileName} | Config: {item.Configuration}{suppression}");

            if (item.Level == 0)
            {
                Console.WriteLine();
            }
        }
    }

    private static void ManageTargetFiles()
    {
        if (!_processors.HasDatabase || TargetFiles is not TargetFileProcessor targetProcessor)
        {
            Console.WriteLine("No database configured. Please setup database first (option 2).");
            return;
        }

        var menu = MenuFactoryExtensions.CreateStandardMenu("Target Files Management")
            .AddOption("1", "View all target files", targetProcessor.ViewTargetFilesInteractive)
            .AddOption("2", "Add target file", targetProcessor.AddTargetFileInteractive)
            .AddOption("3", "Remove target file", targetProcessor.RemoveTargetFileInteractive)
            .AddOption("4", "Validate target files (check if files exist)", targetProcessor.ValidateTargetFilesInteractive)
            .AddBackOption("5", "Back to main menu");

        menu.RunMenu();
    }

    private static void ManageConfiguration()
    {
        var menu = MenuFactoryExtensions.CreateStandardMenu("Configuration Management")
            .AddOption("1", "View current configuration", ConfigurationSetup.ViewCurrentConfiguration)
            .AddOption("2", "Set database path", ConfigurationSetup.SetDatabasePath)
            .AddOption("3", "Create example config file", ConfigurationSetup.CreateExampleConfigFile)
            .AddOption("4", "Reload configuration from file", () =>
                ConfigurationSetup.ReloadConfiguration(_processors.HandleConfigurationReload))
            .AddOption("5", "Save current configuration", ConfigurationSetup.SaveCurrentConfiguration)
            .AddBackOption("6", "Back to main menu");

        menu.RunMenu();
    }
}
