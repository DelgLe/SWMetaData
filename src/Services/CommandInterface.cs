using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class CommandInterface
{
    private static readonly ProcessManager _processors = new();
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
        Logger.LogRuntime("Database path loaded from config", _processors.DatabasePath);

        // Initialize Python scripts
        PythonScriptRunner.InitializeDefaultScripts();

        var mainMenu = new MenuFactory("SolidWorks Metadata Reader")
            .AddOption("1", "Process single file (display only)", () => InteractiveFiles.ProcessSingleFile(swApp))
            .AddOption("2", "Setup/Create database", DatabaseSetup.SetupDatabase)
            .AddOption("3", "Process file and save to database", () => InteractiveFiles.ProcessFileWithDatabase(swApp))
            .AddOption("4", "Manage target files", ManageTargetFiles)
            .AddOption("5", "Process all target files (batch)", () => TargetFiles.ProcessAllTargetFiles(swApp))
            .AddOption("6", "Configuration settings", ManageConfiguration)
            .AddOption("7", "Run Python scripts", PythonScriptRunner.ShowPythonScriptsMenu)
            .AddOption("8", "Exit", () =>
            {
                Logger.LogRuntime("Application exit requested", "User input");
                return false;
            });

        mainMenu.RunMenu();
    }

    private static void ManageTargetFiles()
    {
        var menu = MenuFactoryExtensions.CreateStandardMenu("Target Files Management")
            .AddOption("1", "View all target files", TargetFiles.ViewTargetFilesInteractive)
            .AddOption("2", "Add target file", TargetFiles.AddTargetFileInteractive)
            .AddOption("3", "Remove target file", TargetFiles.RemoveTargetFileInteractive)
            .AddOption("4", "Validate target files (check if files exist)", TargetFiles.ValidateTargetFilesInteractive)
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
