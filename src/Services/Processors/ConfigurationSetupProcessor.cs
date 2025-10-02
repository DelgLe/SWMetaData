using System;
using System.IO;

public class ConfigurationSetupProcessor(AppConfig config, Action<string?> onDatabasePathChanged)
{
    private AppConfig _config = config ?? throw new ArgumentNullException(nameof(config));
    private Action<string?> _onDatabasePathChanged = onDatabasePathChanged ?? throw new ArgumentNullException(nameof(onDatabasePathChanged));

    public void ViewCurrentConfiguration()
    {
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

    public void SetDatabasePath()
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

        _config.DatabasePath = dbPath;
        _onDatabasePathChanged(dbPath);

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

    public void CreateExampleConfigFile()
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

    public void ReloadConfiguration(Action<AppConfig> onConfigReloaded)
    {
        Console.Write("Enter config file path (or press Enter for default): ");
        string? configPath = Console.ReadLine()?.Trim();
        
        if (string.IsNullOrEmpty(configPath))
            configPath = null; // Use default

        try
        {
            var newConfig = ConfigManager.LoadConfigFromJSON(configPath);
            _config = newConfig;
            onConfigReloaded(newConfig);
            Console.WriteLine("Configuration reloaded successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reloading configuration: {ex.Message}");
        }
    }

    public void SaveCurrentConfiguration()
    {
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