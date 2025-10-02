using System;
using System.IO;
using System.Text.Json;

/// <summary>
/// Configuration settings for the SolidWorks Metadata application
/// </summary>
public class AppConfig
{
    public string? DatabasePath { get; set; }
    public bool AutoCreateDatabase { get; set; } = true;
    public string? DefaultTargetFilesPath { get; set; }
    public string? LogDirectory { get; set; }
    public ProcessingSettings Processing { get; set; } = new ProcessingSettings();
}

/// <summary>
/// Processing-related configuration settings
/// </summary>
public class ProcessingSettings
{
    public bool ProcessBomForAssemblies { get; set; } = true;
    public bool IncludeCustomProperties { get; set; } = true;
    public bool ValidateFilesExist { get; set; } = true;
    public int BatchProcessingTimeout { get; set; } = 30000; // milliseconds
}

/// <summary>
/// Configuration manager for loading and saving application settings
/// </summary>
public static class ConfigManager
{
    private const string DefaultConfigFilePath = "C:\\Users\\alexanderd\\Source\\Repos\\DelgLe\\SWMetaData\\swmetadata-config.json";
    private static AppConfig? _currentConfig;

    /// <summary>
    /// Load configuration from JSON file
    /// </summary>
    /// <param name="configPath">Path to config file (optional, uses default if not provided)</param>
    /// <returns>Loaded configuration</returns>
    public static AppConfig LoadConfig(string? configPath = null)
    {
        configPath ??= DefaultConfigFilePath;

        if (!File.Exists(configPath))
        {
            Console.WriteLine($"Config file not found: {configPath}");
            Console.WriteLine("Creating default configuration...");
            var defaultConfig = CreateDefaultConfig();
            SaveConfig(defaultConfig, configPath);
            return defaultConfig;
        }

        try
        {
            string jsonContent = File.ReadAllText(configPath);
            var config = JsonSerializer.Deserialize<AppConfig>(jsonContent, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true
            });

            _currentConfig = config ?? CreateDefaultConfig();
            Console.WriteLine($"Configuration loaded from: {configPath}");
            
            // Validate database path if specified
            if (!string.IsNullOrEmpty(_currentConfig.DatabasePath))
            {
                string fullPath = Path.GetFullPath(_currentConfig.DatabasePath);
                if (!File.Exists(fullPath) && _currentConfig.AutoCreateDatabase)
                {
                    Console.WriteLine($"Database file will be created at: {fullPath}");
                }
                else if (!File.Exists(fullPath))
                {
                    Console.WriteLine($"Warning: Database file not found: {fullPath}");
                }
            }

            return _currentConfig;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading config: {ex.Message}");
            Console.WriteLine("Using default configuration...");
            return CreateDefaultConfig();
        }
    }

    /// <summary>
    /// Save configuration to JSON file
    /// </summary>
    /// <param name="config">Configuration to save</param>
    /// <param name="configPath">Path to save config file (optional, uses default if not provided)</param>
    public static void SaveConfig(AppConfig config, string? configPath = null)
    {
        configPath ??= DefaultConfigFilePath;

        try
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            string jsonContent = JsonSerializer.Serialize(config, options);
            File.WriteAllText(configPath, jsonContent);
            
            _currentConfig = config;
            Console.WriteLine($"Configuration saved to: {configPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving config: {ex.Message}");
        }
    }

    /// <summary>
    /// Get current loaded configuration
    /// </summary>
    public static AppConfig GetCurrentConfig()
    {
        return _currentConfig ?? LoadConfig();
    }

    /// <summary>
    /// Create default configuration
    /// </summary>
    private static AppConfig CreateDefaultConfig()
    {
        return new AppConfig
        {
            DatabasePath = Path.Combine(Environment.CurrentDirectory, "swmetadata.db"),
            AutoCreateDatabase = true,
            DefaultTargetFilesPath = null,
            LogDirectory = Path.Combine(Environment.CurrentDirectory, "logs"),
            Processing = new ProcessingSettings
            {
                ProcessBomForAssemblies = true,
                IncludeCustomProperties = true,
                ValidateFilesExist = true,
                BatchProcessingTimeout = 30000
            }
        };
    }

    /// <summary>
    /// Create example configuration file
    /// </summary>
    /// <param name="filePath">Path where to create the example</param>
    public static void CreateExampleConfig(string filePath)
    {
        var exampleConfig = new AppConfig
        {
            DatabasePath = @"C:\SolidWorksData\metadata.db",
            AutoCreateDatabase = true,
            DefaultTargetFilesPath = @"C:\SolidWorksData\target_files.csv",
            Processing = new ProcessingSettings
            {
                ProcessBomForAssemblies = true,
                IncludeCustomProperties = true,
                ValidateFilesExist = true,
                BatchProcessingTimeout = 60000
            }
        };

        SaveConfig(exampleConfig, filePath);
        Console.WriteLine($"Example configuration created at: {filePath}");
    }
}