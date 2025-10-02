
/// <summary>
/// Configuration settings for the SolidWorks Metadata application
/// </summary>
public class AppConfig
{
    // Database Settings
    public string? DatabasePath { get; set; }
    public bool AutoCreateDatabase { get; set; } = true;
    public string? DefaultTargetFilesPath { get; set; }
    
    // Logging Settings
    public string? LogDirectory { get; set; }
    
    // Processing Settings
    public bool ProcessBomForAssemblies { get; set; } = true;
    public bool IncludeCustomProperties { get; set; } = true;
    public bool ValidateFilesExist { get; set; } = true;
    public int BatchProcessingTimeout { get; set; } = 30000; // milliseconds
    public int PeriodicCleanupInterval { get; set; } = 10; // files between cleanup
    
    // SolidWorks Settings
    public bool HideUserInterface { get; set; } = true; // swApp.Visible = false
    public bool DisableUserControl { get; set; } = true; // swApp.UserControl = false
    public bool UseReadOnlyMode { get; set; } = true; // Opens documents in read-only mode
    public bool LoadHiddenComponents { get; set; } = false; // Skip hidden components for speed
    public bool LoadExternalReferencesInMemory { get; set; } = true; // Faster reference loading
    
    // Cleanup Settings
    public int CleanupDelayMs { get; set; } = 500; // milliseconds to wait during cleanup
    public int DocumentCloseDelayMs { get; set; } = 200; // milliseconds to wait after document close
}