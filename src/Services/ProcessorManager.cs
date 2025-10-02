
/// <summary>
/// Centralized container for application processors and their shared state.
/// Handles refreshing instances when configuration or database path changes.
/// </summary>
public class ProcessManager
{
    private AppConfig? _config;
    private string? _databasePath;

    private TargetFileProcessor? _targetFileProcessor;
    private SWDisplayProcessor? _interactiveFileProcessor;
    private ConfigurationSetupProcessor? _configurationProcessor;
    private DatabaseSetupProcessor? _databaseSetupProcessor;

    /// <summary>
    /// Current configuration assigned to the context.
    /// </summary>
    public AppConfig Config => _config ?? throw new InvalidOperationException("ProcessorContext has not been initialized with configuration.");

    /// <summary>
    /// Current database path tracked by the context.
    /// </summary>
    public string DatabasePath => _databasePath ?? throw new InvalidOperationException("DatabasePath is not initialized.");

    /// <summary>
    /// Indicates whether a database path is currently available.
    /// </summary>
    public bool HasDatabase => !string.IsNullOrWhiteSpace(_databasePath);


    /// <summary>
    /// Exposes the configuration processor (always available after initialization).
    /// </summary>
    public ConfigurationSetupProcessor ConfigurationProcessor =>
        _configurationProcessor ?? throw new InvalidOperationException("ConfigurationProcessor is not initialized.");


    /// <summary>
    /// Exposes the target file processor if initialized (always available after initialization).
    /// </summary>
    public TargetFileProcessor TargetFileProcessor => _targetFileProcessor ??
        _targetFileProcessor ?? throw new InvalidOperationException("TargetFileProcessor is not initialized.");

    /// <summary>
    /// Exposes the interactive file processor (always available after initialization).
    /// </summary>
    public SWDisplayProcessor InteractiveFileProcessor =>
        _interactiveFileProcessor ?? throw new InvalidOperationException("InteractiveFileProcessor is not initialized.");

    /// <summary>
    /// Exposes the database setup processor (always available after initialization).
    /// </summary>
    public DatabaseSetupProcessor DatabaseSetupProcessor =>
        _databaseSetupProcessor ?? throw new InvalidOperationException("DatabaseSetupProcessor is not initialized.");

    /// <summary>
    /// Initialize the context with the supplied configuration and refresh all processors.
    /// </summary>
    public void Initialize(AppConfig config)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _databasePath = _config.DatabasePath;
        RefreshInteractiveProcessor();
        RefreshConfigurationProcessor();
        RefreshDatabaseSetupProcessor();
        RefreshTargetProcessor();
    }

    /// <summary>
    /// Handle configuration reloads by replacing the current configuration and refreshing all processors.
    /// </summary>
    /// <param name="newConfig">The newly loaded configuration instance.</param>
    public void HandleConfigurationReload(AppConfig newConfig)
    {
        Initialize(newConfig);
    }

    /// <summary>
    /// Handle database path changes raised by processors and refresh dependent services.
    /// </summary>
    /// <param name="newDatabasePath">The new database path value.</param>
    public void HandleDatabasePathChanged(string? newDatabasePath)
    {
        _databasePath = string.IsNullOrWhiteSpace(newDatabasePath) ? null : newDatabasePath;

        if (_config == null)
        {
            throw new InvalidOperationException("Configuration must be initialized before updating the database path.");
        }

        _config.DatabasePath = _databasePath;

        RefreshTargetProcessor();
        RefreshInteractiveProcessor();
    }

    /// <summary>
    /// Convenience helper for executing an action when the target file processor is available.
    /// </summary>
    /// <param name="action">Action to perform when the processor exists.</param>
    /// <returns>True if the processor existed and the action ran, otherwise false.</returns>
    public bool TryWithTargetProcessor(Action<TargetFileProcessor> action)
    {
        if (_targetFileProcessor == null)
        {
            return false;
        }

        action(_targetFileProcessor);
        return true;
    }

    private void RefreshTargetProcessor()
    {
        if (!HasDatabase)
        {
            _targetFileProcessor = null;
            return;
        }

        _targetFileProcessor = new TargetFileProcessor(Config);
    }

    private void RefreshInteractiveProcessor()
    {
        _interactiveFileProcessor = new SWDisplayProcessor(Config);
    }

    private void RefreshConfigurationProcessor()
    {
        _configurationProcessor = new ConfigurationSetupProcessor(Config, HandleDatabasePathChanged);
    }

    private void RefreshDatabaseSetupProcessor()
    {
        _databaseSetupProcessor = new DatabaseSetupProcessor(Config, path => HandleDatabasePathChanged(path));
    }
}
