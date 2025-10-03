
using System.Diagnostics;

/// <summary>
/// Utility class for running Python scripts from the C# application
/// </summary>
public static class PythonScriptRunner
{
    private static readonly Dictionary<string, PythonScript> _registeredScripts = new();
    private static readonly object _lock = new();

    /// <summary>
    /// Registers a Python script for execution
    /// </summary>
    public static void RegisterScript(string filePath, string displayName, string description)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));
        if (string.IsNullOrEmpty(displayName))
            throw new ArgumentNullException(nameof(displayName));
        if (string.IsNullOrEmpty(description))
            throw new ArgumentNullException(nameof(description));

        lock (_lock)
        {
            var script = new PythonScript(filePath, displayName, description);
            _registeredScripts[script.FileName] = script;
            Logger.LogRuntime($"Registered Python script: {script.FileName} -> {filePath}", "PythonScriptRunner");
        }
    }

    /// <summary>
    /// Registers multiple scripts at once
    /// </summary>
    public static void RegisterScripts(IEnumerable<(string filePath, string displayName, string description)> scripts)
    {
        foreach (var (filePath, displayName, description) in scripts)
        {
            RegisterScript(filePath, displayName, description);
        }
    }

    /// <summary>
    /// Gets all registered scripts
    /// </summary>
    public static IReadOnlyDictionary<string, PythonScript> GetRegisteredScripts()
    {
        lock (_lock)
        {
            return new Dictionary<string, PythonScript>(_registeredScripts);
        }
    }

    /// <summary>
    /// Runs a Python script by file path
    /// </summary>
    public static void RunPythonScript(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            Logger.WriteAndLogUserMessage("Error: No file path provided");
            return;
        }

        try
        {
            Logger.LogRuntime($"Starting Python script: {filePath}", "PythonScriptRunner");

            if (!File.Exists(filePath))
            {
                Logger.WriteAndLogUserMessage($"Error: Python script not found: {filePath}");
                return;
            }

            // Check if Python is available
            if (!IsPythonAvailable())
            {
                Logger.WriteAndLogUserMessage("Error: Python is not installed or not in PATH");
                return;
            }

            string workingDirectory = Path.GetDirectoryName(filePath) ?? Directory.GetCurrentDirectory();

            var startInfo = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{filePath}\"",
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = false
            };

            Logger.WriteAndLogUserMessage($"Executing: python \"{filePath}\"");
            Logger.WriteAndLogUserMessage("Please wait...");

            using (var process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    Logger.WriteAndLogUserMessage("Error: Failed to start Python process");
                    return;
                }

                // Read output asynchronously to avoid blocking
                process.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Logger.WriteAndLogUserMessage($"[Python] {e.Data}");
                    }
                };

                process.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Logger.WriteAndLogUserMessage($"[Python Error] {e.Data}");
                    }
                };

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                // Wait for the process to complete
                if (!process.WaitForExit(300000)) // 5 minute timeout
                {
                    Logger.WriteAndLogUserMessage("Warning: Python script execution timed out");
                    process.Kill();
                }
                else
                {
                    Logger.LogRuntime($"Python script completed with exit code: {process.ExitCode}", "PythonScriptRunner");
                    if (process.ExitCode == 0)
                    {
                        Logger.WriteAndLogUserMessage("Python script completed successfully");
                    }
                    else
                    {
                        Logger.WriteAndLogUserMessage($"Python script completed with errors (exit code: {process.ExitCode})");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "PythonScriptRunner.RunPythonScript");
            Logger.WriteAndLogUserMessage($"Error running Python script: {ex.Message}");
        }
    }

    /// <summary>
    /// Runs a registered Python script by name
    /// </summary>
    public static void RunRegisteredScript(string scriptName)
    {
        PythonScript? script;
        lock (_lock)
        {
            _registeredScripts.TryGetValue(scriptName, out script);
        }

        if (script == null)
        {
            Logger.WriteAndLogUserMessage($"Error: Registered script '{scriptName}' not found");
            return;
        }

        RunPythonScript(script.FilePath);
    }

    /// <summary>
    /// Initializes default scripts from configuration
    /// </summary>
    public static void InitializeDefaultScripts()
    {
        var config = ConfigManager.GetCurrentConfig();

        // Auto-register from configured directory

        // Register scripts from config
        if (config?.PythonScripts != null)
        {
            foreach (var scriptConfig in config.PythonScripts)
            {
                if (!string.IsNullOrEmpty(scriptConfig.FilePath) &&
                    !string.IsNullOrEmpty(scriptConfig.DisplayName) &&
                    !string.IsNullOrEmpty(scriptConfig.Description))
                {
                    try
                    {
                        RegisterScript(scriptConfig.FilePath, scriptConfig.DisplayName, scriptConfig.Description);
                        Logger.LogRuntime($"Auto-registered config script: {scriptConfig.DisplayName}", "PythonScriptRunner");
                    }
                    catch (Exception ex)
                    {
                        Logger.LogRuntime($"Failed to register config script {scriptConfig.FilePath}: {ex.Message}", "PythonScriptRunner");
                    }
                }
            }
        }

        // You can also manually register specific scripts here
        // Examples:
        // RegisterScript(@"c:\path\to\custom_script.py", "custom", "Run custom data analysis");
        // RegisterScript(@"d:\scripts\backup.py", "backup", "Execute backup procedure");
    }

    /// <summary>
    /// Auto-registers scripts from a directory
    /// </summary>
    public static void AutoRegisterScriptsFromDirectory(string directoryPath, string descriptionPrefix = "")
    {
        try
        {
            if (!Directory.Exists(directoryPath))
            {
                Logger.LogRuntime($"Directory not found for auto-registration: {directoryPath}", "PythonScriptRunner");
                return;
            }

            var scripts = Directory.GetFiles(directoryPath, "*.py")
                                .Where(file => !Path.GetFileName(file).StartsWith("_"))
                                .Select(file => (
                                    filePath: file,
                                    displayName: Path.GetFileNameWithoutExtension(file).Replace("_", " "),
                                    description: $"{descriptionPrefix}{Path.GetFileNameWithoutExtension(file).Replace("_", " ")}"
                                  ));

            RegisterScripts(scripts);
            Logger.LogRuntime($"Auto-registered {scripts.Count()} scripts from {directoryPath}", "PythonScriptRunner");
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "PythonScriptRunner.AutoRegisterScriptsFromDirectory");
        }
    }

    /// <summary>
    /// Creates and runs a menu for managing Python scripts
    /// </summary>
    public static void ShowPythonScriptsMenu()
    {
        // Auto-register scripts from configured directory
        var config = ConfigManager.GetCurrentConfig();

        // Get registered scripts
        var registeredScripts = GetRegisteredScripts();

        var menu = MenuFactoryExtensions.CreateStandardMenu("Python Scripts");

        // Add registered scripts to menu
        int optionNumber = 1;
        foreach (var script in registeredScripts.Values.OrderBy(s => s.DisplayName))
        {
            menu.AddOption(optionNumber.ToString(), script.Description,
                () => RunRegisteredScript(script.DisplayName));
            optionNumber++;
        }

        // Add management options
        menu.AddOption("A", "Add custom script", AddCustomScript);
        menu.AddOption("L", "List all registered scripts", ListRegisteredScripts);
        menu.AddBackOption("0", "Back to main menu");

        menu.RunMenu();
    }

    /// <summary>
    /// Checks if Python is available on the system
    /// </summary>
    private static bool IsPythonAvailable()
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var process = Process.Start(startInfo))
            {
                if (process != null)
                {
                    process.WaitForExit();
                    return process.ExitCode == 0;
                }
            }
        }
        catch
        {
            // Python not found
        }
        return false;
    }

    /// <summary>
    /// Allows user to add a custom script through interactive prompts
    /// </summary>
    private static void AddCustomScript()
    {
        Logger.WriteAndLogUserMessage("Add Custom Python Script");
        Logger.WriteAndLogUserMessage("========================");

        Console.Write("Enter the full path to the Python script: ");
        string? filePath = Console.ReadLine()?.Trim();

        if (string.IsNullOrEmpty(filePath))
        {
            Logger.WriteAndLogUserMessage("No path entered. Operation cancelled.");
            return;
        }

        // Validate the file exists
        if (!File.Exists(filePath))
        {
            Logger.WriteAndLogUserMessage($"Error: File does not exist: {filePath}");
            return;
        }

        // Validate it's a Python file
        if (!filePath.EndsWith(".py", StringComparison.OrdinalIgnoreCase))
        {
            Logger.WriteAndLogUserMessage($"Error: File must have .py extension: {filePath}");
            return;
        }

        Console.Write("Enter a description for this script: ");
        string? description = Console.ReadLine()?.Trim();

        if (string.IsNullOrEmpty(description))
        {
            description = Path.GetFileNameWithoutExtension(filePath).Replace("_", " ");
        }

        // Generate display name from filename
        string displayName = Path.GetFileNameWithoutExtension(filePath);

        try
        {
            RegisterScript(filePath, displayName, description);
            Logger.WriteAndLogUserMessage($"Successfully registered script: {description}");
            Logger.WriteAndLogUserMessage($"Path: {filePath}");
        }
        catch (Exception ex)
        {
            Logger.WriteAndLogUserMessage($"Error registering script: {ex.Message}");
        }

        Console.WriteLine("Press any key to continue...");
        Console.ReadKey();
    }

    /// <summary>
    /// Lists all currently registered scripts
    /// </summary>
    private static void ListRegisteredScripts()
    {
        var registeredScripts = GetRegisteredScripts();

        Logger.WriteAndLogUserMessage("Registered Python Scripts");
        Logger.WriteAndLogUserMessage("========================");

        if (registeredScripts.Count == 0)
        {
            Logger.WriteAndLogUserMessage("No scripts currently registered.");
        }
        else
        {
            int count = 1;
            foreach (var script in registeredScripts.Values.OrderBy(s => s.DisplayName))
            {
                Logger.WriteAndLogUserMessage($"{count}. {script.DisplayName}");
                Logger.WriteAndLogUserMessage($"   Description: {script.Description}");
                Logger.WriteAndLogUserMessage($"   Path: {script.FilePath}");
                Logger.WriteAndLogUserMessage("");
                count++;
            }
        }

        Console.WriteLine("Press any key to continue...");
        Console.ReadKey();
    }
}