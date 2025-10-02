using SolidWorks.Interop.sldworks;

public class SWDisplayProcessor
{
    private readonly AppConfig _config;

    public SWDisplayProcessor(AppConfig config)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
    }

    private string DatabasePath => !string.IsNullOrWhiteSpace(_config.DatabasePath)
        ? _config.DatabasePath!
        : throw new InvalidOperationException("Database path is not configured.");


    public void ProcessSingleFile(SldWorks swApp)
    {
        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
            DisplayMetadata(filePath, metadata);

            if (!_config.ProcessBomForAssemblies)
            {
                OfferBomOption(swApp, filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Details: {ex.InnerException.Message}");
            }
        }
    }

    public void ProcessFileWithDatabase(SldWorks swApp)
    {
        if (string.IsNullOrWhiteSpace(DatabasePath))
        {
            Logger.LogWarning("No database configured. Please setup database first (option 2).");
            return;
        }

        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            Logger.LogInfo("Reading metadata from SolidWorks file...");
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);

            DisplayMetadata(filePath, metadata);

            using var dbManager = new DatabaseGateway(DatabasePath);
            Console.WriteLine("Saving to database...");
            long documentId = dbManager.InsertDocumentMetadata(metadata);
            Console.WriteLine($"Document saved with ID: {documentId}");

            if (!_config.ProcessBomForAssemblies)
            {
                ProcessAssemblyBom(swApp, dbManager, filePath);
            }

            Console.WriteLine("File processing complete!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Details: {ex.InnerException.Message}");
            }
        }
    }

    private void OfferBomOption(SldWorks swApp, string filePath)
    {
        try
        {
            Console.WriteLine("Would you like to view the Bill of Materials (BOM) for this assembly?");

            var menu = MenuFactoryExtensions.CreateStandardMenu("Assembly BOM Options")
                .AddOption("1", "View unsuppressed components only (BOM)", () =>
                {
                    DisplayBom(swApp, filePath, includeSuppressed: false);
                    return false;
                })
                .AddOption("2", "View all components (including suppressed)", () =>
                {
                    DisplayBom(swApp, filePath, includeSuppressed: true);
                    return false;
                })
                .AddOption("3", "Skip BOM", () => false);

            menu.RunMenu();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error offering BOM option: {ex.Message}");
        }
    }

    private void ProcessAssemblyBom(SldWorks swApp, DatabaseGateway dbManager, string filePath)
    {
        Console.WriteLine("\nAssembly detected. Processing BOM...");
        var bomItems = SWMetadataReader.GetBomItems(swApp, filePath, includeSuppressed: false, message => Console.WriteLine(message));

        if (bomItems.Count == 0)
        {
            Logger.LogWarning("No BOM items found in assembly");
            return;
        }

        dbManager.InsertBomItems(filePath, bomItems);
        Logger.LogInfo($"Saved {bomItems.Count} BOM items to database");

        Console.Write("Display BOM? (y/n): ");
        if (Console.ReadLine()?.Trim().ToLower() == "y")
        {
            DisplayBomFromList(bomItems, "Bill of Materials (from database)");
        }
    }



    private void DisplayBom(SldWorks swApp, string filePath, bool includeSuppressed)
    {
        try
        {
            var bomItems = SWMetadataReader.GetBomItems(swApp, filePath, includeSuppressed, message =>
            {
                Logger.WriteAndLogUserMessage(message.TrimStart());
            });

            string title = includeSuppressed
                ? "Complete Component List (Including Suppressed)"
                : "Bill of Materials (Unsuppressed Components Only)";

            DisplayBomFromList(bomItems, title);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating BOM: {ex.Message}");
        }
    }

    public static void DisplayBomFromList(List<BomItem> bomItems, string title)
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

    private void DisplayMetadata(string filePath, Dictionary<string, string> metadata)
    {
        Console.WriteLine($"\n--- Metadata for: {Path.GetFileName(filePath)} ---");

        if (metadata.Count == 0)
        {
            Console.WriteLine("No metadata found.");
            return;
        }

        DisplayGroup("File Information", metadata, new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified" });
        DisplayGroup("Document Properties", metadata, new[] { "Title", "Author", "Subject", "Comments", "Keywords" });

        if (metadata.TryGetValue("Material", out var material))
        {
            Console.WriteLine($"\nMaterial: {material}");
        }

        if (metadata.TryGetValue("ComponentCount", out var componentCount))
        {
            Console.WriteLine($"\nComponent Count: {componentCount}");
        }

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

    private void DisplayGroup(string groupName, Dictionary<string, string> metadata, string[] keys)
    {
        var entries = new List<string>();
        foreach (var key in keys)
        {
            if (metadata.TryGetValue(key, out var value))
            {
                entries.Add($"  {key}: {value}");
            }
        }

        if (entries.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            entries.ForEach(Console.WriteLine);
        }
    }

    private void DisplayGroup(string groupName, Dictionary<string, string> metadata, Func<string, bool> filter)
    {
        var entries = new List<string>();
        foreach (var kvp in metadata)
        {
            if (filter(kvp.Key))
            {
                entries.Add($"  {kvp.Key}: {kvp.Value}");
            }
        }

        if (entries.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            entries.ForEach(Console.WriteLine);
        }
    }

    private bool IsStandardProperty(string key)
    {
        var standardProps = new[]
        {
            "FileName",
            "FilePath",
            "DocumentType",
            "FileSize",
            "LastModified",
            "Title",
            "Author",
            "Subject",
            "Comments",
            "Keywords"
        };

        return Array.IndexOf(standardProps, key) >= 0;
    }
}
