using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public class InteractiveFileProcessor
{
    private readonly AppConfig? _config;

    public InteractiveFileProcessor(AppConfig? config)
    {
        _config = config;
    }

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

            if (ShouldOfferBom(metadata))
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

    public void ProcessFileWithDatabase(SldWorks swApp, string databasePath)
    {
        if (string.IsNullOrWhiteSpace(databasePath))
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

            using var dbManager = new DatabaseGateway(databasePath);
            Console.WriteLine("Saving to database...");
            long documentId = dbManager.InsertDocumentMetadata(metadata);
            Console.WriteLine($"Document saved with ID: {documentId}");

            if (ShouldProcessBom(metadata))
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

    private bool ShouldOfferBom(Dictionary<string, string> metadata)
    {
        if (!metadata.TryGetValue("DocumentType", out var docType))
        {
            return false;
        }

        if (!docType.Contains("ASSEMBLY", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return _config?.ProcessBomForAssemblies ?? true;
    }

    private bool ShouldProcessBom(Dictionary<string, string> metadata)
    {
        return ShouldOfferBom(metadata);
    }

    private void OfferBomOption(SldWorks swApp, string filePath)
    {
        try
        {
            Console.WriteLine("\n--- Assembly BOM Options ---");
            Console.WriteLine("Would you like to view the Bill of Materials (BOM) for this assembly?");
            Console.WriteLine("1. View unsuppressed components only (BOM)");
            Console.WriteLine("2. View all components (including suppressed)");
            Console.WriteLine("3. Skip BOM");
            Console.Write("Enter your choice (1-3): ");

            string choice = Console.ReadLine()?.Trim() ?? string.Empty;

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

    private void ProcessAssemblyBom(SldWorks swApp, DatabaseGateway dbManager, string filePath)
    {
        Console.WriteLine("\nAssembly detected. Processing BOM...");

        ModelDoc2? swModel = null;
        try
        {
            int errors = 0;
            int warnings = 0;

            var documentType = SWDocumentManager.GetDocumentType(filePath);
            Logger.LogInfo($"Opening {documentType} for BOM processing: {Path.GetFileName(filePath)}");

            swModel = swApp.OpenDoc6(
                filePath,
                (int)documentType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                string.Empty,
                ref errors,
                ref warnings);

            if (swModel != null)
            {
                var bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                if (bomItems.Count > 0)
                {
                    dbManager.InsertBomItems(filePath, bomItems);
                    Logger.LogInfo($"Saved {bomItems.Count} BOM items to database");

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

                if (documentType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    SWDocumentManager.ForceCleanupHangingDocuments(swApp, "Assembly failed to open - ");
                }
            }
        }
        finally
        {
            SWDocumentManager.CloseDocumentSafely(swApp, swModel);
        }
    }

    private void DisplayBom(SldWorks swApp, string filePath, bool includeSuppressed)
    {
        ModelDoc2? swModel = null;
        try
        {
            int errors = 0;
            int warnings = 0;

            var docType = SWDocumentManager.GetDocumentType(filePath);
            Logger.WriteAndLogUserMessage($"Opening {docType} for BOM display: {Path.GetFileName(filePath)}");

            swModel = swApp.OpenDoc6(
                filePath,
                (int)docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                string.Empty,
                ref errors,
                ref warnings);

            if (swModel == null)
            {
                Console.WriteLine($"Error: Could not open assembly for BOM generation (Errors: {errors}, Warnings: {warnings}).");

                if (docType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    SWDocumentManager.ForceCleanupHangingDocuments(swApp, "Assembly failed to open - ");
                }

                return;
            }

            List<BomItem> bomItems;
            string title;

            if (includeSuppressed)
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
            SWDocumentManager.CloseDocumentSafely(swApp, swModel);
        }
    }

    private void DisplayBomFromList(List<BomItem> bomItems, string title)
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
