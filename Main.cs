using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public class MainClass
{
    public static void Main(string[] args)
    {
        Console.WriteLine("SolidWorks Metadata Reader");
        Console.WriteLine("==========================");
        Console.WriteLine("Enter file paths to read metadata, or type 'exit' to quit.\n");

        SldWorks swApp = null;
        try
        {
            Console.WriteLine("Initializing SolidWorks connection...");
            swApp = SolidWorksMetadataReader.CreateSolidWorksInstance();
            swApp.Visible = false;
            Console.WriteLine("SolidWorks connection established\n");

            RunInteractiveLoop(swApp);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
        finally
        {
            SolidWorksMetadataReader.CleanupSolidWorks(swApp);
            Console.WriteLine("\nApplication closed. Press any key to exit...");
            Console.ReadKey();
        }
    }

    private static void RunInteractiveLoop(SldWorks swApp)
    {
        while (true)
        {
            Console.Write("Enter file path (or 'exit' to quit): ");
            string input = Console.ReadLine()?.Trim() ?? "";

            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter a file path or 'exit'.\n");
                continue;
            }

            if (input.Equals("exit", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Exiting application...");
                break;
            }



            try
            {
                var metadata = SolidWorksMetadataReader.ReadMetadata(swApp, input);
                DisplayMetadata(input, metadata);
                
                // Check if this is an assembly and offer BOM option
                if (metadata.TryGetValue("DocumentType", out var docType) && 
                    docType.Contains("ASSEMBLY"))
                {
                    OfferBomOption(swApp, input);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                    Console.WriteLine($"Details: {ex.InnerException.Message}");
            }

            Console.WriteLine();
        }
    }

    private static void DisplayMetadata(string filePath, Dictionary<string, string> metadata)
    {
        Console.WriteLine($"\n--- Metadata for: {System.IO.Path.GetFileName(filePath)} ---");

        if (metadata.Count == 0)
        {
            Console.WriteLine("No metadata found.");
            return;
        }

        DisplayGroup("File Information", metadata, new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified" });
        DisplayGroup("Document Properties", metadata, new[] { "Title", "Author", "Subject", "Comments", "Keywords" });

        // Dedicated section for material and components
        if (metadata.ContainsKey("Material"))
        {
            Console.WriteLine($"\nMaterial: {metadata["Material"]}");
        }
        
        if (metadata.ContainsKey("ComponentCount"))
        {
            Console.WriteLine($"\nComponent Count: {metadata["ComponentCount"]}");
        }

        // Dedicated section for configurations (robust display)
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

    private static void DisplayGroup(string groupName, Dictionary<string, string> metadata, string[] keys)
    {
        var found = new List<string>();
        foreach (string key in keys)
        {
            if (metadata.ContainsKey(key))
            {
                found.Add($"  {key}: {metadata[key]}");
            }
        }

        if (found.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            found.ForEach(Console.WriteLine);
        }
    }

    private static void DisplayGroup(string groupName, Dictionary<string, string> metadata, Func<string, bool> filter)
    {
        var customProps = new List<string>();
        foreach (var kvp in metadata)
        {
            if (filter(kvp.Key))
            {
                customProps.Add($"  {kvp.Key}: {kvp.Value}");
            }
        }

        if (customProps.Count > 0)
        {
            Console.WriteLine($"\n{groupName}:");
            customProps.ForEach(Console.WriteLine);
        }
    }

    private static void OfferBomOption(SldWorks swApp, string filePath)
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

    private static void DisplayBom(SldWorks swApp, string filePath, bool includeSupressed)
    {
        try
        {
            // Re-open the document (it should already be cached/quick to open)
            ModelDoc2? swModel = null;
            int errors = 0, warnings = 0;
            
            swDocumentTypes_e docType = SolidWorksMetadataReader.GetDocumentType(filePath);
            swModel = swApp.OpenDoc6(filePath, (int)docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            if (swModel == null)
            {
                Console.WriteLine("Error: Could not open assembly for BOM generation.");
                return;
            }

            List<BomItem> bomItems;
            string title;

            if (includeSupressed)
            {
                bomItems = AssemblyTraverser.GetAllComponents(swModel, true);
                title = "Complete Component List (Including Suppressed)";
            }
            else
            {
                bomItems = AssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                title = "Bill of Materials (Unsuppressed Components Only)";
            }

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

            // Close the document
            if (swModel != null)
                swApp.CloseDoc(swModel.GetTitle());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating BOM: {ex.Message}");
        }
    }

    private static bool IsStandardProperty(string key)
    {
        var standardProps = new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified",
                                   "Title", "Author", "Subject", "Comments", "Keywords" };
        return Array.IndexOf(standardProps, key) >= 0;
    }
}