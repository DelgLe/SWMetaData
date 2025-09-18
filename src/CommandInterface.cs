using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class CommandInterface
{
    private static string? _databasePath = null;

    public static void RunInteractiveLoop(SldWorks swApp)
    {
        // Show initial options
        ShowMainMenu();

        while (true)
        {
            Console.Write("\nEnter your choice: ");
            string input = Console.ReadLine()?.Trim() ?? "";

            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("Please enter a valid option.\n");
                ShowMainMenu();
                continue;
            }

            if (input.Equals("exit", StringComparison.OrdinalIgnoreCase) || input == "4")
            {
                Console.WriteLine("Exiting application...");
                break;
            }

            switch (input)
            {
                case "1":
                    ProcessSingleFile(swApp);
                    break;
                case "2":
                    SetupDatabase();
                    break;
                case "3":
                    ProcessFileWithDatabase(swApp);
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
                    ShowMainMenu();
                    break;
            }
        }
    }

    private static void ShowMainMenu()
    {
        Console.WriteLine("\n=== SolidWorks Metadata Reader ===");
        Console.WriteLine("1. Process single file (display only)");
        Console.WriteLine("2. Setup/Create database");
        Console.WriteLine("3. Process file and save to database");
        Console.WriteLine("4. Exit");
        Console.WriteLine($"Current database: {(_databasePath ?? "Not set")}");
    }

    private static void ProcessSingleFile(SldWorks swApp)
    {
        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
            DisplayMetadata(filePath, metadata);
            
            // Check if this is an assembly and offer BOM option
            if (metadata.TryGetValue("DocumentType", out var docType) && 
                docType.Contains("ASSEMBLY"))
            {
                OfferBomOption(swApp, filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"Details: {ex.InnerException.Message}");
        }
    }

    private static void SetupDatabase()
    {
        Console.WriteLine("\n=== Database Setup ===");
        Console.Write("Enter database file name (will be created in current folder): ");
        string dbName = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(dbName))
        {
            Console.WriteLine("Database name cannot be empty.");
            return;
        }

        // Ensure .db extension
        if (!dbName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
        {
            dbName += ".db";
        }

        try
        {
            _databasePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), dbName);
            
            // Create database and initialize tables
            using (var dbManager = new SWDatabaseManager(_databasePath))
            {
                Console.WriteLine($"Database created successfully: {_databasePath}");
                Console.WriteLine("Tables initialized:");
                Console.WriteLine("  - sw_documents (document metadata)");
                Console.WriteLine("  - sw_custom_properties (custom properties)");
                Console.WriteLine("  - sw_bom_items (bill of materials)");
                Console.WriteLine("  - sw_configurations (part/assembly configurations)");
                Console.WriteLine("  - sw_materials (material information)");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating database: {ex.Message}");
            _databasePath = null;
        }
    }

    private static void ProcessFileWithDatabase(SldWorks swApp)
    {
        if (string.IsNullOrEmpty(_databasePath))
        {
            Console.WriteLine("No database configured. Please setup database first (option 2).");
            return;
        }

        Console.Write("Enter file path: ");
        string filePath = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(filePath))
        {
            Console.WriteLine("Please enter a valid file path.");
            return;
        }

        try
        {
            Console.WriteLine("Reading metadata from SolidWorks file...");
            var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
            
            // Display the metadata
            DisplayMetadata(filePath, metadata);

            // Save to database
            using var dbManager = new SWDatabaseManager(_databasePath);
            Console.WriteLine("Saving to database...");
            long documentId = dbManager.InsertDocumentMetadata(metadata);
            Console.WriteLine($"Document saved with ID: {documentId}");

            // Handle BOM for assemblies
            if (metadata.TryGetValue("DocumentType", out var docType) && 
                docType.Contains("ASSEMBLY"))
            {
                Console.WriteLine("\nAssembly detected. Processing BOM...");
                
                // Re-open document for BOM processing
                ModelDoc2? swModel = null;
                int errors = 0, warnings = 0;
                
                swDocumentTypes_e documentType = SWMetadataReader.GetDocumentType(filePath);
                swModel = swApp.OpenDoc6(filePath, (int)documentType,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                if (swModel != null)
                {
                    try
                    {
                        var bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                        if (bomItems.Count > 0)
                        {
                            dbManager.InsertBomItems(filePath, bomItems);
                            Console.WriteLine($"Saved {bomItems.Count} BOM items to database");
                            
                            // Offer to display BOM
                            Console.Write("Display BOM? (y/n): ");
                            if (Console.ReadLine()?.Trim().ToLower() == "y")
                            {
                                DisplayBomFromList(bomItems, "Bill of Materials (from database)");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No BOM items found in assembly");
                        }
                    }
                    finally
                    {
                        swApp.CloseDoc(swModel.GetTitle());
                    }
                }
                else
                {
                    Console.WriteLine("Could not re-open document for BOM processing");
                }
            }

            Console.WriteLine("File processing complete!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"Details: {ex.InnerException.Message}");
        }
    }

    public static void DisplayMetadata(string filePath, Dictionary<string, string> metadata)
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

    private static bool IsStandardProperty(string key)
    {
        var standardProps = new[] { "FileName", "FilePath", "DocumentType", "FileSize", "LastModified",
                                   "Title", "Author", "Subject", "Comments", "Keywords" };
        return Array.IndexOf(standardProps, key) >= 0;
    }

    public static void OfferBomOption(SldWorks swApp, string filePath)
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

    public static void DisplayBom(SldWorks swApp, string filePath, bool includeSupressed)
    {
        try
        {
            // Re-open the document (it should already be cached/quick to open)
            ModelDoc2? swModel = null;
            int errors = 0, warnings = 0;
            
            swDocumentTypes_e docType = SWMetadataReader.GetDocumentType(filePath);
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
                bomItems = SWAssemblyTraverser.GetAllComponents(swModel, true);
                title = "Complete Component List (Including Suppressed)";
            }
            else
            {
                bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                title = "Bill of Materials (Unsuppressed Components Only)";
            }

            DisplayBomFromList(bomItems, title);

            // Close the document
            if (swModel != null)
                swApp.CloseDoc(swModel.GetTitle());
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
    }
}
