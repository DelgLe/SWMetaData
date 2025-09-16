using SolidWorks.Interop.sldworks;

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
            string input = Console.ReadLine()?.Trim();

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
        DisplayGroup("Custom Properties", metadata, key => !IsStandardProperty(key));

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
}