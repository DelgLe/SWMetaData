using System;
using SolidWorks.Interop.sldworks;

public class MainClass
{
    public static void Main(string[] args)
    {
        // Load config and initialize logger
        var config = ConfigManager.LoadConfig();
        Logger.Initialize(config);
        Logger.LogRuntime("Application starting", "Main");
        
        Console.WriteLine("SolidWorks Metadata Reader");
        Console.WriteLine("==========================");
        Console.WriteLine("Enter file paths to read metadata, or type 'exit' to quit.\n");

        SldWorks? swApp = null;
        try
        {
            Console.WriteLine("Initializing SolidWorks connection...");
            Logger.LogRuntime("Initializing SolidWorks connection", "Main");
            swApp = SWMetadataReader.CreateSolidWorksInstance();
            swApp.Visible = false;
            Console.WriteLine("SolidWorks connection established\n");
            Logger.LogInfo("SolidWorks connection established", "Main");

            CommandInterface.RunInteractiveLoop(swApp);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
            Logger.LogException(ex, "Main - Fatal error");
        }
        finally
        {
            if (swApp != null)
            {
                Logger.LogRuntime("Cleaning up SolidWorks connection", "Main");
                SWMetadataReader.CleanupSolidWorks(swApp);
            }
            Console.WriteLine("\nApplication closed. Press any key to exit...");
            Logger.LogRuntime("Application closed", "Main");
            Console.ReadKey();
        }
    }
}
