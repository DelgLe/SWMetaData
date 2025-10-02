using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public class MainClass
{
    public static void Main(string[] args)
    {
        // Load config and initialize logger
        var config = ConfigManager.LoadConfigFromJSON();
        Logger.Initialize(config);
        Logger.WriteAndLogUserMessage("SolidWorks Metadata Reader");
        Logger.LogRuntime("Application starting", "Main");

        SldWorks? swApp = null;
        try
        {
            Logger.LogRuntime("Initializing SolidWorks connection", "Main");
            swApp = SWMetadataReader.CreateSolidWorksInstance();
            
            // Performance optimizations - disable UI updates and interactions
            swApp.Visible = false;                           // Hide SolidWorks UI window
            swApp.UserControl = false;                       // Prevent user interaction during processing
            
            Logger.WriteAndLogUserMessage("SolidWorks connection established with performance optimizations", "Main");

            CommandInterface.RunInteractiveLoop(swApp);

        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Main - Fatal error");
        }
        finally
        {
            if (swApp != null)
            {
                Logger.LogRuntime("Cleaning up SolidWorks connection", "Main");
                SWMetadataReader.CleanupSolidWorks(swApp);
            }
            Logger.LogRuntime("Application closed", "Main");
            Console.ReadKey();
        }
    }
}
