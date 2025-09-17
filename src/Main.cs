using System;
using SolidWorks.Interop.sldworks;

public class MainClass
{
    public static void Main(string[] args)
    {
        Console.WriteLine("SolidWorks Metadata Reader");
        Console.WriteLine("==========================");
        Console.WriteLine("Enter file paths to read metadata, or type 'exit' to quit.\n");

        SldWorks? swApp = null;
        try
        {
            Console.WriteLine("Initializing SolidWorks connection...");
            swApp = SWMetadataReader.CreateSolidWorksInstance();
            swApp.Visible = false;
            Console.WriteLine("SolidWorks connection established\n");

            CommandInterface.RunInteractiveLoop(swApp);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
        finally
        {
            if (swApp != null)
            {
                SWMetadataReader.CleanupSolidWorks(swApp);
            }
            Console.WriteLine("\nApplication closed. Press any key to exit...");
            Console.ReadKey();
        }
    }
}
