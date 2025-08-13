using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SolidWorksMetadataReader
{
    public static Dictionary<string, string> ReadMetadata(string filePath)
    {
        SldWorks swApp = null;
        ModelDoc2 swModel = null;
        var metadata = new Dictionary<string, string>();
        
        try
        {
            // Connect to SolidWorks
            swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            swApp.Visible = false; // Run in background
            
            // Open document
            int errors = 0, warnings = 0;
            swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART, 
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            
            if (swModel != null)
            {
                // Get custom properties
                CustomPropertyManager propMgr = swModel.Extension.get_CustomPropertyManager("");
                object[] propNames = propMgr.GetNames();
                
                if (propNames != null)
                {
                    foreach (string propName in propNames)
                    {
                        string propValue = "";
                        propMgr.Get4(propName, false, out propValue, out _);
                        metadata[propName] = propValue;
                    }
                }
                
                // Get summary info
                metadata["Title"] = swModel.SummaryInfo[(int)swSummaryInfoField_e.swSumInfoTitle];
                metadata["Author"] = swModel.SummaryInfo[(int)swSummaryInfoField_e.swSumInfoAuthor];
                metadata["Comments"] = swModel.SummaryInfo[(int)swSummaryInfoField_e.swSumInfoComment];
            }
        }
        finally
        {
            // Cleanup
            if (swModel != null) swApp.CloseDoc(swModel.GetTitle());
            if (swApp != null) Marshal.ReleaseComObject(swApp);
        }
        
        return metadata;
    }
}