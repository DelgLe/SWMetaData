using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public static class SWAssemblyTraverser
{
    /// <summary>
    /// Traverses an assembly and returns a list of unsuppressed components (BOM-like list)
    /// </summary>
    /// <param name="swModel">The assembly ModelDoc2 object</param>
    /// <param name="includeSubassemblies">Whether to traverse into subassemblies</param>
    /// <returns>List of BomItem objects representing unsuppressed components</returns>
    public static List<BomItem> GetUnsuppressedComponents(ModelDoc2 swModel, bool includeSubassemblies = true)
    {
        var bomItems = new List<BomItem>();
        
        try
        {
            // Verify this is an assembly
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return bomItems; // Return empty list for non-assemblies
            }

            AssemblyDoc swAssembly = (AssemblyDoc)swModel;
            
            // Get all components at the root level
            object[] rootComponents = (object[])swAssembly.GetComponents(false);
            
            if (rootComponents != null)
            {
                foreach (Component2 comp in rootComponents)
                {
                    TraverseComponent(comp, bomItems, 0, includeSubassemblies);
                }
            }

            // Group by identical components and calculate quantities
            var groupedItems = bomItems
                .Where(item => !item.IsSuppressed)
                .GroupBy(item => new { item.FileName, item.Configuration })
                .Select(group => new BomItem
                {
                    ComponentName = group.First().ComponentName,
                    FileName = group.Key.FileName,
                    Configuration = group.Key.Configuration,
                    FilePath = group.First().FilePath,
                    Quantity = group.Count(),
                    Level = group.Min(item => item.Level), // Use the minimum level for grouped items
                    IsSuppressed = false,
                    SuppressionState = group.First().SuppressionState
                })
                .OrderBy(item => item.Level)
                .ThenBy(item => item.ComponentName)
                .ToList();

            return groupedItems;
        }
        catch (Exception)
        {
            // Return empty list on error
            return bomItems;
        }
    }

    /// <summary>
    /// Gets all components (including suppressed ones) with their suppression states
    /// </summary>
    /// <param name="swModel">The assembly ModelDoc2 object</param>
    /// <param name="includeSubassemblies">Whether to traverse into subassemblies</param>
    /// <returns>List of all BomItem objects with suppression state information</returns>
    public static List<BomItem> GetAllComponents(ModelDoc2 swModel, bool includeSubassemblies = true)
    {
        var bomItems = new List<BomItem>();
        
        try
        {
            // Verify this is an assembly
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return bomItems; // Return empty list for non-assemblies
            }

            AssemblyDoc swAssembly = (AssemblyDoc)swModel;
            
            // Get all components at the root level
            object[] rootComponents = (object[])swAssembly.GetComponents(false);
            
            if (rootComponents != null)
            {
                foreach (Component2 comp in rootComponents)
                {
                    TraverseComponent(comp, bomItems, 0, includeSubassemblies);
                }
            }

            return bomItems.OrderBy(item => item.Level)
                          .ThenBy(item => item.ComponentName)
                          .ToList();
        }
        catch (Exception)
        {
            // Return empty list on error
            return bomItems;
        }
    }

    /// <summary>
    /// Recursively traverses a component and its children
    /// </summary>
    /// <param name="component">The component to traverse</param>
    /// <param name="bomItems">List to add BOM items to</param>
    /// <param name="level">Current depth level</param>
    /// <param name="includeSubassemblies">Whether to traverse into subassemblies</param>
    private static void TraverseComponent(Component2 component, List<BomItem> bomItems, int level, bool includeSubassemblies)
    {
        try
        {
            if (component == null) return;

            // Get component information
            string componentName = component.Name2 ?? string.Empty;
            string fileName = System.IO.Path.GetFileNameWithoutExtension(component.GetPathName()) ?? string.Empty;
            string configuration = component.ReferencedConfiguration ?? string.Empty;
            string filePath = component.GetPathName() ?? string.Empty;
            
            // Get suppression state
            swComponentSuppressionState_e suppressionState = (swComponentSuppressionState_e)component.GetSuppression();
            bool isSuppressed = (suppressionState == swComponentSuppressionState_e.swComponentSuppressed);

            // Create BOM item
            var bomItem = new BomItem
            {
                ComponentName = componentName,
                FileName = fileName,
                Configuration = configuration,
                FilePath = filePath,
                Quantity = 1, // Will be calculated later when grouping
                Level = level,
                IsSuppressed = isSuppressed,
                SuppressionState = suppressionState
            };

            bomItems.Add(bomItem);

            // Traverse children if this is a subassembly and we want to include subassemblies
            if (includeSubassemblies && !isSuppressed)
            {
                object[] childComponents = (object[])component.GetChildren();
                if (childComponents != null)
                {
                    foreach (Component2 childComp in childComponents)
                    {
                        TraverseComponent(childComp, bomItems, level + 1, includeSubassemblies);
                    }
                }
            }
        }
        catch (Exception)
        {
            // Skip this component on error
        }
    }

    /// <summary>
    /// Formats the BOM items for display
    /// </summary>
    /// <param name="bomItems">List of BOM items to format</param>
    /// <param name="includeQuantity">Whether to include quantity in the output</param>
    /// <param name="includeLevel">Whether to include level indentation</param>
    /// <returns>Formatted string representation of the BOM</returns>
    public static string FormatBomList(List<BomItem> bomItems, bool includeQuantity = true, bool includeLevel = true)
    {
        if (bomItems == null || bomItems.Count == 0)
            return "No components found.";

        var result = new System.Text.StringBuilder();
        result.AppendLine("Bill of Materials:");
        result.AppendLine(new string('-', 80));

        foreach (var item in bomItems)
        {
            string indent = includeLevel ? new string(' ', item.Level * 2) : "";
            string quantity = includeQuantity ? $"({item.Quantity}) " : "";
            string suppression = item.IsSuppressed ? " [SUPPRESSED]" : "";
            
            result.AppendLine($"{indent}{quantity}{item.ComponentName} - {item.FileName} [{item.Configuration}]{suppression}");
        }

        return result.ToString();
    }
}
