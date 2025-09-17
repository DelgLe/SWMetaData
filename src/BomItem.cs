using SolidWorks.Interop.swconst;

public class BomItem
{
    public string ComponentName { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string Configuration { get; set; } = string.Empty;
    public string FilePath { get; set; } = string.Empty;
    public int Quantity { get; set; } = 1;
    public int Level { get; set; } = 0;
    public bool IsSuppressed { get; set; } = false;
    public swComponentSuppressionState_e SuppressionState { get; set; }
}
