using System;

/// <summary>
/// Represents information about a target file from the target_files table
/// </summary>
public class TargetFileInfo
{
    public int TargetID { get; set; }
    public string? EngID { get; set; }
    public string? FileName { get; set; }
    public string? FilePath { get; set; }
    public string? DrawID { get; set; }
    public string? Notes { get; set; }
    public string? SourceDirectory { get; set; }
    public string? FolderName { get; set; }
    public int FileCount { get; set; }
}