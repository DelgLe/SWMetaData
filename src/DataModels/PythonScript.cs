/// <summary>
/// Represents a registered Python script with its path and description
/// </summary>

public class PythonScript(string filePath, string displayName, string description)
{
    public string FilePath { get; } = filePath ?? throw new ArgumentNullException(nameof(filePath));
    public string Description { get; } = description ?? throw new ArgumentNullException(nameof(description));
    public string DisplayName { get; } = displayName ?? throw new ArgumentNullException(nameof(displayName));
    public string FileName => Path.GetFileNameWithoutExtension(FilePath);

}