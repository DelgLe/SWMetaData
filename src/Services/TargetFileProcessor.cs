
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public class TargetFileProcessor
{
    private readonly string _databasePath;
    private readonly AppConfig? _config;

    public TargetFileProcessor(string databasePath, AppConfig? config)
    {
        if (string.IsNullOrWhiteSpace(databasePath))
            throw new ArgumentException("Database path cannot be null or empty", nameof(databasePath));

        _databasePath = databasePath;
        _config = config;
    }

    public List<TargetFileInfo> GetAllTargetFiles()
    {
        using var dbManager = new DatabaseGateway(_databasePath);
        return dbManager.GetTargetFiles();
    }

    public List<TargetFileInfo> GetValidTargetFiles()
    {
        using var dbManager = new DatabaseGateway(_databasePath);
        return dbManager.GetValidTargetFiles();
    }

    public TargetFileInfo AddTargetFile(string filePath, string? engId, string? drawId, string? notes)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        string normalizedPath = Path.GetFullPath(filePath);

        var targetFile = new TargetFileInfo
        {
            EngID = string.IsNullOrWhiteSpace(engId) ? null : engId,
            FileName = Path.GetFileName(normalizedPath),
            FilePath = normalizedPath,
            DrawID = string.IsNullOrWhiteSpace(drawId) ? null : drawId,
            Notes = string.IsNullOrWhiteSpace(notes) ? null : notes,
            SourceDirectory = Path.GetDirectoryName(normalizedPath),
            FolderName = Path.GetFileName(Path.GetDirectoryName(normalizedPath)),
            FileCount = 1
        };

        using var dbManager = new DatabaseGateway(_databasePath);
        dbManager.AddTargetFile(targetFile);

        return targetFile;
    }

    public bool RemoveTargetFile(int targetId)
    {
        using var dbManager = new DatabaseGateway(_databasePath);
        var existingFiles = dbManager.GetTargetFiles();
        bool exists = existingFiles.Exists(t => t.TargetID == targetId);

        if (!exists)
            return false;

        dbManager.RemoveTargetFile(targetId);
        return true;
    }

    public TargetFileValidationResult ValidateTargetFiles()
    {
        var allFiles = GetAllTargetFiles();
        var missingFiles = new List<TargetFileInfo>();

        foreach (var file in allFiles)
        {
            if (string.IsNullOrEmpty(file.FilePath) || !File.Exists(file.FilePath))
            {
                missingFiles.Add(file);
            }
        }

        return new TargetFileValidationResult
        {
            TotalCount = allFiles.Count,
            ValidCount = allFiles.Count - missingFiles.Count,
            MissingFiles = missingFiles
        };
    }

    public BatchProcessingResult ProcessFiles(
        SldWorks swApp,
        IReadOnlyList<TargetFileInfo> targetFiles,
        Action<string>? progressWriter = null)
    {
        if (swApp == null)
            throw new ArgumentNullException(nameof(swApp));
        if (targetFiles == null)
            throw new ArgumentNullException(nameof(targetFiles));

        progressWriter ??= Console.WriteLine;

        int processed = 0;
        int errors = 0;
        int excluded = 0;
        int total = targetFiles.Count;
        int cleanupInterval = Math.Max(1, _config?.PeriodicCleanupInterval ?? 10);

        using var dbManager = new DatabaseGateway(_databasePath);

        for (int index = 0; index < targetFiles.Count; index++)
        {
            var targetFile = targetFiles[index];
            string? filePath = targetFile.FilePath;

            try
            {
                progressWriter($"\nProcessing {index + 1}/{total}: {targetFile.FileName}");

                if (string.IsNullOrEmpty(filePath))
                {
                    progressWriter("  Skipping - no file path");
                    Logger.LogWarning("Target file skipped due to missing path", targetFile.EngID ?? targetFile.TargetID.ToString());
                    errors++;
                    continue;
                }

                if (SWMetadataReader.IsFileExcluded(filePath))
                {
                    Logger.LogInfo($"Skipped excluded file: {targetFile.FileName}", Path.GetExtension(filePath));
                    progressWriter("  Skipped - file extension is excluded");
                    excluded++;
                    errors++;
                    continue;
                }

                var metadata = SWMetadataReader.ReadMetadata(swApp, filePath);
                long documentId = dbManager.InsertDocumentMetadata(metadata);
                progressWriter($"  Success! Document ID: {documentId}");

                processed++;

                if (ShouldProcessBom(metadata))
                {
                    ProcessAssemblyBom(swApp, dbManager, filePath, progressWriter);
                }

                if (processed % cleanupInterval == 0)
                {
                    SWDocumentManager.ForceCleanupHangingDocuments(swApp, $"  Periodic cleanup after {processed} files - ");
                    SWMetadataReader.PeriodicCleanup(processed, "BatchProcessing");
                }
            }
            catch (Exception ex)
            {
                progressWriter($"  Error: {ex.Message}");
                Logger.LogException(ex, $"Batch processing error - {filePath}");
                errors++;

                SWDocumentManager.ForceCleanupHangingDocuments(swApp, "  Error cleanup - ");
                SWMetadataReader.PeriodicCleanup(processed + errors, "ErrorHandling");
            }
        }

        return new BatchProcessingResult
        {
            TotalFiles = total,
            ProcessedFiles = processed,
            ErrorCount = errors,
            ExcludedCount = excluded
        };
    }

    private bool ShouldProcessBom(Dictionary<string, string> metadata)
    {
        if (!(_config?.ProcessBomForAssemblies ?? true))
            return false;

        if (!metadata.TryGetValue("DocumentType", out var docType))
            return false;

        return docType.Contains("ASSEMBLY", StringComparison.OrdinalIgnoreCase);
    }

    private void ProcessAssemblyBom(
        SldWorks swApp,
        DatabaseGateway dbManager,
        string filePath,
        Action<string> progressWriter)
    {
        ModelDoc2? swModel = null;
        try
        {
            int errors = 0, warnings = 0;
            var documentType = SWDocumentManager.GetDocumentType(filePath);
            progressWriter($"  Opening {documentType}: {Path.GetFileName(filePath)}");

            swModel = swApp.OpenDoc6(
                filePath,
                (int)documentType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                string.Empty,
                ref errors,
                ref warnings);

            if (swModel != null)
            {
                var bomItems = SWAssemblyTraverser.GetUnsuppressedComponents(swModel, true);
                if (bomItems.Count > 0)
                {
                    dbManager.InsertBomItems(filePath, bomItems);
                    progressWriter($"  Saved {bomItems.Count} BOM items");
                }
                else
                {
                    progressWriter("  No BOM items found");
                }
            }
            else
            {
                progressWriter($"  Could not open assembly (Errors: {errors}, Warnings: {warnings})");

                if (documentType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    SWDocumentManager.ForceCleanupHangingDocuments(swApp, "  Assembly failed to open - ");
                }
            }
        }
        catch (Exception ex)
        {
            progressWriter($"  BOM processing error: {ex.Message}");
            Logger.LogException(ex, $"ProcessAssemblyBom - {filePath}");
        }
        finally
        {
            SWDocumentManager.CloseDocumentSafely(swApp, swModel);
        }
    }

    // UI Methods for interactive target file management
    public void ViewTargetFilesInteractive()
    {
        try
        {
            var targetFiles = GetAllTargetFiles();
            
            if (targetFiles.Count == 0)
            {
                Console.WriteLine("No target files found in database.");
                return;
            }

            Console.WriteLine($"\n=== Target Files ({targetFiles.Count} files) ===");
            Console.WriteLine(new string('-', 100));
            Console.WriteLine("{0,-3} {1,-15} {2,-30} {3,-40}", "ID", "EngID", "File Name", "File Path");
            Console.WriteLine(new string('-', 100));

            foreach (var file in targetFiles)
            {
                string engId = file.EngID ?? "N/A";
                string fileName = file.FileName ?? "N/A";
                string filePath = file.FilePath ?? "N/A";

                // Truncate long paths for display
                if (filePath.Length > 40)
                    filePath = "..." + filePath.Substring(filePath.Length - 37);

                Console.WriteLine("{0,-3} {1,-15} {2,-30} {3,-40}",
                    file.TargetID, engId, fileName, filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error viewing target files: {ex.Message}");
        }
    }

    public void AddTargetFileInteractive()
    {
        try
        {
            Console.Write("Enter file path: ");
            string filePath = Console.ReadLine()?.Trim() ?? "";

            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("File path cannot be empty.");
                return;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Warning: File does not exist at the specified path.");
                Console.Write("Continue anyway? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() != "y")
                    return;
            }

            Console.Write("Enter Engineering ID (optional): ");
            string? engId = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(engId)) engId = null;

            Console.Write("Enter Drawing ID (optional): ");
            string? drawId = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(drawId)) drawId = null;

            Console.Write("Enter notes (optional): ");
            string? notes = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(notes)) notes = null;

            AddTargetFile(filePath, engId, drawId, notes);
            Console.WriteLine("Target file added successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error adding target file: {ex.Message}");
        }
    }

    public void RemoveTargetFileInteractive()
    {
        try
        {
            var targetFiles = GetAllTargetFiles();

            if (targetFiles.Count == 0)
            {
                Console.WriteLine("No target files to remove.");
                return;
            }

            Console.WriteLine("\nTarget Files:");
            foreach (var file in targetFiles.Take(20)) // Show first 20
            {
                Console.WriteLine($"{file.TargetID}: {file.FileName} ({file.EngID ?? "No EngID"})");
            }

            if (targetFiles.Count > 20)
                Console.WriteLine($"... and {targetFiles.Count - 20} more files");

            Console.Write("Enter Target ID to remove: ");
            if (int.TryParse(Console.ReadLine(), out int targetId))
            {
                if (RemoveTargetFile(targetId))
                {
                    Console.WriteLine("Target file removed successfully!");
                }
                else
                {
                    Console.WriteLine("Target file not found.");
                }
            }
            else
            {
                Console.WriteLine("Invalid Target ID.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error removing target file: {ex.Message}");
        }
    }

    public void ValidateTargetFilesInteractive()
    {
        try
        {
            var result = ValidateTargetFiles();

            Logger.WriteAndLogUserMessage($"\nValidation Results:");
            Logger.WriteAndLogUserMessage($"Total target files: {result.TotalCount}");
            Logger.WriteAndLogUserMessage($"Valid files (exist on disk): {result.ValidCount}");
            Logger.WriteAndLogUserMessage($"Missing files: {result.TotalCount - result.ValidCount}");

            if (result.MissingFiles.Count > 0)
            {
                Console.WriteLine("\nMissing files:");
                foreach (var file in result.MissingFiles)
                {
                    Console.WriteLine($"  ID {file.TargetID}: {file.FilePath}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error validating target files: {ex.Message}");
        }
    }
}

public sealed class TargetFileValidationResult
{
    public int TotalCount { get; init; }
    public int ValidCount { get; init; }
    public IReadOnlyList<TargetFileInfo> MissingFiles { get; init; } = Array.Empty<TargetFileInfo>();
}

public sealed class BatchProcessingResult
{
    public int TotalFiles { get; init; }
    public int ProcessedFiles { get; init; }
    public int ErrorCount { get; init; }
    public int ExcludedCount { get; init; }
}
