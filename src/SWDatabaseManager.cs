using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

public class SWDatabaseManager : IDisposable
{
    private readonly string _connectionString;
    private SQLiteConnection _connection;
    private bool _disposed = false;

    public SWDatabaseManager(string databasePath)
    {
        if (string.IsNullOrWhiteSpace(databasePath))
            throw new ArgumentException("Database path cannot be null or empty", nameof(databasePath));

        _connectionString = $"Data Source={databasePath};Version=3;";
        InitializeDatabase();
    }

    /// <summary>
    /// Initialize the database and create tables if they don't exist
    /// </summary>
    private void InitializeDatabase()
    {
        try
        {
            _connection = new SQLiteConnection(_connectionString);
            _connection.Open();
            CreateTables();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to initialize database: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Create all necessary tables for storing SolidWorks data
    /// </summary>
    private void CreateTables()
    {
        var commands = new[]
        {
            // Assembly/Part metadata table
            @"CREATE TABLE IF NOT EXISTS sw_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL UNIQUE,
                file_name TEXT NOT NULL,
                document_type TEXT NOT NULL,
                file_size INTEGER,
                last_modified DATETIME,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP
            )",

            // Custom properties table
            @"CREATE TABLE IF NOT EXISTS sw_custom_properties (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                property_name TEXT NOT NULL,
                property_value TEXT,
                FOREIGN KEY (document_id) REFERENCES sw_documents (id) ON DELETE CASCADE,
                UNIQUE(document_id, property_name)
            )",

            // BOM items table
            @"CREATE TABLE IF NOT EXISTS sw_bom_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_assembly_id INTEGER NOT NULL,
                component_name TEXT NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                configuration TEXT,
                quantity INTEGER DEFAULT 1,
                level INTEGER DEFAULT 0,
                is_suppressed BOOLEAN DEFAULT 0,
                suppression_state TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (parent_assembly_id) REFERENCES sw_documents (id) ON DELETE CASCADE
            )",

            // Configurations table
            @"CREATE TABLE IF NOT EXISTS sw_configurations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                configuration_name TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES sw_documents (id) ON DELETE CASCADE,
                UNIQUE(document_id, configuration_name)
            )",

            // Materials table
            @"CREATE TABLE IF NOT EXISTS sw_materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_id INTEGER NOT NULL,
                material_name TEXT NOT NULL,
                FOREIGN KEY (document_id) REFERENCES sw_documents (id) ON DELETE CASCADE
            )",

            // Create indexes for better performance
            @"CREATE INDEX IF NOT EXISTS idx_documents_filepath ON sw_documents(file_path)",
            @"CREATE INDEX IF NOT EXISTS idx_custom_props_docid ON sw_custom_properties(document_id)",
            @"CREATE INDEX IF NOT EXISTS idx_bom_items_assembly ON sw_bom_items(parent_assembly_id)",
            @"CREATE INDEX IF NOT EXISTS idx_configurations_docid ON sw_configurations(document_id)",
            @"CREATE INDEX IF NOT EXISTS idx_materials_docid ON sw_materials(document_id)"
        };

        foreach (string commandText in commands)
        {
            using var command = new SQLiteCommand(commandText, _connection);
            command.ExecuteNonQuery();
        }
    }

    /// <summary>
    /// Insert document metadata into the database
    /// </summary>
    /// <param name="metadata">Dictionary containing document metadata</param>
    /// <returns>The document ID of the inserted/updated record</returns>
    public long InsertDocumentMetadata(Dictionary<string, string> metadata)
    {
        if (metadata == null)
            throw new ArgumentNullException(nameof(metadata));

        if (!metadata.ContainsKey("FilePath"))
            throw new ArgumentException("Metadata must contain 'FilePath' key", nameof(metadata));

        using var transaction = _connection.BeginTransaction();
        try
        {
            // Insert or update document record
            long documentId = InsertOrUpdateDocument(metadata);

            // Insert custom properties
            InsertCustomProperties(documentId, metadata);

            // Insert configurations if present
            if (metadata.ContainsKey("Configurations"))
            {
                InsertConfigurations(documentId, metadata["Configurations"]);
            }

            // Insert material if present
            if (metadata.ContainsKey("Material"))
            {
                InsertMaterial(documentId, metadata["Material"]);
            }

            transaction.Commit();
            return documentId;
        }
        catch (Exception)
        {
            transaction.Rollback();
            throw;
        }
    }

    /// <summary>
    /// Insert BOM items for an assembly
    /// </summary>
    /// <param name="assemblyFilePath">Path to the assembly file</param>
    /// <param name="bomItems">List of BOM items to insert</param>
    public void InsertBomItems(string assemblyFilePath, List<BomItem> bomItems)
    {
        if (string.IsNullOrWhiteSpace(assemblyFilePath))
            throw new ArgumentException("Assembly file path cannot be null or empty", nameof(assemblyFilePath));

        if (bomItems == null || bomItems.Count == 0)
            return;

        // Get the assembly document ID
        long assemblyId = GetDocumentId(assemblyFilePath);
        if (assemblyId == 0)
            throw new InvalidOperationException($"Assembly document not found in database: {assemblyFilePath}");

        using var transaction = _connection.BeginTransaction();
        try
        {
            // Clear existing BOM items for this assembly
            ClearBomItems(assemblyId);

            // Insert new BOM items
            const string sql = @"
                INSERT INTO sw_bom_items 
                (parent_assembly_id, component_name, file_name, file_path, configuration, 
                 quantity, level, is_suppressed, suppression_state)
                VALUES (@assemblyId, @componentName, @fileName, @filePath, @configuration, 
                        @quantity, @level, @isSuppressed, @suppressionState)";

            using var command = new SQLiteCommand(sql, _connection);
            command.Parameters.Add("@assemblyId", System.Data.DbType.Int64);
            command.Parameters.Add("@componentName", System.Data.DbType.String);
            command.Parameters.Add("@fileName", System.Data.DbType.String);
            command.Parameters.Add("@filePath", System.Data.DbType.String);
            command.Parameters.Add("@configuration", System.Data.DbType.String);
            command.Parameters.Add("@quantity", System.Data.DbType.Int32);
            command.Parameters.Add("@level", System.Data.DbType.Int32);
            command.Parameters.Add("@isSuppressed", System.Data.DbType.Boolean);
            command.Parameters.Add("@suppressionState", System.Data.DbType.String);

            foreach (var item in bomItems)
            {
                command.Parameters["@assemblyId"].Value = assemblyId;
                command.Parameters["@componentName"].Value = item.ComponentName ?? string.Empty;
                command.Parameters["@fileName"].Value = item.FileName ?? string.Empty;
                command.Parameters["@filePath"].Value = item.FilePath ?? string.Empty;
                command.Parameters["@configuration"].Value = item.Configuration ?? string.Empty;
                command.Parameters["@quantity"].Value = item.Quantity;
                command.Parameters["@level"].Value = item.Level;
                command.Parameters["@isSuppressed"].Value = item.IsSuppressed;
                command.Parameters["@suppressionState"].Value = item.SuppressionState.ToString();

                command.ExecuteNonQuery();
            }

            transaction.Commit();
        }
        catch (Exception)
        {
            transaction.Rollback();
            throw;
        }
    }

    /// <summary>
    /// Get BOM items for a specific assembly
    /// </summary>
    /// <param name="assemblyFilePath">Path to the assembly file</param>
    /// <returns>List of BOM items from the database</returns>
    public List<BomItem> GetBomItems(string assemblyFilePath)
    {
        if (string.IsNullOrWhiteSpace(assemblyFilePath))
            throw new ArgumentException("Assembly file path cannot be null or empty", nameof(assemblyFilePath));

        long assemblyId = GetDocumentId(assemblyFilePath);
        if (assemblyId == 0)
            return new List<BomItem>();

        const string sql = @"
            SELECT component_name, file_name, file_path, configuration, 
                   quantity, level, is_suppressed, suppression_state
            FROM sw_bom_items 
            WHERE parent_assembly_id = @assemblyId
            ORDER BY level, component_name";

        var bomItems = new List<BomItem>();

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@assemblyId", assemblyId);

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var item = new BomItem
            {
                ComponentName = reader["component_name"].ToString() ?? string.Empty,
                FileName = reader["file_name"].ToString() ?? string.Empty,
                FilePath = reader["file_path"].ToString() ?? string.Empty,
                Configuration = reader["configuration"].ToString() ?? string.Empty,
                Quantity = Convert.ToInt32(reader["quantity"]),
                Level = Convert.ToInt32(reader["level"]),
                IsSuppressed = Convert.ToBoolean(reader["is_suppressed"])
            };

            // Parse suppression state
            if (Enum.TryParse<SolidWorks.Interop.swconst.swComponentSuppressionState_e>(
                reader["suppression_state"].ToString(), out var suppressionState))
            {
                item.SuppressionState = suppressionState;
            }

            bomItems.Add(item);
        }

        return bomItems;
    }

    /// <summary>
    /// Get document metadata from the database
    /// </summary>
    /// <param name="filePath">Path to the document file</param>
    /// <returns>Dictionary containing the metadata, or null if not found</returns>
    public Dictionary<string, string> GetDocumentMetadata(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        const string sql = @"
            SELECT d.*, p.property_name, p.property_value
            FROM sw_documents d
            LEFT JOIN sw_custom_properties p ON d.id = p.document_id
            WHERE d.file_path = @filePath";

        var metadata = new Dictionary<string, string>();
        bool documentFound = false;

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@filePath", filePath);

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (!documentFound)
            {
                // Add document info once
                metadata["FilePath"] = reader["file_path"].ToString() ?? string.Empty;
                metadata["FileName"] = reader["file_name"].ToString() ?? string.Empty;
                metadata["DocumentType"] = reader["document_type"].ToString() ?? string.Empty;
                metadata["FileSize"] = reader["file_size"].ToString() ?? string.Empty;
                metadata["LastModified"] = reader["last_modified"].ToString() ?? string.Empty;
                documentFound = true;
            }

            // Add custom property if present
            string propName = reader["property_name"].ToString();
            string propValue = reader["property_value"].ToString();
            if (!string.IsNullOrEmpty(propName))
            {
                metadata[propName] = propValue ?? string.Empty;
            }
        }

        return documentFound ? metadata : null;
    }

    #region Target Files Management

    /// <summary>
    /// Get all target files from the target_files table
    /// </summary>
    /// <returns>List of target file information</returns>
    public List<TargetFileInfo> GetTargetFiles()
    {
        const string sql = @"
            SELECT TargetID, EngID, file_name, file_path, DrawID, notes, 
                   source_directory, folder_name, file_count
            FROM target_files 
            ORDER BY TargetID";

        var targetFiles = new List<TargetFileInfo>();

        using var command = new SQLiteCommand(sql, _connection);
        using var reader = command.ExecuteReader();
        
        while (reader.Read())
        {
            targetFiles.Add(new TargetFileInfo
            {
                TargetID = Convert.ToInt32(reader["TargetID"]),
                EngID = reader["EngID"]?.ToString(),
                FileName = reader["file_name"]?.ToString(),
                FilePath = reader["file_path"]?.ToString(),
                DrawID = reader["DrawID"]?.ToString(),
                Notes = reader["notes"]?.ToString(),
                SourceDirectory = reader["source_directory"]?.ToString(),
                FolderName = reader["folder_name"]?.ToString(),
                FileCount = Convert.ToInt32(reader["file_count"])
            });
        }

        return targetFiles;
    }

    /// <summary>
    /// Add a new target file to the target_files table
    /// </summary>
    /// <param name="targetFile">Target file information to add</param>
    public void AddTargetFile(TargetFileInfo targetFile)
    {
        if (targetFile == null)
            throw new ArgumentNullException(nameof(targetFile));

        const string sql = @"
            INSERT INTO target_files 
            (EngID, file_name, file_path, DrawID, notes, source_directory, folder_name, file_count)
            VALUES (@engId, @fileName, @filePath, @drawId, @notes, @sourceDirectory, @folderName, @fileCount)";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@engId", targetFile.EngID ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@fileName", targetFile.FileName ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@filePath", targetFile.FilePath ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@drawId", targetFile.DrawID ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@notes", targetFile.Notes ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@sourceDirectory", targetFile.SourceDirectory ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@folderName", targetFile.FolderName ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@fileCount", targetFile.FileCount);
        
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Quick add a target file with just file path
    /// </summary>
    /// <param name="filePath">Path to the SolidWorks file</param>
    /// <param name="engId">Optional engineering ID</param>
    public void AddTargetFile(string filePath, string? engId = null)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        var targetFile = new TargetFileInfo
        {
            EngID = engId,
            FileName = System.IO.Path.GetFileName(filePath),
            FilePath = filePath,
            SourceDirectory = System.IO.Path.GetDirectoryName(filePath),
            FolderName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(filePath)),
            FileCount = 1
        };

        AddTargetFile(targetFile);
    }

    /// <summary>
    /// Update an existing target file
    /// </summary>
    /// <param name="targetFile">Updated target file information</param>
    public void UpdateTargetFile(TargetFileInfo targetFile)
    {
        if (targetFile == null)
            throw new ArgumentNullException(nameof(targetFile));

        const string sql = @"
            UPDATE target_files 
            SET EngID = @engId, file_name = @fileName, file_path = @filePath, 
                DrawID = @drawId, notes = @notes, source_directory = @sourceDirectory, 
                folder_name = @folderName, file_count = @fileCount
            WHERE TargetID = @targetId";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@targetId", targetFile.TargetID);
        command.Parameters.AddWithValue("@engId", targetFile.EngID ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@fileName", targetFile.FileName ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@filePath", targetFile.FilePath ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@drawId", targetFile.DrawID ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@notes", targetFile.Notes ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@sourceDirectory", targetFile.SourceDirectory ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@folderName", targetFile.FolderName ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@fileCount", targetFile.FileCount);
        
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Remove a target file from the target_files table
    /// </summary>
    /// <param name="targetId">Target ID to remove</param>
    public void RemoveTargetFile(int targetId)
    {
        const string sql = "DELETE FROM target_files WHERE TargetID = @targetId";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@targetId", targetId);
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Get target files that exist as actual files on disk
    /// </summary>
    /// <returns>List of valid target files</returns>
    public List<TargetFileInfo> GetValidTargetFiles()
    {
        var allTargets = GetTargetFiles();
        var validTargets = new List<TargetFileInfo>();

        foreach (var target in allTargets)
        {
            if (!string.IsNullOrEmpty(target.FilePath) && System.IO.File.Exists(target.FilePath))
            {
                validTargets.Add(target);
            }
        }

        return validTargets;
    }

    /// <summary>
    /// Get count of target files
    /// </summary>
    /// <returns>Total count of target files</returns>
    public int GetTargetFileCount()
    {
        const string sql = "SELECT COUNT(*) FROM target_files";
        
        using var command = new SQLiteCommand(sql, _connection);
        return Convert.ToInt32(command.ExecuteScalar());
    }

    #endregion

    #region Private Helper Methods

    private long InsertOrUpdateDocument(Dictionary<string, string> metadata)
    {
        const string insertSql = @"
            INSERT OR REPLACE INTO sw_documents 
            (file_path, file_name, document_type, file_size, last_modified)
            VALUES (@filePath, @fileName, @documentType, @fileSize, @lastModified)";

        using var command = new SQLiteCommand(insertSql, _connection);
        command.Parameters.AddWithValue("@filePath", metadata.GetValueOrDefault("FilePath", ""));
        command.Parameters.AddWithValue("@fileName", metadata.GetValueOrDefault("FileName", ""));
        command.Parameters.AddWithValue("@documentType", metadata.GetValueOrDefault("DocumentType", ""));
        command.Parameters.AddWithValue("@fileSize", ParseLong(metadata.GetValueOrDefault("FileSize", "0")));
        command.Parameters.AddWithValue("@lastModified", metadata.GetValueOrDefault("LastModified", ""));

        command.ExecuteNonQuery();

        // Get the document ID
        return GetDocumentId(metadata["FilePath"]);
    }

    private void InsertCustomProperties(long documentId, Dictionary<string, string> metadata)
    {
        // Clear existing custom properties
        using (var deleteCommand = new SQLiteCommand(
            "DELETE FROM sw_custom_properties WHERE document_id = @documentId", _connection))
        {
            deleteCommand.Parameters.AddWithValue("@documentId", documentId);
            deleteCommand.ExecuteNonQuery();
        }

        // Insert custom properties (exclude system fields)
        var systemFields = new HashSet<string> 
        { 
            "FilePath", "FileName", "DocumentType", "FileSize", "LastModified", 
            "Configurations", "Material", "ComponentCount" 
        };

        const string sql = @"
            INSERT INTO sw_custom_properties (document_id, property_name, property_value)
            VALUES (@documentId, @propertyName, @propertyValue)";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.Add("@documentId", System.Data.DbType.Int64);
        command.Parameters.Add("@propertyName", System.Data.DbType.String);
        command.Parameters.Add("@propertyValue", System.Data.DbType.String);

        foreach (var kvp in metadata)
        {
            if (!systemFields.Contains(kvp.Key))
            {
                command.Parameters["@documentId"].Value = documentId;
                command.Parameters["@propertyName"].Value = kvp.Key;
                command.Parameters["@propertyValue"].Value = kvp.Value ?? string.Empty;
                command.ExecuteNonQuery();
            }
        }
    }

    private void InsertConfigurations(long documentId, string configurationsString)
    {
        if (string.IsNullOrWhiteSpace(configurationsString))
            return;

        // Clear existing configurations
        using (var deleteCommand = new SQLiteCommand(
            "DELETE FROM sw_configurations WHERE document_id = @documentId", _connection))
        {
            deleteCommand.Parameters.AddWithValue("@documentId", documentId);
            deleteCommand.ExecuteNonQuery();
        }

        // Insert configurations
        string[] configurations = configurationsString.Split(',');
        const string sql = @"
            INSERT INTO sw_configurations (document_id, configuration_name)
            VALUES (@documentId, @configurationName)";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.Add("@documentId", System.Data.DbType.Int64);
        command.Parameters.Add("@configurationName", System.Data.DbType.String);

        foreach (string config in configurations)
        {
            string trimmedConfig = config.Trim();
            if (!string.IsNullOrEmpty(trimmedConfig))
            {
                command.Parameters["@documentId"].Value = documentId;
                command.Parameters["@configurationName"].Value = trimmedConfig;
                command.ExecuteNonQuery();
            }
        }
    }

    private void InsertMaterial(long documentId, string materialName)
    {
        if (string.IsNullOrWhiteSpace(materialName))
            return;

        // Clear existing material
        using (var deleteCommand = new SQLiteCommand(
            "DELETE FROM sw_materials WHERE document_id = @documentId", _connection))
        {
            deleteCommand.Parameters.AddWithValue("@documentId", documentId);
            deleteCommand.ExecuteNonQuery();
        }

        // Insert material
        const string sql = @"
            INSERT INTO sw_materials (document_id, material_name)
            VALUES (@documentId, @materialName)";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@documentId", documentId);
        command.Parameters.AddWithValue("@materialName", materialName);
        command.ExecuteNonQuery();
    }

    private long GetDocumentId(string filePath)
    {
        const string sql = "SELECT id FROM sw_documents WHERE file_path = @filePath";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@filePath", filePath);

        object result = command.ExecuteScalar();
        return result != null ? Convert.ToInt64(result) : 0;
    }

    private void ClearBomItems(long assemblyId)
    {
        const string sql = "DELETE FROM sw_bom_items WHERE parent_assembly_id = @assemblyId";

        using var command = new SQLiteCommand(sql, _connection);
        command.Parameters.AddWithValue("@assemblyId", assemblyId);
        command.ExecuteNonQuery();
    }

    private static long ParseLong(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return 0;

        // Remove " bytes" suffix if present
        value = value.Replace(" bytes", "").Trim();
        
        return long.TryParse(value, out long result) ? result : 0;
    }

    #endregion

    #region IDisposable Implementation

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed && disposing)
        {
            _connection?.Close();
            _connection?.Dispose();
            _disposed = true;
        }
    }

    #endregion
}
