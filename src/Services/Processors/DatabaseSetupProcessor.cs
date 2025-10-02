
/// <summary>
/// Handles all database setup, initialization, and validation operations
/// </summary>
public class DatabaseSetupProcessor(AppConfig config, Action<string?> onDatabasePathChanged)
{
    private readonly AppConfig _config = config ?? throw new ArgumentNullException(nameof(config));
    private readonly Action<string?> _onDatabasePathChanged = onDatabasePathChanged ?? throw new ArgumentNullException(nameof(onDatabasePathChanged));

    /// <summary>
    /// Main database setup orchestration - handles both config and new database scenarios
    /// </summary>
    public void SetupDatabase()
    {
        // Check if config has a database path
        bool hasConfigDatabase = !string.IsNullOrEmpty(_config.DatabasePath);
        
        if (hasConfigDatabase)
        {
            Console.WriteLine($"Configuration database path found: {_config.DatabasePath}");
            
            var menu = MenuFactoryExtensions.CreateStandardMenu("Database Setup")
                .AddOption("1", "Use existing config database path (setup tables if needed)", () => {
                    SetupConfigDatabase();
                    return false; // Exit menu after selection
                })
                .AddOption("2", "Create new database in current folder", () => {
                    SetupNewDatabase();
                    return false; // Exit menu after selection
                });
                
            menu.RunMenu();
        }
        else
        {
            // No config database, go straight to creating new database
            SetupNewDatabase();
        }
    }

    /// <summary>
    /// Setup database using path from configuration
    /// </summary>
    public void SetupConfigDatabase()
    {
        if (string.IsNullOrEmpty(_config.DatabasePath))
        {
            Console.WriteLine("No config database path available.");
            return;
        }

        try
        {
            string configDbPath = _config.DatabasePath;
            bool dbExists = File.Exists(configDbPath);
            
            Console.WriteLine($"Setting up database at: {configDbPath}");
            
            if (dbExists)
            {
                Console.WriteLine("Database file already exists. Checking/creating tables...");
            }
            else
            {
                Console.WriteLine("Database file doesn't exist. Creating new database...");
                
                // Create directory if it doesn't exist
                CreateDatabaseDirectory(configDbPath);
            }
            
            // Create/initialize database and tables
            CreateDatabaseWithSchema(configDbPath);
            
            try
            {
                _onDatabasePathChanged(configDbPath);
                
                if (dbExists)
                {
                    Console.WriteLine("Database connected successfully!");
                    Console.WriteLine("Tables verified/created.");
                }
                else
                {
                    Console.WriteLine("Database created successfully!");
                    Console.WriteLine("Tables initialized.");
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Database setup failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error setting up config database: {ex.Message}");
        }
    }

    /// <summary>
    /// Create a new database in the current folder
    /// </summary>
    public void SetupNewDatabase()
    {
        Console.Write("Enter database file name (will be created in current folder): ");
        string dbName = Console.ReadLine()?.Trim() ?? "";

        if (string.IsNullOrEmpty(dbName))
        {
            Logger.LogWarning("Database name cannot be empty.");
            return;
        }

        // Ensure .db extension
        if (!dbName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
        {
            dbName += ".db";
        }

        try
        {
            string newDbPath = Path.Combine(Directory.GetCurrentDirectory(), dbName);
            
            if (File.Exists(newDbPath))
            {
                Console.Write($"Database file '{dbName}' already exists. Overwrite? (y/n): ");
                if (Console.ReadLine()?.Trim().ToLower() != "y")
                {
                    Console.WriteLine("Database setup cancelled.");
                    return;
                }
            }
            
            // Create database and initialize tables
            CreateDatabaseWithSchema(newDbPath);
            
            try
            {
                _onDatabasePathChanged(newDbPath);
                Logger.LogInfo($"Database created successfully: {newDbPath}");
                Logger.LogInfo("Tables initialized.");
                
                // Ask if user wants to update config
                OfferConfigurationUpdate(newDbPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Database creation failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error creating database: {ex.Message}");
        }
    }


    private void CreateDatabaseDirectory(string databasePath)
    {
        string? directory = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
            Console.WriteLine($"Created directory: {directory}");
        }
    }

    /// <summary>
    /// Create and initialize a new database with optimized settings and full schema
    /// </summary>
    /// <param name="databasePath">Path where the database should be created</param>
    /// <returns>DatabaseInitializationResult with success status and details</returns>
    public static void CreateDatabaseWithSchema(string databasePath)
    {
        try
        {
            string connectionString = $"Data Source={databasePath};Version=3;";

            using var connection = new System.Data.SQLite.SQLiteConnection(connectionString);
            connection.Open();

            // Create all tables and indexes
            CreateDatabaseSchema(connection);
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error creating database: {ex.Message}");
            throw;
        }
    }


    /// <summary>
    /// Create the complete database schema including all tables and indexes
    /// </summary>
    private static void CreateDatabaseSchema(System.Data.SQLite.SQLiteConnection connection)
    {
        var schemaCommands = new[]
        {
            // Target files table (must be created first as other tables may reference it)
            @"CREATE TABLE IF NOT EXISTS target_files (
                TargetID INTEGER PRIMARY KEY AUTOINCREMENT,
                EngID TEXT,
                file_name TEXT,
                file_path TEXT UNIQUE,
                DrawID TEXT,
                notes TEXT,
                source_directory TEXT,
                folder_name TEXT,
                file_count INTEGER DEFAULT 1,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP
            )",

            // SolidWorks documents metadata table
            @"CREATE TABLE IF NOT EXISTS sw_documents (
                DocID INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL UNIQUE,
                file_name TEXT NOT NULL,
                document_type TEXT NOT NULL,
                file_size INTEGER,
                last_modified DATETIME,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP
            )",

            // Custom properties table
            @"CREATE TABLE IF NOT EXISTS sw_custom_properties (
                property_id INTEGER PRIMARY KEY AUTOINCREMENT,
                DocID INTEGER NOT NULL,
                property_name TEXT NOT NULL,
                property_value TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (DocID) REFERENCES sw_documents (DocID) ON DELETE CASCADE,
                UNIQUE(DocID, property_name)
            )",

            // BOM items table for assembly components
            @"CREATE TABLE IF NOT EXISTS sw_bom_items (
                bom_item_id INTEGER PRIMARY KEY AUTOINCREMENT,
                Parent_DocID INTEGER NOT NULL,
                component_name TEXT NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                configuration TEXT,
                quantity INTEGER DEFAULT 1,
                level INTEGER DEFAULT 0,
                is_suppressed BOOLEAN DEFAULT 0,
                suppression_state TEXT,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (Parent_DocID) REFERENCES sw_documents (DocID) ON DELETE CASCADE
            )",

            // Configurations table
            @"CREATE TABLE IF NOT EXISTS sw_configurations (
                configuration_id INTEGER PRIMARY KEY AUTOINCREMENT,
                DocID INTEGER NOT NULL,
                configuration_name TEXT NOT NULL,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (DocID) REFERENCES sw_documents (DocID) ON DELETE CASCADE,
                UNIQUE(DocID, configuration_name)
            )",

            // Materials table
            @"CREATE TABLE IF NOT EXISTS sw_materials (
                DocID INTEGER PRIMARY KEY,
                material_name TEXT NOT NULL,
                created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (DocID) REFERENCES sw_documents (DocID) ON DELETE CASCADE
            )",

            // Performance indexes
            @"CREATE INDEX IF NOT EXISTS idx_target_files_path ON target_files(file_path)",
            @"CREATE INDEX IF NOT EXISTS idx_documents_filepath ON sw_documents(file_path)",
            @"CREATE INDEX IF NOT EXISTS idx_documents_type ON sw_documents(document_type)",
            @"CREATE INDEX IF NOT EXISTS idx_custom_props_docid ON sw_custom_properties(DocID)",
            @"CREATE INDEX IF NOT EXISTS idx_custom_props_name ON sw_custom_properties(property_name)",
            @"CREATE INDEX IF NOT EXISTS idx_bom_parent_docid ON sw_bom_items(Parent_DocID)",
            @"CREATE INDEX IF NOT EXISTS idx_bom_level ON sw_bom_items(level)",
            @"CREATE INDEX IF NOT EXISTS idx_configurations_docid ON sw_configurations(DocID)",
            @"CREATE INDEX IF NOT EXISTS idx_materials_docid ON sw_materials(DocID)"
        };

        foreach (string commandText in schemaCommands)
        {
            using var command = new System.Data.SQLite.SQLiteCommand(commandText, connection);
            command.ExecuteNonQuery();
        }
    }


    private void OfferConfigurationUpdate(string newDbPath)
    {
        Console.Write("Update configuration to use this new database? (y/n): ");
        if (Console.ReadLine()?.Trim().ToLower() == "y")
        {
            _config.DatabasePath = newDbPath;
            Console.WriteLine("Configuration updated. Use 'Configuration settings > Save current configuration' to persist changes.");
        }
    }
}
