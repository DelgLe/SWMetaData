# SolidWorks Metadata Reader - Configuration Guide

## JSON Configuration Support

The application now supports JSON-based configuration files to specify database paths and processing settings.

### Configuration File Location

By default, the application looks for `swmetadata-config.json` in the same directory as the executable.

### Configuration Options

```json
{
  "databasePath": "C:\\path\\to\\your\\database.db",
  "autoCreateDatabase": true,
  "defaultTargetFilesPath": "C:\\path\\to\\target_files.csv",
  "processing": {
    "processBomForAssemblies": true,
    "includeCustomProperties": true,
    "validateFilesExist": true,
    "batchProcessingTimeout": 60000
  }
}
```

### Configuration Properties

| Property | Type | Description | Default |
|----------|------|-------------|---------|
| `databasePath` | string | Full path to SQLite database file | `./swmetadata.db` |
| `autoCreateDatabase` | boolean | Create database if it doesn't exist | `true` |
| `defaultTargetFilesPath` | string | Path to target files CSV/list | `null` |
| `processing.processBomForAssemblies` | boolean | Generate BOM for assembly files | `true` |
| `processing.includeCustomProperties` | boolean | Extract custom properties | `true` |
| `processing.validateFilesExist` | boolean | Check if files exist before processing | `true` |
| `processing.batchProcessingTimeout` | number | Timeout in milliseconds for batch operations | `30000` |

### Using Configuration

#### Method 1: Automatic Loading
1. Create `swmetadata-config.json` in the application directory
2. Run the application - it will automatically load the configuration
3. The specified database path will be used automatically

#### Method 2: Interactive Setup
1. Run the application
2. Choose option **6. Configuration settings**
3. Use the configuration management menu to:
   - View current configuration
   - Set database path interactively
   - Create example config files
   - Reload configuration from custom paths
   - Save current settings

### Example Configuration Files

#### Basic Configuration
```json
{
  "databasePath": "C:\\SolidWorksData\\metadata.db",
  "autoCreateDatabase": true
}
```

#### Advanced Configuration
```json
{
  "databasePath": "C:\\Projects\\SolidWorks\\database\\metadata.db",
  "autoCreateDatabase": true,
  "defaultTargetFilesPath": "C:\\Projects\\SolidWorks\\target_files.txt",
  "processing": {
    "processBomForAssemblies": true,
    "includeCustomProperties": true,
    "validateFilesExist": true,
    "batchProcessingTimeout": 120000
  }
}
```

### Working with Existing Databases

If you have an existing SQLite database with the `target_files` table:

1. Create a configuration file pointing to your existing database:
```json
{
  "databasePath": "C:\\path\\to\\existing\\database.db",
  "autoCreateDatabase": false
}
```

2. The application will connect to your existing database and use the existing `target_files` table

3. Use **Option 4: Manage target files** to view and manage your existing target files

4. Use **Option 5: Process all target files** to batch process all files in your target table

### Tips

- Use forward slashes `/` or double backslashes `\\\\` in JSON file paths
- The application will create missing directories if `autoCreateDatabase` is true
- Configuration is loaded once at startup - use **Option 6** to reload changes
- The database path can be absolute or relative to the application directory