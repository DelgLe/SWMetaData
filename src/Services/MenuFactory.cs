using System;
using System.Collections.Generic;

/// <summary>
/// Factory class for creating and managing user menu options
/// </summary>
public class MenuFactory
{
    private readonly List<MenuOption> _options;
    private readonly string _title;

    public MenuFactory(string title)
    {
        _title = title;
        _options = new List<MenuOption>();
    }

    /// <summary>
    /// Add a menu option with an action
    /// </summary>
    public MenuFactory AddOption(string choice, string description, Action action)
    {
        _options.Add(new MenuOption(choice, description, action));
        return this;
    }

    /// <summary>
    /// Add a menu option with a function that returns a boolean (for exit/back conditions)
    /// </summary>
    public MenuFactory AddOption(string choice, string description, Func<bool> action)
    {
        _options.Add(new MenuOption(choice, description, action));
        return this;
    }

    /// <summary>
    /// Add a separator/back option that exits the menu
    /// </summary>
    public MenuFactory AddBackOption(string choice = "0", string description = "Back to previous menu")
    {
        _options.Add(new MenuOption(choice, description, () => false)); // Return false to exit loop
        return this;
    }

    /// <summary>
    /// Display the menu and handle user input in a loop
    /// </summary>
    public void RunMenu()
    {
        while (true)
        {
            ShowMenu();
            Console.Write("Enter your choice: ");
            string input = Console.ReadLine()?.Trim() ?? "";

            var selectedOption = _options.Find(opt => opt.Choice.Equals(input, StringComparison.OrdinalIgnoreCase));
            
            if (selectedOption != null)
            {
                try
                {
                    // Execute the action and check if we should continue the loop
                    bool continueLoop = selectedOption.Execute();
                    if (!continueLoop)
                    {
                        break; // Exit the menu loop
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error executing option: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Invalid option. Please try again.");
            }
        }
    }

    /// <summary>
    /// Display the menu once and return the user's choice (no loop)
    /// </summary>
    public string ShowMenuAndGetChoice()
    {
        ShowMenu();
        Console.Write("Enter your choice: ");
        return Console.ReadLine()?.Trim() ?? "";
    }

    /// <summary>
    /// Execute a specific choice programmatically
    /// </summary>
    public bool ExecuteChoice(string choice)
    {
        var selectedOption = _options.Find(opt => opt.Choice.Equals(choice, StringComparison.OrdinalIgnoreCase));
        if (selectedOption != null)
        {
            return selectedOption.Execute();
        }
        return true; // Continue if choice not found
    }

    private void ShowMenu()
    {
        Console.WriteLine($"\n=== {_title} ===");
        foreach (var option in _options)
        {
            Console.WriteLine($"{option.Choice}. {option.Description}");
        }
    }
}

/// <summary>
/// Represents a single menu option
/// </summary>
public class MenuOption
{
    public string Choice { get; }
    public string Description { get; }
    private readonly Func<bool> _action;

    public MenuOption(string choice, string description, Action action)
    {
        Choice = choice;
        Description = description;
        _action = () => 
        {
            action.Invoke();
            return true; // Continue menu loop by default
        };
    }

    public MenuOption(string choice, string description, Func<bool> action)
    {
        Choice = choice;
        Description = description;
        _action = action;
    }

    public bool Execute()
    {
        return _action.Invoke();
    }
}

/// <summary>
/// Static factory methods for common menu patterns
/// </summary>
public static class MenuFactoryExtensions
{
    /// <summary>
    /// Create a standard menu with numbered options and a back/exit option
    /// </summary>
    public static MenuFactory CreateStandardMenu(string title)
    {
        return new MenuFactory(title);
    }

    /// <summary>
    /// Create a yes/no confirmation menu
    /// </summary>
    public static bool CreateConfirmationMenu(string message, string defaultChoice = "n")
    {
        Console.Write($"{message} (y/n) [{defaultChoice}]: ");
        string input = Console.ReadLine()?.Trim().ToLower() ?? defaultChoice;
        
        if (string.IsNullOrEmpty(input))
            input = defaultChoice;
            
        return input == "y" || input == "yes";
    }
}