using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.SqlServer.Management.UserSettings;


internal class Program
{
  
  // Constants for working mode
  private const int MODE_DELETE = 0;
  private const int MODE_READ = 1;
  
  // Constants for error messages
  private const int ERR_OK = 0;
  private const int ERR_SHOW_HELP = 1;
  private const int ERR_NO_SERVERNAME = 2;
  private const int ERR_INVALID_PATH_CHARS = 3;
  private const int ERR_NO_PATH = 4;
  private const int ERR_FILE_NOT_FOUND = 5;
  private const int ERR_DESERIALIZATION = 6;
  private const int ERR_SERIALIZATION = 7;
  
  // Default values of program parameters
  private static string ExeName = "<Exe-Name>";
  private static string ServerToProcess;
  private static string SettingsFilePath = ".";
  private static string SettingsFileName = "SqlStudio.bin";
  private static string SettingsFile;
  private static int Mode = MODE_READ;
  
 
  // Main program
  static void Main(string[] args)
  {
    // Parse command line and set variables
    int chkStartupConditions = ParseCommandLine();
    
    // Break on any error
    if (chkStartupConditions != ERR_OK)
    {
      ShowMessage(chkStartupConditions);
      return;
    }

    // Assemble path to SSMS settings file
    SettingsFile = Path.Combine(SettingsFilePath, SettingsFileName);

    // Break if that file doesn't exist
    if (!File.Exists(SettingsFile))
    {
      ShowMessage(ERR_FILE_NOT_FOUND);
      return;
    }
    
    // AT FIRST install event handler for assembly resolving ...
    AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
    
    // ...and AFTER that call the function containing the code that uses these assemblies
    DoJob();
  }


  // This is the actual worker function of the program. It deletes entries of
  // servers in the settings file of SSMS.
  private static void DoJob()
  {
    // Counter for number of deleted server entries
    int delCntr = 0;
      
    // The settings file is a serialized instance of class SqlStudio.
    // Here it gets deserialized.
    SqlStudio settings;
    BinaryFormatter binaryFormatter = new BinaryFormatter();
    FileStream inStream = new FileStream(SettingsFile, FileMode.Open);

    try     {settings = (SqlStudio) binaryFormatter.Deserialize(inStream);}
    catch   {ShowMessage(ERR_DESERIALIZATION); return;}
    finally {inStream.Close();}
    
    // List of server entries to delete
    List<ServerConnectionItem> toProcess = new List<ServerConnectionItem>();

    // Iterate over server list of server types in settings file
    // The list elements are key-value pairs
    foreach (var pair in settings.SSMS.ConnectionOptions.ServerTypes)
    {
      // Retrieve value of current list element
      ServerTypeItem serverType = pair.Value;

      // Iterate over list of servers of current type. Add all servers that fits
      // the search criterium to a list.
      foreach (ServerConnectionItem server in serverType.Servers)
      {
        if (ServerToProcess.Equals("*"))
          toProcess.Add(server);

        else if (server.Instance.StartsWith(ServerToProcess, StringComparison.InvariantCultureIgnoreCase))
          toProcess.Add(server);
      }

      // Process the list filled above
      foreach (ServerConnectionItem server in toProcess)
      {
        switch (Mode)
        {
          // In delete mode delete the server entry from the settings class instance
          case MODE_DELETE:
            serverType.Servers.RemoveItem(server);
            delCntr++;
            break;
            
          // In read mode output the server entry to console window
          case MODE_READ:
            Console.WriteLine(server.Instance);
            break;
        }
      }
        
      // Clear intermediate list for next iteration of server types
      toProcess.Clear();
    }

    // In delete mode the changed settings have to be written to the settings file
    if (Mode == MODE_DELETE)
    {
      // Serialize the settings class instance to a file stream
      FileStream outStream = new FileStream(SettingsFile, FileMode.Create);

      try     {binaryFormatter.Serialize(outStream, settings);}
      catch   {ShowMessage(ERR_SERIALIZATION); return;}
      finally {outStream.Close();}

      // Finally show a status message
      Console.WriteLine("Deleted entries: {0}", delCntr);
    }
  }
  
  
  // This function parses the command line, retrieves the name the program has
  // been started as, checks the provided arguments to be formally correct, and
  // stores them into variables
  private static int ParseCommandLine()
  {
    try
    {
      string[] arguments = Environment.GetCommandLineArgs();
      
      for(int cnt = 0; cnt < arguments.Length; cnt++)
      {
        // The first argument contains the name the program has been started as
        if (cnt == 0)
        {
          // If there is anything we take it
          if (!arguments[cnt].Equals(String.Empty))
            ExeName = Path.GetFileName(arguments[cnt]);

          // If there are no more arguments return to caller and show the help message
          if (arguments.Length < 2) return ERR_SHOW_HELP;
        }
        
        // Even if there are other arguments, if the user requested help we will show it
        else if (arguments[cnt].Equals("/?") ||
                 arguments[cnt].Equals("-?"))
        {
          return ERR_SHOW_HELP;
        }
            
        // The current argument is the path to the folder of the SSMS settings file
        else if (arguments[cnt].StartsWith("/p:", StringComparison.InvariantCultureIgnoreCase) ||
                 arguments[cnt].StartsWith("-p:", StringComparison.InvariantCultureIgnoreCase))
        {
          // Remove double quotes from path
          SettingsFilePath = arguments[cnt].Substring(3).Trim(new char[] {'"'});
          SettingsFilePath = Path.GetFullPath(SettingsFilePath);
        }
        
        // The current argument is the complete path to the SSMS settings file
        else if (arguments[cnt].StartsWith("/f:", StringComparison.InvariantCultureIgnoreCase) ||
                 arguments[cnt].StartsWith("-f:", StringComparison.InvariantCultureIgnoreCase))
        {
          // Remove double quotes from path
          SettingsFilePath = arguments[cnt].Substring(3).Trim(new char[] {'"'});
          SettingsFilePath = Path.GetFullPath(SettingsFilePath);
          
          // Split complete path into folder path and file name
          SettingsFileName = Path.GetFileName(SettingsFilePath);
          SettingsFilePath = Path.GetDirectoryName(SettingsFilePath);
        }
        
        // The current argument is the program's working mode
        // d => delete
        else if (arguments[cnt].Equals("/d", StringComparison.InvariantCultureIgnoreCase) ||
                 arguments[cnt].Equals("-d", StringComparison.InvariantCultureIgnoreCase))
        {
          Mode = MODE_DELETE;
        }
        
        // r => read
        else if (arguments[cnt].Equals("/r", StringComparison.InvariantCultureIgnoreCase) ||
                 arguments[cnt].Equals("-r", StringComparison.InvariantCultureIgnoreCase))
        {
          Mode = MODE_READ;
        }
        
        // The current arguments defines to process all server entries
        else if (arguments[cnt].Equals("/all", StringComparison.InvariantCultureIgnoreCase) ||
                 arguments[cnt].Equals("-all", StringComparison.InvariantCultureIgnoreCase))
        {
          ServerToProcess = "*";
        }
        
        // Everything that doesn't fit the patterns above is considered to be a
        // server name
        else
        {
          ServerToProcess = arguments[cnt];
        }
      }
      
      // Final check: the string with the server name must not be empty
      if (String.IsNullOrEmpty(ServerToProcess))
        return ERR_NO_SERVERNAME;
      else
        return ERR_OK;
    }
    
    catch
    {
      // We reach this point if the path to the folder of the SSMS settings file
      // or the complete path to this file was not provided or the path contains
      // invalid characters
      if (String.IsNullOrEmpty(SettingsFilePath) ||
          String.IsNullOrEmpty(SettingsFileName))
        return ERR_NO_PATH;
      else
        return ERR_INVALID_PATH_CHARS;
    }
  }
  
  
  // All error messages and the help message are displayed by this function.
  // For error messages the standard error device is used.
  private static void ShowMessage(int errNum)
  {
    if (errNum != ERR_SHOW_HELP)
      Console.SetOut(Console.Error);
    
    Console.WriteLine();

    switch (errNum)
    {
      case ERR_SERIALIZATION:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Failed to write settings file.");
        goto default;
      
      case ERR_DESERIALIZATION:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Failed to read settings file.");
        goto default;
        
      case ERR_FILE_NOT_FOUND:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Settings file not found.");
        goto default;
      
      case ERR_NO_PATH:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Please provide the path respectively the name of the settings file.");
        goto default;

      case ERR_INVALID_PATH_CHARS:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Please provide only valid characters for the path respectively the name of the");
        Console.WriteLine(@"settings file.");
        goto default;

      case ERR_NO_SERVERNAME:
        Console.WriteLine(@"Error!");
        Console.WriteLine(@"Please provide the name of the SQL Server whose entries you want to delete from");
        Console.WriteLine(@"settings file. You can also provide the switch /all to delete the entries of all");
        Console.WriteLine(@"servers.");
        goto default;
      
      case ERR_SHOW_HELP: 
        Console.WriteLine(@"Read/delete all entries of the provided SQL Server or of all SQL Servers from");
        Console.WriteLine(@"the settings file of SQL Server Management Studio 2014.");
        goto default;
      
      default:
        Console.WriteLine();
        Console.WriteLine(@"Usage: {0} <Server>|/all [/r|/d] [/p:<Path>|/f:<Path\FileName>]", ExeName);
        Console.WriteLine();
        Console.WriteLine(@"  Server     - Entries in the settings file starting with this string will be");
        Console.WriteLine(@"               processed");
        Console.WriteLine(@"  /all         All entries in the settings file will be processed");
        Console.WriteLine(@"  /r         - Read entries (default)");
        Console.WriteLine(@"  /d         - Delete entries");
        Console.WriteLine(@"  Path       - Path to the directory where the settings file SqlStudio.bin is");
        Console.WriteLine(@"               stored");
        Console.WriteLine(@"  FileName   - Name of settings file");
        Console.WriteLine();
        Console.WriteLine(@"If the path or the file name contains spaces it has to be enclosed in double");
        Console.WriteLine(@"quotes.");
        break;
    }
  }
  
  
  // This is an event handler that will be called when an assebly is about to be
  // loaded. It's the only way to load the assembly SqlWorkbench.Interfaces. The
  // assembly Microsoft.SqlServer.Management.UserSettings is added as well so it
  // has not to be stored in the same directory like the program itself.
  private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
  {
    if (args.Name.StartsWith("SqlWorkbench.Interfaces", StringComparison.InvariantCultureIgnoreCase))
    {
      string assemblyPath = String.Concat(GetProgramFilesDir(), @"\Microsoft SQL Server\120\Tools\Binn\ManagementStudio\SqlWorkbench.Interfaces.dll");
      return Assembly.LoadFrom(assemblyPath);
    }
    
    else if (args.Name.StartsWith("Microsoft.SqlServer.Management.UserSettings", StringComparison.InvariantCultureIgnoreCase))
    {
      string assemblyPath = String.Concat(GetProgramFilesDir(), @"\Microsoft SQL Server\120\Tools\Binn\ManagementStudio\Microsoft.SqlServer.Management.UserSettings.dll");
      return Assembly.LoadFrom(assemblyPath);
    }
    
    else
    {
      return Assembly.Load(args.Name);
    }
  }
  
  
  // Retrieve the path to the program files directory for 32 bit programs
  private static string GetProgramFilesDir()
  {
    string programFilesDir = Environment.GetEnvironmentVariable("ProgramFiles(x86)");
    
    if (String.IsNullOrEmpty(programFilesDir))
      programFilesDir = Environment.GetEnvironmentVariable("ProgramFiles");

    return programFilesDir;
  }
}
