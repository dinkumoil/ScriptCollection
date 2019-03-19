# What is this repository for?

This repository is a collection of (maybe) useful scripts I wrote over the years. It will be extended from time to time.


# The scripts at a glance

## CSV viewer

The purpose of this script is to be included in the Windows Explorer context menu for CSV files. It loads CSV files with MS Excel, formats all columns as text and fits the column widths to their content. Greatly simplifies viewing of CSV files.


## Cut

This is a small script that provides the functionality of the UNIX `cut` tool with plain Windows Batch script.

The script accepts the following command line switches:

* `/b x`  -  Set the starting line (default: the file's first line)
* `/e x`  -  Set the terminating line (default: the file's last line)
* `/n`    -  Show line numbers (default: off)


## EventConsumer

These scripts demonstrate the usage of WMI permanent event consumers. The script `NewProcessCreationEventMonitorInstaller.vbs` installs such an event consumer that monitors the system for the execution of a new _cmd.exe_ process. When this happens the handler script `NewProcessCreationEventHandler.vbs` is executed.

The noticeable benefit of WMI **permanent** event consumers is that they have to be installed only once, after that they are integrated into the system permanently.


## ExplorerContextMenu

This directory contains three Windows Explorer context menu extensions:

* `DeleteADS`  -  Delete NTFS alternate data streams. Can be used to delete the Zone Identifier of files downloaded from the internet to avoid annoying warnings when opening them.
* `OpenConsoleHere`  -  Open a normal console or a console with admin rights in the current directory/the directory of the current file.
* `UpdateExplorerIcons`  -  Force Windows Explorer to reload the data of all desktop and file list icons.

To install the context menu extensions double-click the _install.cmd_ files in the respective directory.

For _DeleteADS_ you will need to download the _Streams.exe_ tool by SysInternals. For _UpdateExplorerIcons_ you will need to download the _WinApiExec.exe_ tool and the _HStart.exe_ tool. See the equally named _*.txt_ files for download links and installation instructions.


## FileSelector

This script provides a file selector written in native Windows Batch script. You can browse your hard disk(s) in a console window and select a file or folder. It is also possible to create and delete files and folders by means of a _Tools_ menu.

A directory is presented as a scrollable list of entries, each preceded with a number.

* To **navigate** to a **directory** input its number and press _ENTER_.
* To **select** a **file** input its number and press _ENTER_. The _FileSelector_ script returns to its caller.
* To **select** a **directory** input character _C_ + its number and press _ENTER_. The _FileSelector_ script returns to its caller.

It can be choosen to return the selected item via a variable or a file. The file method is more reliable because this way it is more easy to return files/folders whose paths contain characters which have a special meaning in Batch script (e.g. as operators).

A demo script that shows how to use the _FileSelector_ is included.


## GetSysTimes

This folder contains two Batch scripts with inline VBScript:

* Retrieving the last system boot up time.
* Retrieving the system uptime.


## HeaderFiles

This is a collection of VBScript classes, utility functions and OS constants.

* `ADO.vbs`  -  ADO constants
* `WMI.vbs`  -  WMI constants
* `Utils.vbs`  -  Utility functions (e.g. QuickSort)
* `ClassDatabaseFileDSN.vbs`  -  Working with databases using a file DSN
* `ClassDatabaseSqlServer.vbs`  -  Working with MS SQL Server databases using OLEDB provider
* `ClassFileNameMatch.vbs`  -  Pattern matching for file and path name wildcards
* `ClassFileVersionInfo.vbs`  -  Retrieve file version informations from EXE and DLL files
* `ClassIniFile.vbs`  -  Working with INI files (read, write, change)
* `ClassSimilarity.vbs`  -  Phonetical comparison (SoundEx, Kölner Phonetic, Levenshtein distance)


## LoadAllNppPlugins

This script is able to download and unzip all _Notepad++_ plugin packages that are listed in the official _Notepad++_ plugin list. This list is provided as a DLL file and part of every _Notepad++_ installation since version 7.6. For downloading and unzipping the plugins the script uses the same helper program like _Notepad++_ itself, _Gup.exe_ which is part of every _Notepad++_ installation as well.

The script can also be used on a system without an installed copy of _Notepad++_. In this case it needs at least _Gup_ (its ZIP file can be downloaded [here](https://github.com/notepad-plus-plus/wingup/releases), contains also all other files _Gup.exe_ needs) and the DLL file with the plugin list (its ZIP file can be downloaded [here](https://github.com/notepad-plus-plus/nppPluginList/releases)).

To extract the plugin list (a JSON document) from its DLL file, the script uses an external helper program named _WinApiExec_. With this tool it is possible to perform Win32 API calls from within a script. If this tool is not present in the intended directory the script will download and unpack its ZIP file.

The script accepts the following command line parameters:

* `/N:"Path"`  -  Set path of the installation directory of Notepad++ (default: `C:\Program Files (x86)\Notepad++`)
* `/G:"Path"`  -  Set full path to _Gup.exe_ (default: `C:\Program Files (x86)\Notepad++\updater\GUP.exe`)
* `/L:"Path"`  -  Set full path to DLL file with plugin list (default: `C:\Program Files (x86)\Notepad++\plugins\config\nppPluginList.dll`)
* `/J:"Path"`  -  Set full path to JSON file with plugin list (default: `%TEMP%\nppPluginList.json`)
* `/P:"Path"`  -  Set path of the directory _Gup.exe_ should unpack the plugins to (default: `<Script-path>\Plugins`)

The directory where _Gup.exe_ unpacks the plugin packages will be created automatically.

The options `/L` and `/J` are exclusive, the one occuring later is taken into account. Also the option `/N` influences the paths for _Gup.exe_ and the plugin list's DLL file. If you want to set non-standard paths for one or both of them **and** a non-standard path for Notepad++ provide the `/G` and `/L` options **after** the `/N` option.

The script outputs status messages during its work. To prevent that these messages get displayed with message boxes (which have to be closed one by one) it should be started with the following command line:

`cscript /nologo [Path-to-Script]LoadAllNppPlugins.vbs [Arguments]`

If the console Windows Script Host is set as default the script can be started with a double-click.


## MigrateNppPlugins

With this script it is possible to migrate all Notepad++ plugins to the new plugin directory structure required by Notepad++ v7.6.3 and above. It should be run **after** upgrading Notepad++ to that version. The script is able to migrate plugins of

* local installations of Notepad++ up to v7.6.2.
* portable installations of Notepad++ up to v7.6.2.
* hybrid installations of Notepad++ up to v7.5.9 (plugin DLL files in the user profile).

The script not only migrates the plugin DLL file itself to the new location but also companion files and folders, i.e. files required for the plugin to work properly as well as help and documentation files, if there are any. This works **only** if the files/folders are named **exactly** like the plugin's DLL file.

**Please note:** There are plugins out there that store their companion files under e.g. `<Notepad++-install-dir>\plugins\<plugin-name>` and when trying to load them they use a hard-coded path. That means they will not find these files anymore after the script has moved them to the new location. In this case Notepad++ respectively the plugin will show some kind of error message during start up or the plugin simply will not work as desired, e.g. showing its help file will fail. You should try then to move the companion files under suspicion back to their previous location.

The normal use case for the script is to be run in interactive mode. The script searches for local and hybrid installations of Notepad++ under `%ProgramFiles%` and `%ProgramFiles(x86)`. If it doesn't find any of them it asks for the path to a portable installation. But even if it finds a local or hybrid installation it asks if the user prefers to migrate the plugins of a portable installation. In case of a local or hybrid installation the script restarts itself and triggers an User Account Control (UAC) prompt to elevate the user rights it runs under. Then it starts to migrate the plugins.

If you run the script with the following command line you can use it in an automated way. **Please note:** If you want to migrate a local or hybrid installation you have to run the script with administrative user rights.

**`MigrateNppPlugins.cmd "Source Path" "Destination Path" "Installation Type"`**

* `Source Path`  -  Path to source plugin directory
* `Destination Path`  -  Path to destination plugin directory
* `Installation Type`  -  Can be one of `Local`, `Localv7.6`, `Localv7.6.1`, `Hybrid`, `Portable`


`Source Path` depends on the installation type and version number of the old Notepad++ installation:
</br>

| Version   | Local installation                 | Hybrid installation           | Portable installation         |
|----------:|:---------------------------------- |:-----------------------------:|:----------------------------- |
| <= v7.5.9 | `%ProgramFiles%\plugins`           | `%AppData%\Notepad++\plugins` | `<Npp-install-path>\plugins`  |
|    v7.6   | `%LocalAppData%\Notepad++\plugins` |             n/a               | `<Npp-install-path>\plugins`  |
|    v7.6.1 | `%ProgramData%\Notepad++\plugins`  |             n/a               | `<Npp-install-path>\plugins`  |
|    v7.6.2 | `%ProgramData%\Notepad++\plugins`  |             n/a               | `<Npp-install-path>\plugins`  |


`Destination Path` depends on the installation type of the new Notepad++ v7.6.3 (or above) installation:
</br>

| Local installation     | Portable installation      |
|:---------------------- |:-------------------------- |
|`%ProgramFiles%\plugins`|`<Npp-install-path>\plugins`|


`Installation Type` depends on the installation type and version number of the old Notepad++ installation:
</br>

| Version   | Local installation | Hybrid installation  | Portable installation  |
|----------:|:------------------ |:--------------------:|:---------------------- |
| <= v7.5.9 | `Local`            | `Hybrid`             | `Portable`             |
|    v7.6   | `Localv7.6`        |        n/a           | `Portable`             |
|    v7.6.1 | `Localv7.6.1`      |        n/a           | `Portable`             |
|    v7.6.2 | `Localv7.6.1`      |        n/a           | `Portable`             |


## SetSQLServerFirewallRules

The main script in this folder (`SetSqlFwRules.vbs`) sets the required rules in the Windows firewall to make MS SQL Server reachable for networked clients.

The script sets rules for all SQL Server instances it finds on the system.

It differentiates between main instances and named instances of SQL Server (only the latter ones need the port for SQL Server Browser to be opened).

It also differentiates between machines being domain members and machines in a peer-to-peer network (only the latter ones need the port for NetBios name service to be opened).

The script adds the rules to all firewall profiles but tries to avoid the _Public_ profile. It only adds rules to the _Public_ profile if it is the only available one.


## SSMSLoginTool

This is a build script to compile a small program written in C# using the C# compiler provided with every .NET installation.

SQL Server Management Studio is a great tool but it has an annoying bug in the login box. Under unknown circumstances it looses sometimes the saved password for an already known SQL Server. Everytime you try to log in, the password field for this server is blank and even if you provide it again and set the option "Remember password" it will not be stored.

It is possible to overcome this problem by deleting the buggy entries from the settings file of SQL Server Management Studio. But this file contains a serialized .NET class which is not human readable nor editable with a simple text editor. A .NET assembly containing the SSMS settings file class is required to do that.

Fortunately I found a piece of code on StackOverflow that solved the problem and was able to extend it to a rather comfortably to use console program. The program can be used to read or delete all server entries in the settings file or only the entry of a certain server.

The program in its current version is only suitable for SQL Server 2014. But that is only because of some paths that have to be changed for other SQL Server versions. In the past I used the program in conjunction with SQL Server 2008 R2 as well. The following changes have to be made to adapt it to another version of SQL Server:

* In the file _SSMSLoginTool.cs_ the method _CurrentDomain_AssemblyResolve_ contains the paths to two assemblies the program uses. These files come together with SQL Server, thus they are located in its installation directory. Please adapt these paths to your version of SQL Server.
* The file _SSMSLoginTool.proj_ contains in the XML node _\\\\ItemGroup\ReferencedAssemblies_ the path to one of the assemblies mentioned above, too. It has to be adapted to your version of SQL Server as well.

To build the program start the script _Build.cmd_. It searches the newest available version of _MSBuild_, sets the _TargetFramework_ to the highest available on your machine and compiles the program with the C# compiler included in the newest available .NET installation.

For instructions on how to use the program start it without any parameters or with the parameter _/?_ or _-?_.


## WMIPing

This script imitates the _PING_ command using a WMI class. It provides detailed error messages and sets a return code according to success or failure of the command. 