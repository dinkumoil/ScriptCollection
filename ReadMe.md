# What is this repository for?

This repository is a collection of (maybe) useful scripts I wrote over the years. It will be extended from time to time.


# The scripts at a glance

## CSV viewer

The purpose of this script is to be included in the Windows Explorer context menu for CSV files. It loads CSV files with MS Excel, formats all columns as text and fits the column widths to their content. Greatly simplifies viewing of CSV files.


## Cut

This is a small script that provides the functionality of the UNIX `cut` tool with plain Windows Batch script.

The script accepts the following command line switches:

* /b x  -  Set the starting line (default: the file's first line)
* /e x  -  Set the terminating line (default: the file's last line)
* /n    -  Show line numbers (default: off)


## EventConsumer

These scripts demonstrate the usage of WMI permanent event consumers. The script `NewProcessCreationEventMonitorInstaller.vbs` installs such an event consumer that monitors the system for the execution of a new _cmd.exe_ process. When this happens the handler script `NewProcessCreationEventHandler.vbs` is executed.

The noticeable benefit of WMI **permanent** event consumers is that they have to be installed only once, after that they are integrated into the system permanently.


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

The script accepts the following commandline parameters:

* /N:"Path"  -  Set path of the installation directory of Notepad++ (default: `C:\Program Files (x86)\Notepad++`)
* /G:"Path"  -  Set full path to _Gup.exe_ (default: `C:\Program Files (x86)\Notepad++\updater\GUP.exe`)
* /L:"Path"  -  Set full path to DLL file with plugin list (default: `C:\Program Files (x86)\Notepad++\plugins\config\nppPluginList.dll`)
* /J:"Path"  -  Set full path to JSON file with plugin list (default: `%TEMP%\nppPluginList.json`)
* /P:"Path"  -  Set path of the directory _Gup.exe_ should unpack the plugins to (default: `<Script-path>\Plugins`)

The directory where _Gup.exe_ unpacks the plugin packages will be created automatically.

The script outputs status messages during its work. To prevent that these messages get displayed with message boxes (which have to be closed one by one) it should be started with the following command line:

`cscript /nologo [Path-to-Script]LoadAllNppPlugins.vbs [Arguments]`

If the console Windows Script Host is set as default the script can be started with a double-click.


## SetSQLServerFirewallRules

The main script in this folder (`SetSqlFwRules.vbs`) sets the required rules in the Windows firewall to make MS SQL Server reachable for networked clients.

The script sets rules for all SQL Server instances it finds on the system.

It differentiates between main instances and named instances of SQL Server (only the latter ones need the port for SQL Server Browser to be opened).

It also differentiates between machines being domain members and machines in a peer-to-peer network (only the latter ones need the port for NetBios name service to be opened).

The script adds the rules to all firewall profiles but tries to avoid the _Public_ profile. It only adds rules to the _Public_ profile if it is the only available one.


## WMIPing

This script imitates the _PING_ command using a WMI class. It provides detailed error messages and sets a return code according to success or failure of the command. 