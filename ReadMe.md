# What is this repository for?

This repository is a collection of (maybe) useful scripts I wrote over the years. It will be extended from time to time.


# The scripts at a glance

## CSV viewer

Loads CSV files with MS Excel, formats all columns as text and fits the column widths to their content. Greatly simplifies viewing of CSV files.


## CombineFiles

A small script that is able to merge the content of two files line by line into an output file.


## Cut

A small script that provides the functionality of the UNIX `cut` tool with plain Windows Batch script.


## EventConsumer

These scripts demonstrate the usage of WMI permanent event consumers from within Windows batch scripts and VBS scripts.


## ExplorerContextMenu

This directory contains three Windows Explorer context menu extensions:

* `DeleteADS`  -  Delete NTFS alternate data streams.
* `OpenConsoleHere`  -  Open a normal console or a console with admin rights in the current directory/the directory of the current file.
* `UpdateExplorerIcons`  -  Force Windows Explorer to reload the data of all desktop and file list icons.


## FileSelector

This script provides a file selector written in native Windows Batch script. You can browse your hard disk(s) in a console window and select a file or folder. It is also possible to create and delete files and folders by means of a _Tools_ menu. A demo script that shows how to use the _FileSelector_ is included.


## GetDayOfWeek

This script determines the day of the week of a specific date. The date can be provided via command line or stdin. The day of the week can be displayed in numerical form, as german or english abbreviation or with its full german or english name.


## GetSysTimes

This folder contains two Batch scripts to

* retrieve the last system boot up time.
* retrieve the system uptime.


## HashFile

With this script it is possible to calculate hash values of files using only the native Windows executable _CertUtil.exe_.


## HeaderFiles

This is a collection of VBScript classes, utility functions and OS constants. They can be included into own scripts to avoid writing the same code over and over again.


## Lua Scripts

This is a collection of scripts for the _LuaScript_ plugin for Notepad++ to extend Npp's functionality.


## MigrateNppPlugins

With this script it is possible to migrate all Notepad++ plugins to the new plugin directory structure required by Notepad++ v7.6.3 and above.


## NppExec Development

The files in this directory provide support for syntax highlighting and a function list parser for code of _NppExec_ Notepad++ plugin.


## NppExec Scripts

This is a collection of scripts for the _NppExec_ plugin for Notepad++ to extend Npp's functionality.


## NppPluginManagement

These scripts allow to manage plugins for Notepad++. The benefit over using Notepad++'s build in _Plugins Admin_ is that you can always install/update to the plugin versions from the most recent plugins list and that also older versions of Notepad++ prior to v7.6 are supported.


## Pascal Development

The files in this directory provide function list parsers for Notepad++. There is a parser for source code files of Pascal and derived languages like _Delphi_ and _Free Pascal_ and a parser for _Delphi_ forms (i.e. `*.dfm`) files.


## SetSQLServerFirewallRules

Script to set the required rules in the Windows firewall to make MS SQL Server reachable for network clients.


## SSMSLoginTool

This is a build script to compile a small program written in C#. It helps solving the problem that SQL Server Management Studio looses sometimes the saved password for an already known SQL Server.


## WMIPing

This script imitates the _PING_ command using a WMI class. In opposite to the _PING_ command it provides detailed error messages and sets a return code according to success or failure of the command.
