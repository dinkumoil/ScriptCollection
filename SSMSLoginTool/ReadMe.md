# SSMSLoginTool

This is a build script to compile a small program written in C# using the C# compiler provided with every .NET installation.

SQL Server Management Studio is a great tool but it has an annoying bug in the login box. Under unknown circumstances it looses sometimes the saved password for an already known SQL Server. Everytime you try to log in, the password field for this server is blank and even if you provide it again and set the option "Remember password" it will not be stored.

It is possible to overcome this problem by deleting the buggy entries from the settings file of SQL Server Management Studio. But this file contains a serialized .NET class which is not human readable nor editable with a simple text editor. A .NET assembly containing the SSMS settings file class is required to do that.

Fortunately I found a piece of code on StackOverflow that solved the problem. I was able to extend it to a rather comfortably to use console program. The program can be used to read or delete all server entries in the settings file or only the entry of a certain server.

The program in its current version is only suitable for SQL Server 2014. But that is only because of some paths that have to be changed for other SQL Server versions. In the past I used the program in conjunction with SQL Server 2008 R2 as well. The following changes have to be made to adapt it to another version of SQL Server:

* In the file _SSMSLoginTool.cs_ the method _CurrentDomain_AssemblyResolve_ contains the paths to two assemblies the program uses. These files come together with SQL Server, thus they are located in its installation directory. Please adapt these paths to your version of SQL Server.
* The file _SSMSLoginTool.proj_ contains in the XML node _\\\\ItemGroup\ReferencedAssemblies_ the path to one of the assemblies mentioned above, too. It has to be adapted to your version of SQL Server as well.

To build the program start the script _Build.cmd_. It searches the newest available version of _MSBuild_, sets _ToolsVersion_ and _TargetFramework_ to the highest available on your machine and thus compiles the program with the C# compiler included in the newest available .NET installation for the .NET framework with the highest version number available.

For instructions on how to use the program start it without any parameters or with the parameter _/?_ or _-?_.
