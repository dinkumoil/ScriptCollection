
================================================================================
Delete entries from history of login box in SQL Server Management Studio (SSMS)
The current version is suitable for SQL Server 2014 (64 bit)
================================================================================


Links:

http://stackoverflow.com/questions/6230159/how-to-delete-server-entries-in-sql-server-management-studios-connect-to-serve/6374534#6374534
http://stackoverflow.com/questions/1413831/c-sharp-add-a-reference-using-only-code-no-ide-add-reference-functions



Path of SSMS settings file (SqlStudio.bin):

C:\Users\<user-name>\AppData\Roaming\Microsoft\SQL Server Management Studio\12.0



Command line for compiling without project file:

"C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" /nologo /target:exe /out:"SSMSLogin.exe" "SSMSLogin.cs" /reference:"C:\Program Files\Microsoft SQL Server\120\Tools\Binn\ManagementStudio\Microsoft.SqlServer.Management.UserSettings.dll"
