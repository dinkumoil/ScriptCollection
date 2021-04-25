# SetSQLServerFirewallRules

The main script in this folder (`SetSqlFwRules.vbs`) sets the required rules in the Windows firewall to make MS SQL Server reachable for network clients.

The script sets rules for all SQL Server instances it finds on the system.

It differentiates between main instances and named instances of SQL Server (only the latter ones need the port for SQL Server Browser to be opened).

It also differentiates between machines being domain members and machines in a peer-to-peer network (only the latter ones need the port for NetBios name service to be opened).

The script adds the rules to all firewall profiles but tries to avoid the _Public_ profile. It only adds rules to the _Public_ profile if it is the only available one.
