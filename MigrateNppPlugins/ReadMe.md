# MigrateNppPlugins

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
