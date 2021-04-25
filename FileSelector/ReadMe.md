# FileSelector

This script provides a file selector written in native Windows Batch script. You can browse your hard disk(s) in a console window and select a file or folder. It is also possible to create and delete files and folders by means of a _Tools_ menu.

A directory is presented as a scrollable list of entries, each preceded with a number.

* To **navigate** to a **directory** input its number and press _ENTER_.
* To **select** a **file** input its number and press _ENTER_. The _FileSelector_ script returns to its caller.
* To **select** a **directory** input character _C_ + its number and press _ENTER_. The _FileSelector_ script returns to its caller.

It can be choosen to return the selected item via a variable or a file. The file method is more reliable because this way it is more easy to return files/folders whose paths contain characters which have a special meaning in Batch script (e.g. as operators).

A demo script that shows how to use the _FileSelector_ is included.
