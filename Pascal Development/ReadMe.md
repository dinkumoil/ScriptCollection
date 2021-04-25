# Pascal Development

The file `pascal.xml` provides a function list parser for Notepad++. After installing this file you will be able to use the function list panel of Notepad++ to navigate through your code files.

**Please note:** The parser has some known issues and does not work always as expected. I will **not react** to bug reports but fixes are welcome. If you find some unwanted behaviour please fix it by yourself and file a pull request.


## Installation

The following steps depend on the version of Notepad++ you use.

**Notepad++ versions prior to v7.9.1**

1. Open file `pascal.xml` with Notepad++.
2. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
    **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
3. Open file `functionList.xml` from that folder with Notepad++.
4. From the file `pascal.xml`, copy the whole `/NotepadPlus/functionList/parser` XML node into the file `functionList.xml` as a child node of the `/NotepadPlus/functionList/parsers` XML node.
5. In file `functionList.xml` go to the XML node `/NotepadPlus/functionList/associationMap`. There you will find XML nodes like `<association id="xxx" langID="xxx" />`.
6. Add a new node like this: `<association id="pascal_syntax" langID="11" />`.

**Notepad++ version v7.9.1 and above**

1. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
2. Open folder `functionList` and move the file `pascal.xml` to this folder.
