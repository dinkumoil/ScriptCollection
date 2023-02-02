# Pascal Development

The files in this directory provide function list parsers for Notepad++. After installing these files you will be able to use the function list panel of Notepad++ to navigate through your code files.

* The files `pascal_old.xml` and `pascal_new.xml` provide function list parsers for _Pascal/Delphi_ source code files. For the difference of these files see the notes below.
* The file `delphiForm.xml` provides a function list parser for _Delphi_ forms files, i.e. `*.dfm` files.

**Please note:** The parsers may have issues and thus may not work always as expected. I will **not react** to bug reports but fixes are welcome. If you find some unwanted behaviour please fix it by yourself and file a pull request.


## Installation

### _Pascal/Delphi_ function list parser

**Please note:** There are two versions of this parser, one for Notepad++ versions up to v8.4.8 and one for version v8.4.9 and above. You should use these parsers only with the recommended Notepad++ versions (see installation instructions below)! Due to some bugs in older versions of Notepad++ the new parser may not work reliable or cause Notepad++ to hang when used with unsuitable versions.

The following steps depend on the version of Notepad++ you use.

**Notepad++ versions prior to v7.9.1**

1. Open file `pascal_old.xml` with Notepad++.
2. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
3. Open file `functionList.xml` from that folder with Notepad++.
4. From the file `pascal_old.xml`, copy the whole `/NotepadPlus/functionList/parser` XML node into the file `functionList.xml` as a child node of the `/NotepadPlus/functionList/parsers` XML node.
5. In file `functionList.xml` go to the XML node `/NotepadPlus/functionList/associationMap`. There you will find XML nodes like `<association id="xxx" langID="xxx" />`.
6. Add a new node like this: `<association id="pascal_syntax" langID="11" />`.

**Notepad++ version v7.9.1 up to v8.4.8**

1. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
2. Open folder `functionList` and move the file `pascal_old.xml` to this folder.
3. Rename file `pascal_old.xml` to `pascal.xml`.

**Notepad++ version v8.4.9 and above**

1. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
2. Open folder `functionList` and move the file `pascal_new.xml` to this folder.
3. Rename file `pascal_new.xml` to `pascal.xml`.


### _Delphi_ forms function list parser

Notepad++ provides no build-in lexer (aka syntax highlighter) for _Delphi_ forms files. For this reason you need to install at first an _UDL_ (user defined lexer) for this file type to be able to take advantage of the function list parser. You can find a _Delphi_ forms UDL in the [official repository for _Notepad++_ UDLs](https://github.com/notepad-plus-plus/userDefinedLanguages/tree/master/UDLs). The main page of this repository also provides instructions how to install a UDL.

When you are done with installing the UDL, you can install the function list parser. The following steps depend on the version of Notepad++ you use.

**Notepad++ versions prior to v7.9.1**

1. Open file `delphiForm.xml` with Notepad++.
2. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
3. Open file `functionList.xml` from that folder with Notepad++.
4. From the file `delphiForm.xml`, copy the whole `/NotepadPlus/functionList/parser` XML node into the file `functionList.xml` as a child node of the `/NotepadPlus/functionList/parsers` XML node.
5. In file `functionList.xml` go to the XML node `/NotepadPlus/functionList/associationMap`. There you will find pairs of XML nodes like `<association id="xxx" userDefinedLangName="xxx"/>` and `<association id="xxx" ext=".xxx" />`.
6. Add a new pair of nodes like this: `<association id="dfm_syntax" userDefinedLangName="Delphi Form" />` and `<association id="dfm_syntax" ext=".dfm" />`.

**Notepad++ version v7.9.1 and above**

1. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
2. Open folder `functionList` and move the file `delphiForm.xml` to this folder.
3. In folder `functionList` open file `overrideMap.xml` with Notepad++.
4. In file `overrideMap.xml` go to XML node `/NotepadPlus/functionList/associationMap`. There you will find XML nodes like `<association id="xxx.xml" userDefinedLangName="xxx" />`.
5. Add a new node like this: `<association id="delphiForm.xml" userDefinedLangName="Delphi Form" />`
