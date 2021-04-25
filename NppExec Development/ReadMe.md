# _NppExec_ Development


## Syntax Highlighting for _NppExec_ Scripts

### Remarks
This folder contains two UDL files for the _NppExec_ plugin scripting language, one for the default theme of Notepad++ (file `nppexec._userdefined.udl_default.xml`) and one for the _Material_ theme (file `nppexec._userdefined.udl_material.xml`). The latter one is not included in a standard Notepad++ installation. You can obtain it from [here](https://github.com/Codextor/npp-material-theme).

### Installation

The following steps depend on the version of Notepad++ you use.

**Notepad++ versions prior to v7.6.4**

1. Start Notepad++.
2. Open menu `Language` and click on menu entry `Define your language...`.
3. Click button `Import...`.
4. Double-click one of the above files.

**Notepad++ version v7.6.4 and above**

You can follow the steps listed above or you can do the following:

**a) If you use an installed version of Notepad++**
1. Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
2. Open folder `userDefineLangs` and copy one of the above files into this folder.

**b) If you use a portable version of Notepad++**
1. Open the folder where your `notepad++.exe` resides.
2. Open folder `userDefineLangs` and copy one of the above files into this folder.


## Function List Parser for _NppExec_ Scripts

### Prerequisites
To get a working function list for _NppExec_ scripts you also need a user defined language (UDL) for syntax highlighting _NppExec_ scripts. Thus you have to install one of the UDL files from above at first.

### Installation

The following steps depend on the version of Notepad++ you use.

**Notepad++ versions prior to v7.9.1**

1. Open file `nppexec.xml` with Notepad++.
2. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
    **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
3. Open file `functionList.xml` from that folder with Notepad++.
4. From the file `nppexec.xml`, copy the whole `/NotepadPlus/functionList/parser` XML node into the file `functionList.xml` as a child node of the `/NotepadPlus/functionList/parsers` XML node.
5. In file `functionList.xml` go to the XML node `/NotepadPlus/functionList/associationMap`. There you will find XML nodes like `<association id="xxx" userDefinedLangName="xxx" />`.
6. Add a new node like this: `<association id="nppexec_syntax" userDefinedLangName="NppExec" />`. **Please note:** The value of the `userDefinedLangName` attribute has to be the same like in the UDL file's `/NotepadPlus/UserLang` XML node's `name` attribute. In case you did not change the UDL file from above, everything is already set up the right way.

**Notepad++ version v7.9.1 and above**

1. **a) If you use an installed version of Notepad++:** Open your user profile folder and navigate to folder `AppData\Roaming\Notepad++`. If the `AppData` folder is not shown in your user profile folder, you need to go to the Windows Explorer folder options and enable the display of hidden files and folders.
   **b) If you use a portable version of Notepad++:** Open the folder where your `notepad++.exe` resides.
2. Open folder `functionList` and move the file `nppexec.xml` to this folder.
3. Open file `overrideMap.xml` from the folder `functionList` and go to the XML node `/NotepadPlus/functionList/associationMap`. There you will find XML nodes like `<association id="xxx.xml" userDefinedLangName="xxx" />`.
4. Add a new node like this: `<association id="nppexec.xml" userDefinedLangName="NppExec" />`. **Please note:** The value of the `id` attribute has to be the file name of the XML file you moved to the `functionList` folder in step 2. The value of the `userDefinedLangName` attribute has to be the same like in the UDL file's `/NotepadPlus/UserLang` XML node's `name` attribute. In case you did not change the UDL file from above, everything is already set up the right way. 
