# LuaScripts

This is a collection of scripts for the _Lua Script_ plugin for Notepad++. To use them copy the file _startup.lua_ to `%AppData%\Notepad++\plugins\config` where the config files of all your Notepad++ plugins reside. The next time Notepad++ is started you will find new entries in `Plugins -> LuaScript` menu. Then you can assign keyboard shortcuts to them.

* `transposeSelections` and helpers  - Transpose selected lines.
* `revertSelections`  -  Revert order of selected lines.
* `selectionAddAll`  -  Select all occurences of word under the caret/currently selected word.
* `selectionAddNext`  - Add next occurence of word under caret/currently selected word to selection.
* `selectionSkipCurrent`  -  Remove word under caret/currently selected word from selection and advance to its next occurence.
* `selectionRemoveCurrent`  -  Only remove word under caret/currently selected word from selection, do not advance.
* `selectMarked`  -  Select all items marked by find marks.
* `documentEndRectExtend`  -  Create column selection to last line of file.
* `documentStartRectExtend`  -  Create column selection to first line of file.
* `getLinePositions`, `autoIndent_OnChar`, `autoIndent_OnUpdateUI` and `checkAutoIndent`  -  Provide auto-indentation for Lua language.
