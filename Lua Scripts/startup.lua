-- Startup script
-- Changes will take effect once Notepad++ is restarted


-- =============================================================================
-- Add menu entries to transpose multiple selections by moving them up or down
-- =============================================================================

-- Constants for transpose directions
local TRANSPOSE_SEL_DIR_UP   = 1
local TRANSPOSE_SEL_DIR_DOWN = 2


-- Get text of selection with index idx
local function getSelectionNText(idx)
  local retVal = editor:textrange(editor.SelectionNStart[idx], editor.SelectionNEnd[idx])
  if not retVal then retVal = "" end

  return retVal
end


-- Set text of selection with index idx and readjust its span
local function setSelectionNText(idx, str)
  editor:SetTargetRange(editor.SelectionNStart[idx], editor.SelectionNEnd[idx])
  editor:ReplaceTarget(str)

  editor.SelectionNAnchor[idx] = editor.TargetStart
  editor.SelectionNCaret[idx]  = editor.SelectionNAnchor[idx] + #str
end


-- Use Insertion Sort to create a sorted list of selection indices
-- Using table.sort is avoided since it requires usage of 1 based indices
local function sortSelections(selCnt)
  local selPos, idxDst
  local selections = {}

  -- Create initial list
  for idx = 0, selCnt - 1 do
    selections[idx] = idx
  end

  -- Sort list
  for idxSrc = 1, selCnt - 1 do
    selPos = editor.SelectionNStart[idxSrc]
    idxDst = idxSrc

    while (idxDst > 0) and (editor.SelectionNStart[selections[idxDst - 1]] > selPos) do
      selections[idxDst] = selections[idxDst - 1]
      idxDst = idxDst - 1
    end

    selections[idxDst] = idxSrc
  end

  return selections
end


-- Transpose selections
-- If there is only 1 selection, the standard LineTranspose function is used
-- If there are more than 2 selections, they get rotated according to parameter
-- direction
local function transposeSelections(direction)
  local selCnt, bufStr
  local startIdx, endIdx, increment
  local srcIdx, dstIdx

  selCnt = editor.Selections

  -- Process special cases:
  --   0 selections -> do nothing
  --   1 selection  -> perform SCI_LINETRANSPOSE
  if selCnt < 1 then
    return
  elseif selCnt == 1 then
    editor:LineTranspose()
    return
  end

  -- Set loop parameters depending on transpose direction
  local params = {{0, selCnt - 1, 1}, {selCnt - 1, 0, -1}}
  startIdx, endIdx, increment = (function(p) return p[1], p[2], p[3] end)(params[direction])

  -- Create list of selection indices sorted by
  -- the starting position of the related selection
  local selections = sortSelections(selCnt)

  -- Set start of undo sequence
  editor:BeginUndoAction()

  -- Perform transpose operation
  bufStr = getSelectionNText(selections[startIdx])
  dstIdx = startIdx

  while dstIdx ~= endIdx do
    srcIdx = (dstIdx + increment) % selCnt
    setSelectionNText(selections[dstIdx], getSelectionNText(selections[srcIdx]))
    dstIdx = dstIdx + increment
  end

  setSelectionNText(selections[dstIdx], bufStr)

  -- Set end of undo sequence
  editor:EndUndoAction()
end


-- Add menu entry for transposing lines upwards
npp.AddShortcut("Transpose Selections Up", "", function()
  transposeSelections(TRANSPOSE_SEL_DIR_UP)
end)


-- Add menu entry for transposing lines downwards
npp.AddShortcut("Transpose Selections Down", "", function()
  transposeSelections(TRANSPOSE_SEL_DIR_DOWN)
end)



-- =============================================================================
-- Add menu entry to revert the order of the lines in a single selection
-- or in multiple selections
-- =============================================================================

-- Revert order of selections
local function revertSelections()
  local selCnt, bufStr, srcStr
  local startIdx, endIdx
  local srcIdx, dstIdx

  selCnt = editor.Selections

  -- Process special case:
  --   less than 2 selections -> do nothing
  if selCnt < 1 then return end

  if selCnt < 2 then
    -- Perform revert operation
    startIdx = editor:LineFromPosition(editor.SelectionStart)
    endIdx   = editor:LineFromPosition(editor.SelectionEnd) - 1

    if endIdx - startIdx < 1 then return end

    -- Set start of undo sequence
    editor:BeginUndoAction()

    for dstIdx = startIdx, math.floor((startIdx + endIdx) / 2) do
      srcIdx = startIdx + (endIdx - dstIdx)

      editor:SetTargetRange(editor:PositionFromLine(srcIdx), editor.LineEndPosition[srcIdx])
      srcStr = editor.TargetText

      editor:SetTargetRange(editor:PositionFromLine(dstIdx), editor.LineEndPosition[dstIdx])
      bufStr = editor.TargetText

      editor:ReplaceTarget(srcStr)

      editor:SetTargetRange(editor:PositionFromLine(srcIdx), editor.LineEndPosition[srcIdx])
      editor:ReplaceTarget(bufStr)
    end

    -- Set end of undo sequence
    editor:EndUndoAction()
  else
    -- Create list of selection indices sorted by
    -- the starting position of the related selection
    local selections = sortSelections(selCnt)

    -- Set start of undo sequence
    editor:BeginUndoAction()

    -- Perform revert operation
    for dstIdx = 0, math.floor(selCnt / 2) - 1 do
      srcIdx = (selCnt - 1) - dstIdx
      bufStr = getSelectionNText(selections[dstIdx])

      setSelectionNText(selections[dstIdx], getSelectionNText(selections[srcIdx]))
      setSelectionNText(selections[srcIdx], bufStr)
    end

    -- Set end of undo sequence
    editor:EndUndoAction()
  end
end


-- Add menu entry for reverting lines
npp.AddShortcut("Revert Selections", "", function()
  revertSelections()
end)



-- =============================================================================
-- Add menu entry to select all occurences of word under or next to cursor or
-- already selected word
-- =============================================================================

-- Add all occurences of selected word to selection
local function selectionAddAll()
  -- From SciTEBase.cxx
  local flags     = SCFIND_WHOLEWORD -- can use 0
  local startWord = -1
  local endWord   = -1
  local s         = ""

  editor.SearchFlags = flags

  if editor.SelectionEmpty or not editor.MultipleSelection then
    startWord = editor:WordStartPosition(editor.CurrentPos, true)
    endWord   = editor:WordEndPosition(startWord, true)

    editor:SetSelection(startWord, endWord)
    
    if not editor.MultipleSelection then
      return
    end
  else
    local i   = editor.MainSelection

    startWord = editor.SelectionNStart[i]
    endWord   = editor.SelectionNEnd[i]
  end

  s = editor:textrange(startWord, endWord)

  while true do
    editor:SetTargetRange(0, editor.TextLength)

    local i            = editor.MainSelection
    local searchRanges = {{editor.SelectionNEnd[i], editor.TargetEnd}, {editor.TargetStart, editor.SelectionNStart[i]}}
    local itemFound    = false

    for _, range in pairs(searchRanges) do
      editor:SetTargetRange(range[1], range[2])

      if editor:SearchInTarget(s) ~= -1 then
        editor:AddSelection(editor.TargetStart, editor.TargetEnd)
        itemFound = true
        break
      end
    end

    if editor.TargetStart == startWord and
       editor.TargetEnd   == endWord   or
       not itemFound                   then
      break
    end
  end

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end


-- Add menu entry 
npp.AddShortcut("Selection Add All", "", function()
  selectionAddAll()
end)



-- =============================================================================
-- Add menu entry to select next occurence of word under or next to cursor or
-- already selected word
-- =============================================================================

-- Add next occurence of selected word to selection
local function selectionAddNext()
  -- From SciTEBase.cxx
  local flags = SCFIND_WHOLEWORD -- can use 0

  editor:SetTargetRange(0, editor.TextLength)
  editor.SearchFlags = flags

  if editor.SelectionEmpty then
    local startWord = editor:WordStartPosition(editor.CurrentPos, true)
    local endWord   = editor:WordEndPosition(startWord, true)

    editor:SetSelection(startWord, endWord)
  else
    local i = editor.MainSelection
    local s = editor:textrange(editor.SelectionNStart[i], editor.SelectionNEnd[i])
    local searchRanges = {{editor.SelectionNEnd[i], editor.TargetEnd}, {editor.TargetStart, editor.SelectionNStart[i]}}

    for _, range in pairs(searchRanges) do
      editor:SetTargetRange(range[1], range[2])

      if editor:SearchInTarget(s) ~= -1 then
        if editor.MultipleSelection then
          editor:AddSelection(editor.TargetStart, editor.TargetEnd)
        else
          editor:SetSelection(editor.TargetStart, editor.TargetEnd)
        end

        -- To scroll main selection in sight
        editor:ScrollRange(editor.TargetStart, editor.TargetEnd)

        break
      end
    end
  end

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end


-- Add menu entry 
npp.AddShortcut("Selection Add Next", "", function()
  selectionAddNext()
end)



-- =============================================================================
-- Add menu entry to select next occurence of word under or next to cursor or
-- already selected word but remove selection from word under cursor
-- =============================================================================

-- Remove current selected word from selection and advance to its next occurence
local function selectionSkipCurrent()
  -- From SciTEBase.cxx
  local flags = SCFIND_WHOLEWORD -- can use 0

  editor:SetTargetRange(0, editor.TextLength)
  editor.SearchFlags = flags

  if editor.SelectionEmpty then
    local startWord   = editor:WordStartPosition(editor.CurrentPos, true)
    local endWord     = editor:WordEndPosition(startWord, true)

    editor:SetSelection(startWord, endWord)
  else
    local i = editor.MainSelection
    local s = editor:textrange(editor.SelectionNStart[i], editor.SelectionNEnd[i])
    local searchRanges = {{editor.SelectionNEnd[i], editor.TargetEnd}, {editor.TargetStart, editor.SelectionNStart[i]}}

    for _, range in pairs(searchRanges) do
      editor:SetTargetRange(range[1], range[2])

      if editor:SearchInTarget(s) ~= -1 then
        if editor.MultipleSelection then
          editor:AddSelection(editor.TargetStart, editor.TargetEnd)
        else
          editor:SetSelection(editor.TargetStart, editor.TargetEnd)
        end

        -- To scroll main selection in sight
        editor:ScrollRange(editor.TargetStart, editor.TargetEnd)

        break
      end
    end

    editor:DropSelectionN(editor.Selections - 2)
  end

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end


-- Add menu entry 
npp.AddShortcut("Selection Skip Current", "", function()
  selectionSkipCurrent()
end)



-- =============================================================================
-- Add menu entry to remove selection from word under cursor
-- =============================================================================

-- Remove word under cursor from selection
local function selectionRemoveCurrent()
  if editor.Selections == 1 or not editor.MultipleSelection then
    editor.SelectionNCaret[editor.Selections - 1] = editor.SelectionNAnchor[editor.Selections - 1]
  else
    editor:DropSelectionN(editor.Selections - 1)
  end
  
  editor:ScrollCaret()

  -- To turn on Notepad++ multi select markers
  editor:LineScroll(0, 1)
  editor:LineScroll(0, -1)
end


-- Add menu entry 
npp.AddShortcut("Selection Remove Current", "", function()
  selectionRemoveCurrent()
end)



-- =============================================================================
-- Add menu entry to select all items marked by find marks
-- =============================================================================

-- Select all items marked by find marks
local function selectMarked()
  local indicatorIdx
  local pos, markStart, markEnd
  local hasMarkedItems
  
  -- Search find marks have indicator number 31
  SCE_UNIVERSAL_FOUND_STYLE = 31

  -- Remove selection but do not change cursor position
  editor.Anchor  = editor.CurrentPos
  hasMarkedItems = false
  
  -- Search find mark indicators and select marked text
  pos = 0
  
  while pos < editor.TextLength do
    if editor:IndicatorValueAt(SCE_UNIVERSAL_FOUND_STYLE, pos) == 1 then
      markStart = editor:IndicatorStart(SCE_UNIVERSAL_FOUND_STYLE, pos)
      markEnd   = editor:IndicatorEnd(SCE_UNIVERSAL_FOUND_STYLE, pos)

      if editor.SelectionEmpty or not editor.MultipleSelection then
        editor:SetSelection(markEnd, markStart)
      else
        editor:AddSelection(markEnd, markStart)
      end

      pos            = markEnd
      hasMarkedItems = true
    end
    
    pos = pos + 1
  end
  
  -- Scroll last selected element into view. This is only because
  -- Npp/Scintilla will do that anyway when moving the cursor next time
  if hasMarkedItems then
    editor:ScrollCaret()
  end
end


-- Add menu entry 
npp.AddShortcut("Select Marked", "", function()
  selectMarked()
end)



-- =============================================================================
-- Auto indentation for Lua
-- =============================================================================

-- Regexs to determine when to indent or unindent
-- From: https://github.com/sublimehq/Packages/blob/master/Lua/Indent.tmPreferences
local decreaseIndentPattern = [[^\s*(elseif|else|end|until|\})\s*$]]
local increaseIndentPattern = [[^\s*(else|elseif|for|(local\s+)?function|if|repeat|while)\b((?!end).)*$|\{\s*$]]

do_increase = false


-- Get the start and end position of a specific line number
local function getLinePositions(line_num)
  local start_pos = editor:PositionFromLine(line_num)
  local end_pos = start_pos + editor:LineLength(line_num)

  return start_pos, end_pos
end


-- Check any time a character is added
local function autoIndent_OnChar(ch)
  if ch == "\n" then
    -- Get the previous line
    local line_num = editor:LineFromPosition(editor.CurrentPos) - 1
    local start_pos, end_pos = getLinePositions(line_num)

    if editor:findtext(increaseIndentPattern, SCFIND_REGEXP, start_pos, end_pos) then
      -- This has to be delayed because N++'s auto-indentation hasn't triggered yet
      do_increase = true
    end
  else
    local line_num = editor:LineFromPosition(editor.CurrentPos)
    local start_pos, end_pos = getLinePositions(line_num)

    if editor:findtext(decreaseIndentPattern, SCFIND_REGEXP, start_pos, end_pos) then
      -- The pattern matched, now check the previous line's indenation
      if line_num > 1 and editor.LineIndentation[line_num - 1] <= editor.LineIndentation[line_num] then
        editor.LineIndentation[line_num] = editor.LineIndentation[line_num] - 2
      end
    end
  end

  return false
end


-- Work around N++'s auto indentation by delaying the indentation change
local function autoIndent_OnUpdateUI(flags)
  if do_increase then
    do_increase = false
    -- Now the the indentation can be increased since N++'s auto-indentation is done by now
    editor:Tab()
  end

  return false
end


-- See if the auto indentation should be turned on or off
local function checkAutoIndent(bufferid)
  if npp.BufferLangType[bufferid] == L_LUA then
    do_increase = false
    -- Register the event handlers
    npp.AddEventHandler("OnChar", autoIndent_OnChar)
    npp.AddEventHandler("OnUpdateUI", autoIndent_OnUpdateUI)
  else
    -- Remove the event handlers
    npp.RemoveEventHandler("OnChar", autoIndent_OnChar)
    npp.RemoveEventHandler("OnUpdateUI", autoIndent_OnUpdateUI)
  end
end


-- Only turn on the auto indentation when it is actually a lua file
-- Check it when the file is switched to and if the language changes

npp.AddEventHandler("OnSwitchFile", function(filename, bufferid)
  checkAutoIndent(bufferid)
  return false
end)


npp.AddEventHandler("OnLangChange", function()
  checkAutoIndent(npp.CurrentBufferID)
  return false
end)
