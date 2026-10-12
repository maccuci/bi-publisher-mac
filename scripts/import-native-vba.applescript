-- Usa a importacao do editor, sem acesso programatico ao projeto VBA.
on sameFile(documentPath, expectedFile)
    try
        if documentPath starts with "/" then
            set documentFile to POSIX file documentPath as alias
        else
            set documentFile to documentPath as alias
        end if
        considering case, diacriticals
            return (POSIX path of documentFile) is (POSIX path of expectedFile)
        end considering
    on error
        return false
    end try
end sameFile

-- O VBA pode expor o rotulo em AXValue ou em um filho, sem AXTitle/name.
on rowLabels(elementRef, remainingDepth, expectedLabel)
    set labels to {}
    tell application "/System/Library/CoreServices/System Events.app"
        repeat with attributeName in {"AXTitle", "AXValue", "AXDescription"}
            try
                set labelValue to value of attribute (contents of attributeName) of elementRef
                if class of labelValue is text then set end of labels to labelValue
            end try
        end repeat
        if labels contains expectedLabel then return labels
        if remainingDepth > 0 then
            try
                set childElements to UI elements of elementRef
                set childCount to count childElements
                if childCount > 8 then set childCount to 8
                repeat with childIndex from 1 to childCount
                    set labels to labels & my rowLabels(item childIndex of childElements, remainingDepth - 1, expectedLabel)
                    if labels contains expectedLabel then exit repeat
                end repeat
            end try
        end if
    end tell
    return labels
end rowLabels

on uniqueRowIndex(labelsByRow, expectedLabel)
    set matchIndex to 0
    repeat with rowIndex from 1 to count labelsByRow
        considering case, diacriticals
            set matches to (item rowIndex of labelsByRow) contains expectedLabel
        end considering
        if matches then
            if matchIndex is not 0 then error "More than one row matches the project/module; operation cancelled."
            set matchIndex to rowIndex
        end if
    end repeat
    return matchIndex
end uniqueRowIndex

on findRow(projectOutline, expectedLabel)
    tell application "/System/Library/CoreServices/System Events.app"
        set projectRows to rows of projectOutline
        set labelsByRow to {}
        repeat with projectRow in projectRows
            set end of labelsByRow to my rowLabels(contents of projectRow, 2, expectedLabel)
        end repeat
        set matchIndex to my uniqueRowIndex(labelsByRow, expectedLabel)
        if matchIndex is 0 then return missing value
        return item matchIndex of projectRows
    end tell
end findRow

on collapsedState(stateValue)
    set stateText to stateValue as text
    if {"0", "false", "off"} contains stateText then return true
    if {"1", "true", "on"} contains stateText then return false
    error "The Project Explorer folder state was not recognized."
end collapsedState

on expandRow(rowRef)
    tell application "/System/Library/CoreServices/System Events.app"
        if exists (first UI element of rowRef whose role is "AXDisclosureTriangle") then
            set triangle to first UI element of rowRef whose role is "AXDisclosureTriangle"
            if my collapsedState(value of triangle) then click triangle
        end if
    end tell
end expandRow

on revealModules(projectOutline)
    set modulesRow to my findRow(projectOutline, "Modules")
    if modulesRow is not missing value then my expandRow(modulesRow)
end revealModules

on run arguments
    if (count arguments) is not 2 then error "Pass the base template and the VBA module."
    set basePath to item 1 of arguments
    set modulePath to item 2 of arguments
    set baseFile to POSIX file basePath as alias
    set moduleFile to POSIX file modulePath as alias
    set editorTitle to "Microsoft Visual Basic - BIPTemplateEditor.base"
    log "Step 1/6: opening and checking the base template."
    tell application "/Applications/Microsoft Word.app"
        -- O parametro file name usa a abertura de documentos do Word.
        open file name (baseFile as text) without add to recent files
        activate
        set baseIsActive to false
        repeat 50 times
            if (count documents) > 0 then
                if my sameFile(full name of active document, baseFile) then
                    set baseIsActive to true
                    exit repeat
                end if
            end if
            delay 0.1
        end repeat
        if not baseIsActive then error "Word did not activate the base file within 5 seconds. Open out/native/BIPTemplateEditor.base.dotm with File > Open and try again."
    end tell
    tell application "/System/Library/CoreServices/System Events.app"
        tell process "Microsoft Word"
            log "Step 2/6: opening the editor and finding the base project."
            set frontmost to true
            if not (exists window editorTitle) then
                click menu item "Visual Basic Editor" of menu 1 of menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1
            end if
            repeat 50 times
                if exists window editorTitle then exit repeat
                delay 0.1
            end repeat
            if not (exists window editorTitle) then error "The VBA editor did not open."
            set editorWindow to window editorTitle
            perform action "AXRaise" of editorWindow
            set value of attribute "AXMain" of editorWindow to true
            set projectOutline to outline 1 of editorWindow
            set projectRow to my findRow(projectOutline, "TemplateProject (BIPTemplateEditor.base)")
            if projectRow is missing value then error "The base project label was not found in Project Explorer; nothing will be imported."
            click projectRow
            if (value of attribute "AXSelected" of projectRow) is not true then set value of attribute "AXSelected" of projectRow to true
            if (value of attribute "AXSelected" of projectRow) is not true then error "Could not select the base project; import cancelled."
            my expandRow(projectRow)
            my revealModules(projectOutline)
            if my findRow(projectOutline, "BIPTemplateEditor") is not missing value then error "The module already exists. A duplicate will not be imported."
            log "Step 3/6: opening Import File and selecting the module."
            set fileMenu to menu 1 of menu bar item "File" of menu bar 1
            if not (exists (first menu item of fileMenu whose name starts with "Import File")) then error "The editor is not focused: File does not show Import File."
            click (first menu item of fileMenu whose name starts with "Import File")
            delay 0.4
            keystroke "g" using {command down, shift down}
            delay 0.4
            keystroke modulePath
            keystroke return
            delay 0.4
            keystroke return
            log "Step 4/6: expanding Modules and confirming the import."
            repeat 5 times
                my revealModules(projectOutline)
                if my findRow(projectOutline, "BIPTemplateEditor") is not missing value then exit repeat
                delay 0.2
            end repeat
            if my findRow(projectOutline, "BIPTemplateEditor") is missing value then error "The import was not confirmed. Check the Word dialog."
            log "Step 5/6: checking the compile command."
            if exists menu bar item "Debug" of menu bar 1 then
                set debugMenu to menu 1 of menu bar item "Debug" of menu bar 1
                if exists (first menu item of debugMenu whose name starts with "Compile") then
                    set compileItem to first menu item of debugMenu whose name starts with "Compile"
                    if enabled of compileItem then click compileItem
                end if
            else
                log "The editor does not expose Debug > Compile; VBA execution must be checked by testing the buttons."
            end if
            delay 0.3
            if exists (first window whose subrole is "AXDialog") then error "Word showed a dialog during compilation. Check the message."
        end tell
    end tell
    tell application "/Applications/Microsoft Word.app"
        log "Step 6/6: saving and closing the base template."
        if (count documents) is 0 then error "The base template was closed; nothing will be saved."
        if not my sameFile(full name of active document, baseFile) then error "The active document changed; close cancelled."
        save active document
        close active document saving yes
    end tell
    return "BIP_VBA_IMPORTED"
end run
