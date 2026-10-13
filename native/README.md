# Native Word for Mac integration

The **BI Publisher** ribbon tab is provided by a global Word `.dotm` template containing VBA and Ribbon XML. A Swift/Foundation executable reads local XML through an AppleScriptTask bridge. Ribbon icons are generated from SF Symbols during the build.

This is the current integration for editing RTF templates in their Word window. Ribbon labels, dialogs, and scripts are in English; labels match the command names below. The [project README](../README.md) covers the repository and acceptance requirements.

## Commands

| Ribbon command | Behavior |
| --- | --- |
| **Import XML** | Open a native file picker and associate an XML path catalog with the active document. |
| **Insert Field** | Insert a legacy text form field at the selection, with the element name as its display label and `<?path?>` in Help Text (F1). If the selection overlaps an existing field, insert after it. |
| **Clear XML** | Remove the XML catalog and file reference from the session. |
| **Inspect Field** | With the cursor inside a field (or the field selected), show its type (data field, loop, group, condition, branch, block end, subtemplate call), expression, the source path in the imported XML with occurrence count and example value, and the enclosing loops/conditions, similar to the field properties of the Windows Template Builder. Reads `FormField.HelpText` (or the status text Template Builder uses) and expands `<?ref:xdoNNNN?>` references from document variables. There is no double-click trigger. |
| **About** | Show information about the prototype. |

Insertion rejects expressions longer than 255 UTF-16 code units. Inspection does not change the document. There is no dedicated editor for existing expressions.

The XML reader returns, for each path, the number of occurrences and the first non-empty value (up to 80 characters) in the catalog format `BIPXML2` (`path<TAB>count<TAB>example`); the VBA importer expects the same version, so install the helper and the template together.

### Loops, grouping, and conditions

- **Repeat (for-each):** choose a repeating XML element. The command surrounds selected text with `F` (`<?for-each:PATH?>`) and `E` (`<?end for-each?>`). In a table it wraps the selected rows, placing the opening field in the first cell and the closing field in the last cell.
- **Group (for-each-group):** choose an element and a grouping key within it. The command inserts `G` (`<?for-each-group:PATH;./KEY?>`) and `EG` (`<?end for-each-group?>`) around the selection or table rows. Within a group, **Repeat** offers `current-group()` to list its items.
- **Condition (if):** choose an XML field and enter an XPath condition, such as `STATUS='ACTIVE'`. The command inserts `C` (`<?if:...?>`) and `EC` (`<?end if?>`). A selection spanning multiple cells applies to whole rows.
- **Choose (choose):** create a choice with `CH`, a first `W` (`<?when:...?>`) branch, `EW`, and `ECH`. To add another `when` or an `otherwise` (`O`/`EO`), select content between `EW` and `ECH` and click the command again. Inside an existing branch, the command creates a nested choice.

Inserted fields, loops, and conditions use paths relative to their enclosing loop where applicable. These commands change the open document; use a working copy and verify the resulting fields before saving.

### Local preview

**Preview PDF** runs [BIPPreview.java](BIPPreview.java) with a separately installed Oracle engine and opens the generated PDF. If rendering fails, the bridge opens the error log. The document must already be saved as RTF and have an imported XML sample. If it has unsaved changes, the command asks to save the active document before rendering.

The engine libraries are copied once from a licensed Oracle Home into a local folder, so the preview does not depend on the original installation (for example, a shared folder from a Windows virtual machine) after the copy:

```sh
bun run oracle:import "/path/to/Oracle_Home"
```

The script copies the jars from `bip/lib` plus `jlib/orai18n.jar`, `orai18n-collation.jar`, and `orai18n-mapping.jar` to `~/Library/Application Support/BIPTemplateEditor/oracle-lib/` and records the source in `SOURCE.txt`. It also accepts a BI Publisher Desktop (Template Builder for Word) folder copied from Windows: it finds `xdocore.jar` and copies every jar from that folder. Run it again to update the engine. Oracle components are not redistributed by this project and are never committed to the repository.

| Dependency | Path |
| --- | --- |
| Java 21 | `/opt/homebrew/opt/openjdk@21/bin/java` (`brew install openjdk@21`) |
| Oracle engine | `~/Library/Application Support/BIPTemplateEditor/oracle-lib/` |

The bridge creates a temporary directory containing `preview.pdf` and `preview.log`; those files are not automatically removed. Remote catalog subtemplate imports (`xdoxsl:///`) are replaced with empty templates, so their sections appear blank. A successful local preview does not establish equivalence to the report in an authorized Oracle environment.

For a separate command-line render from the repository root:

```sh
bun run preview path/to/template.rtf path/to/sample.xml /private/tmp/bip-preview.pdf
```

The command-line script uses the same local engine folder; `BIP_ORACLE_LIB` overrides it for that script only, not for the Word bridge. The output extension selects PDF, HTML, or RTF.

## Recorded validation and current limitations

On **October 9, 2026**, tests in Word on a synthetic RTF copy established the ribbon and About callback, import of six paths from `sample-data.xml`, insertion of a gray-shaded `STATUS` field, and inspection of its Help Text as `<?/DATA/PERSON/STATUS?>`. A saved round-trip and Oracle execution of that Word-edited copy were not established.

On **October 12, 2026**, tests outside Word rendered plugin-style Help Text fields, text/table `for-each` loops, `for-each-group` with `current-group()`, `if`, and `choose/when/otherwise` with the local Oracle engine, using the locally copied engine folder. The preview bridge generated and opened a PDF. A simulation of the inspection's block and `<?ref:...?>` logic on a real Template Builder template resolved all 484 field expressions with balanced blocks. The new structure, preview, and inspection VBA callbacks were not executed or compiled in Word.

Automatic loading after installation/restart, the updated ribbon callbacks, and preservation of unrelated fields and layout after saving still require host verification. The prototype does not provide a protected-copy export workflow or full Oracle Template Builder functionality.

## Prepare the Word template

Requirements: macOS, Microsoft Word for Mac, Bun, Python 3, and a Swift compiler. Run build commands from the repository root.

The VBA project must be imported and saved by Word in a separate macro-enabled base template. Keep it synchronized with `BIPTemplateEditor.bas`; packaging does not compile VBA.

### Manual import or update

1. Open `out/native/BIPTemplateEditor.base.dotm` in Word for editing with **File → Open**, rather than creating a document from it. If there is no base yet, save a separate blank **Word Macro-Enabled Template (.dotm)** at that path.
2. Open **Tools → Macro → Visual Basic Editor**. Select **TemplateProject (BIPTemplateEditor.base)** in Project Explorer. Do not select **Normal**.
3. Use **File → Import File…** to import `native/BIPTemplateEditor.bas`. When updating an existing base, replace its existing `BIPTemplateEditor` module instead of importing a duplicate. Keep a backup of the base before replacing the module.
4. Confirm the module belongs to the correct project. Use **Debug → Compile** if that menu is available, then save the template. If Word blocks import, record the error; do not weaken global macro security or enable programmatic VBA project access to bypass it.
5. Close the base template, then package it:

   ```sh
   bun run native:build out/native/BIPTemplateEditor.base.dotm
   ```

The command compiles and checks the XML reader, generates icons, stages the bridge/preview source, and creates **out/native/BIPTemplateEditor.dotm**. It preserves the base and its VBA project. It rejects a base older than the VBA source, a missing VBA project, or a signed development template whose signature would be invalidated.

`bun run build` packages the default base at the same path. Without a template argument, `bun run native:build` builds the helper assets only. Packaging tests verify package structure; they do not establish VBA compilation or callback execution in Word.

### Optional first-import automation

For a base that exists but has no VBA project, run this in your own Terminal from the repository root:

```sh
bun run native:setup
```

The assistant backs up the base, opens the VBA editor, selects only **TemplateProject (BIPTemplateEditor.base)**, attempts import/compilation, saves and closes the base, then builds. It waits for you to quit Word before installing. It refuses to reimport a base that already contains VBA; use the manual update procedure for that base.

Automation depends on the installed Word's **English menus** and macOS Automation/Accessibility permissions. It verifies the active file's normalized HFS/POSIX identity before import and save, requires one exact project/module match, expands the Modules folder, and reports progress. Unconfirmed import or a 60-second timeout stops the process before build or installation. Interrupt a stalled previous attempt with **Control-C** before retrying.

This assistant has not been validated end to end. Agent access to System Events previously failed with `-10827`; run it in your own Terminal. A module left open after interruption does not prove it was compiled or saved. The assistant does not change macro security or enable programmatic VBA project access.

## Install and test

After packaging the final plugin, save your documents and quit Word with **Command-Q**. Run this in your own Terminal:

```sh
bun run native:install
```

The installer writes:

| Destination | Files |
| --- | --- |
| `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/` | `BIPTemplateEditor.dotm` |
| `~/Library/Application Scripts/com.microsoft.Word/` | `BIPBridge.applescript`, `bip-xml`, `BIPPreview.java` |

Existing copies receive timestamped backups; other templates and `Normal.dotm` are untouched. The installer requires Word to be closed. If Word uses a custom Startup folder, place the global template there. Installation writes outside the repository and may require execution from the user's Terminal.

1. Open Word and a **copy** of `fixtures/synthetic/02-legacy-fields.rtf`.
2. Find **BI Publisher** beside Home. If Word asks about macros, authorize only this reviewed template according to your local policy.
3. Click **Import XML** and select `fixtures/synthetic/sample-data.xml`. Expect six paths, including `/DATA/PERSON/SALARY_AMOUNT`.
4. Place the cursor at a test position, click **Insert Field**, and choose that path. Expect a `SALARY_AMOUNT` field with Help Text `<?/DATA/PERSON/SALARY_AMOUNT?>` outside a loop.
5. Place the cursor inside the new field and click **Inspect Field**. Expect type *Campo de dados*, the expression, and origin `/DATA/PERSON/SALARY_AMOUNT` with one occurrence and example `3750.00`.
6. Save a separate RTF copy, reopen it, and compare all unrelated fields and layout. Test the copy in an authorized Oracle environment with synthetic XML before treating it as usable.

If the tab is missing, check **Tools → Templates and Add-ins…** and enable `BIPTemplateEditor.dotm` under **Global Templates and Add-ins**. Keep only one copy loaded if both `out/native` and Startup versions appear. If Word cannot open an older package, rebuild and reinstall the corrected package while Word is closed; that package error alone does not require another VBA import.

## XML handling and host behavior

The catalog is associated with the document active at import time. Switching documents requires importing the appropriate sample. Importing another sample replaces the catalog; **Clear XML** clears it. Canceling import retains the previous sample.

The reader accepts UTF-8 XML up to 10 MB, with limits of 128 nesting levels, 200,000 elements, 10,000 unique paths, and 4,096 UTF-8 bytes per path. It includes attributes, occurrence counts, and sample values, and rejects DTDs and external entities. Configure namespaces in the template before using qualified references.

Processing stays local. The catalog and source XML path are held for the session, and preview reads that source file again. The bridge exchanges catalog results through temporary files and deletes the result after a successful VBA read; insertion similarly uses a temporary RTF snippet. Preview output/logs remain on disk. Treat temporary artifacts as report data.

The implementation handles these observed Word Mac behaviors:

- `AppleScriptTask` returned to VBA as soon as its dialog closed, before the script finished, with an empty return value. The bridge writes an atomic result file; VBA waits for it.
- `FileLen` returned zero for a missing file without an error. The wait checks `Dir` instead.
- `DoEvents` during that wait crashed Word. The code uses `usleep`.
- `FormFields.Add` created an empty field that would not accept text. Insertion uses an RTF snippet through `InsertFile` to include both the display label and Help Text.

## License

Apache License 2.0; see [LICENSE](../LICENSE) and [NOTICE](../NOTICE). No Oracle components are included or redistributed; the preview uses the engine from the user's own licensed Oracle installation.

## Checks

```sh
bun run native:build
bun run test
```

The first command runs XML self-checks and builds native assets. Tests cover packaging and import automation safeguards. These checks do not replace manual Word tests or the preservation and Oracle acceptance gates in the [PRD](../docs/PRD_BI_Publisher_Word_macOS.md).
