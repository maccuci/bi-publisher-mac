# BIP Template Editor

A native Microsoft Word for Mac prototype for working with Oracle BI Publisher RTF templates. It adds a **BI Publisher** ribbon tab to the RTF document window through a global `.dotm` template, VBA, and Ribbon XML. A local Swift/Foundation helper reads sample XML through an AppleScriptTask bridge.

The native integration is the current product direction. It does not require Office.js, a browser, or a local HTTP server. The ribbon, dialogs, and scripts are in English.

> **Status: alpha.** This is an independent project, not affiliated with or endorsed by Oracle or Microsoft. Several commands have not yet been executed in Word (see below). Work only on copies of templates.
>
> Local preview uses the Oracle BI Publisher engine from **your own licensed Oracle installation**. No Oracle software is included in this repository. See [License and trademarks](#license-and-trademarks).

## Features and status

The native sources provide these commands:

| Command | Purpose |
| --- | --- |
| **Import XML** | Load local sample XML and build a catalog of paths for the active document. |
| **Insert Field** | Insert a legacy text form field with a display label and a BI Publisher expression in Help Text (F1). |
| **Clear XML** | Clear the sample catalog from the session. |
| **Repeat (for-each)** | Wrap selected text or table rows in a loop. |
| **Group (for-each-group)** | Group selected content by an XML field. |
| **Condition (if)** | Wrap selected content in an XPath condition. |
| **Choose (choose)** | Insert conditional branches, including `when` and `otherwise`. |
| **Preview PDF** | Render a saved RTF using a separately installed local Oracle engine. |
| **Inspect Field** | Show a field's type, expression (expanding Template Builder `<?ref:xdoNNNN?>` references), source path in the sample XML with occurrence count and example value, and the loops/conditions around it. |
| **About** | Show information about the prototype. |

Recorded tests establish a limited baseline:

- **Word, October 9, 2026:** the ribbon and About callback ran in an RTF window. A later test on a synthetic RTF copy imported six XML paths, inserted `STATUS`, and read back `<?/DATA/PERSON/STATUS?>` from Help Text.
- **Outside Word, October 12, 2026:** plugin-style fields, text and table loops, grouping with `current-group()`, and conditional branches rendered with the local Oracle engine. The preview bridge generated and opened a PDF.

These results do not establish compatibility for the current source build or acceptance of the MVP. The structure, preview, and updated inspection callbacks still need execution and **Debug → Compile** in Word, and saved edits need reopen, preservation, and authorized Oracle environment checks. See the [native guide](native/README.md) for details.

Editing an existing expression through a dedicated editor, automatic export of a protected RTF copy, remote catalog integration, and Windows support remain outside the implemented native workflow. Local preview leaves remote catalog subtemplate sections blank.

## Install from a release

End users do not need Bun, Swift, Python, or the VBA editor.

1. Download `BIPTemplateEditor-<version>.zip` from the GitHub Releases page and unzip it.
2. Quit Word (**Command-Q**) and run `Install.command`. The package is not signed by Apple, so macOS may block a double-click. Either open Terminal, type `bash `, drag `Install.command` into the window, and press Return; or double-click it, then choose **Open Anyway** in **System Settings → Privacy & Security** and open it again.
3. Open Word and an RTF template, and enable macros for this template when Word asks. The **BI Publisher** tab appears beside Home.

The installer copies the template to Word's Startup folder and the bridge, XML reader, and preview source to `~/Library/Application Scripts/com.microsoft.Word/`, keeping backups of earlier copies. It removes the macOS download quarantine from the package files it installs, since running it is the user's decision to trust them. `Uninstall.command` removes the installed files and can also remove the copied Oracle engine.

For **Preview PDF**, install Java 21 (`brew install openjdk@21`) and run `Import Oracle Engine.command`, then select your licensed `Oracle_Home` folder. The release never contains Oracle components.

## Build and install

Development requires macOS, Microsoft Word for Mac, Bun, Python 3, and a Swift compiler. Local Oracle preview additionally requires Java 21 (`brew install openjdk@21`) and the BI Publisher engine jars from a licensed Oracle installation. Oracle does not need to run on the Mac: the jars are platform-independent Java, so copy them once from an existing Oracle Home (a company BI server, a Windows VM) or from a BI Publisher Desktop (Template Builder for Word) folder. Afterwards the preview no longer depends on that installation. Oracle libraries are not included in this repository:

```sh
bun run oracle:import "/path/to/Oracle_Home"
```

From the repository root:

```sh
bun install
bun run native:build
bun run native:test
```

Without an argument, `native:build` compiles and checks the XML helper, generates ribbon icons, and stages the bridge and preview source. It does not package a Word template.

Import the VBA module into a separate base template in Word, save it, and close it as described in the [native build instructions](native/README.md#prepare-the-word-template). Then run:

```sh
bun run build
```

This packages `out/native/BIPTemplateEditor.base.dotm` into `out/native/BIPTemplateEditor.dotm`. The base must contain the VBA project saved by Word and must be updated whenever the VBA source changes. Packaging preserves the base and refuses a missing or stale VBA project.

Save your documents and quit Word with **Command-Q** before installing. Run this in your own Terminal:

```sh
bun run word:install
```

`native:install` is an alias for the same installer; `bun run native:uninstall` removes it. It copies the template to Word Startup and the helper files to Word's Application Scripts folder, backing up existing copies. It does not modify `Normal.dotm` or change global macro security.

Reopen Word with a **copy** of a synthetic RTF and check the **BI Publisher** tab. Follow the [installation and test procedure](native/README.md#install-and-test) before using the prototype with working templates.

## Publishing a release

The release must be built on the maintainer's Mac because Word must save the VBA project; it cannot be produced by CI.

1. Import the current `native/BIPTemplateEditor.bas` into `out/native/BIPTemplateEditor.base.dotm` in Word, run **Debug → Compile**, save, and close it.
2. Test the commands in Word on synthetic templates.
3. Update `version` in `package.json` and run:

   ```sh
   bun run release
   ```

The script packages the template (refusing a base older than the VBA source), runs the packaging tests, builds a universal (Apple silicon and Intel) XML reader, and writes `out/release/BIPTemplateEditor-<version>.zip` with the template, bridge, reader, preview source, `Install.command`, `Uninstall.command`, `Import Oracle Engine.command`, a `README.txt`, `LICENSE`, and `NOTICE`. It fails if any `.jar` file is present. Attach the zip to a GitHub Release.

Code signing and Apple notarization are not set up; users must allow the unsigned scripts once as described above.

## Repository layout

| Location | Contents |
| --- | --- |
| `native/` | VBA commands, Ribbon XML, AppleScript bridge, XML reader, icon generator, and local Oracle preview source. |
| `scripts/` | Native build, packaging, installation, and VBA import automation. |
| `fixtures/synthetic/` | Fictitious RTF templates and sample XML for testing. |
| `fixtures/private/` | Authorized private fixtures; ignored by Git. |
| `docs/` | Product requirements, compatibility evidence, test protocol, and architecture decisions. |
| `out/native/` | Generated native build files and Word base/final templates; ignored by Git. |
| `LICENSE`, `NOTICE` | Apache License 2.0 and third-party, Oracle, and trademark notices. |

## Verification and document integrity

```sh
bun run test
```

These checks cover XML parsing, template packaging, and import automation safeguards. They do not validate VBA compilation, Word callbacks, saved RTF preservation, or execution in an Oracle test environment.

Legacy BI Publisher expressions belong in form-field **Help Text (F1)**. Word field code such as `FORMTEXT` is a separate value, as is the visible display label.

Work on copies of RTF files. Native insertion modifies the open document, and preview asks to save it if it has unsaved changes. The native workflow does not yet provide automatic original-file protection or a reviewed export copy. Before accepting the MVP, complete the PRD's T-01 through T-05 checks on representative, authorized, anonymized templates: isolated edits, insertion, original protection, structural preservation, and rejection of unsupported input.

RTF and XML processing is local, with no automatic uploads, remote telemetry, or Oracle authentication. Native commands use temporary bridge/snippet files; preview produces local output and logs. Treat those artifacts and exported inventories as potentially sensitive. Keep private templates and report data out of version control.

An earlier Office.js task-pane experiment was abandoned because Word for Mac disables Add-ins in RTF Compatibility Mode; see [ADR-003](docs/adr/003-native-word-ribbon.md).

## License and trademarks

This project is licensed under the [Apache License 2.0](LICENSE). See [NOTICE](NOTICE) for the full third-party and trademark notices.

- BIP Template Editor is independent and is not affiliated with, endorsed by, or supported by Oracle Corporation, Microsoft Corporation, or Apple Inc.
- Oracle, Java, Oracle BI Publisher, Oracle Analytics Publisher, and Oracle Fusion are trademarks of Oracle and/or its affiliates. Microsoft and Word are trademarks of the Microsoft group of companies. Apple, Mac, macOS, and SF Symbols are trademarks of Apple Inc. They are used only to describe compatibility.
- No Oracle software is included. The optional preview runs the BI Publisher engine from an installation you provide; you are responsible for holding a valid Oracle license for it. The Apache License does not apply to Oracle software.
- Ribbon icons are rendered locally from SF Symbols at build time and are not stored in the repository.
- Never commit private or employer-owned templates or report data. `fixtures/private/` is ignored by Git for that purpose.

## Project references

- [Product requirements and acceptance criteria](docs/PRD_BI_Publisher_Word_macOS.md)
- [Native setup and command guide](native/README.md)
- [Compatibility matrix and historical evidence](docs/compatibility-matrix.md)
- [Manual test protocol](docs/test-protocol.md)
- [Native Word ribbon decision](docs/adr/003-native-word-ribbon.md)
