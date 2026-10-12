BIP Template Editor - BI Publisher plugin for Word for Mac (alpha)
==================================================================

Independent project, not affiliated with Oracle or Microsoft. Apache 2.0
license (see LICENSE and NOTICE). Always work on copies of your templates.

INSTALL
1. Quit Word (Cmd+Q).
2. Run "Install.command". The package is not signed by Apple, so macOS may
   block a double-click. Use one of these options:
   a) Simplest: open Terminal, type  bash  and a space, drag
      "Install.command" into the window, and press Return.
   b) Or double-click it; if macOS blocks it, go to System Settings >
      Privacy & Security > "Open Anyway" and open it again.
   The same applies to the other .command files.
3. Open Word and an RTF template. The "BI Publisher" tab appears next to
   Home. When Word asks, enable macros for this template.

If the tab does not appear: Tools > Templates and Add-ins... and check
BIPTemplateEditor.dotm under "Global templates and add-ins".

PREVIEW PDF (optional)
Preview uses the Oracle BI Publisher engine from YOUR licensed Oracle
installation. No Oracle components are included in this package.
1. Install Java 21:  brew install openjdk@21
   (Apple silicon and Intel Homebrew locations are both supported, as is
   any JDK 21 registered with macOS.)
2. Run "Import Oracle Engine.command" (as in step 2 above) and choose the
   Oracle_Home folder (the one containing bip/lib and jlib). The required
   files are copied once to
   ~/Library/Application Support/BIPTemplateEditor/oracle-lib/

UNINSTALL
Quit Word and run "Uninstall.command" (as in install step 2).

Documentation and source code: see the project repository on GitHub.
