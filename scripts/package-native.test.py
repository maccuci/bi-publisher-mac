from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import unittest
import os
import subprocess
import shutil
import sys
import io
from unittest.mock import patch
import xml.etree.ElementTree as ET

spec = spec_from_file_location('package_native', Path(__file__).with_name('package-native.py'))
module = module_from_spec(spec)
spec.loader.exec_module(module)
runner_spec = spec_from_file_location('run_native_import', Path(__file__).with_name('run-native-import.py'))
runner = module_from_spec(runner_spec)
runner_spec.loader.exec_module(runner)

class Packaging(unittest.TestCase):
    def base(self, path, macro=True):
        with ZipFile(path, 'w') as z:
            z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/></Types>')
            z.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="officeDocument" Target="word/document.xml"/></Relationships>')
            z.writestr('word/document.xml', '<document/>')
            if macro:
                z.writestr('word/vbaProject.bin', bytes.fromhex('d0cf11e0a1b11ae1') + b'SYNTHETIC-NOT-EXECUTABLE')

    def test_ribbon_and_original_preserved(self):
        with TemporaryDirectory() as directory:
            source, output = Path(directory)/'base.dotm', Path(directory)/'plugin.dotm'
            self.base(source)
            original = source.read_bytes()
            module.package(source, output)
            self.assertEqual(source.read_bytes(), original)
            with ZipFile(source) as before, ZipFile(output) as after:
                self.assertEqual(before.read('word/vbaProject.bin'), after.read('word/vbaProject.bin'))
                self.assertEqual(before.read('word/document.xml'), after.read('word/document.xml'))
                relationships = ET.fromstring(after.read('_rels/.rels'))
                self.assertTrue(any(n.get('Target') == 'customUI/customUI.xml' for n in relationships))
                self.assertEqual(after.read('customUI/customUI.xml'), (module.ROOT/'native/customUI.xml').read_bytes())
                images = ET.fromstring(after.read('customUI/_rels/customUI.xml.rels'))
                self.assertEqual({n.get('Id') for n in images}, {'iconImportXml', 'iconInsertField', 'iconForEach', 'iconGroup', 'iconIf', 'iconChoose', 'iconPreview', 'iconClearXml', 'iconInspectField', 'iconAbout'})
                for n in images:
                    self.assertTrue(after.read('customUI/'+n.get('Target')).startswith(b'\x89PNG'))
                types = ET.fromstring(after.read('[Content_Types].xml'))
                self.assertEqual(types.tag, '{http://schemas.openxmlformats.org/package/2006/content-types}Types')
                self.assertTrue(any(n.get('Extension') == 'png' for n in types))
            with self.assertRaises(ValueError): module.package(source, source)

    def test_no_fake_macro_plugin(self):
        with TemporaryDirectory() as directory:
            source, output = Path(directory)/'empty.dotm', Path(directory)/'plugin.dotm'
            self.base(source, macro=False)
            with self.assertRaisesRegex(ValueError, 'VBA compiled'): module.package(source, output)
            self.assertFalse(output.exists())

    def test_core_dates_keep_their_type_namespace(self):
        with TemporaryDirectory() as directory:
            source, output = Path(directory)/'base.dotm', Path(directory)/'plugin.dotm'
            self.base(source)
            with ZipFile(source, 'a') as z:
                z.writestr('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Test author</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-09T18:13:00Z</dcterms:created></cp:coreProperties>')
            original = source.read_bytes()
            module.package(source, output)
            self.assertEqual(source.read_bytes(), original)
            with ZipFile(output) as z:
                data = z.read('docProps/core.xml')
            namespaces = dict(value for _, value in ET.iterparse(io.BytesIO(data), events=('start-ns',)))
            root = ET.fromstring(data)
            created = root.find('{http://purl.org/dc/terms/}created')
            type_name = created.get('{http://www.w3.org/2001/XMLSchema-instance}type')
            prefix, local = type_name.split(':')
            self.assertEqual(namespaces.get(prefix), 'http://purl.org/dc/terms/')
            self.assertEqual(local, 'W3CDTF')
            self.assertEqual(created.text, '2026-10-09T18:13:00Z')
            self.assertEqual(root.find('{http://purl.org/dc/elements/1.1/}creator').text, 'BIP Template Editor')

class Setup(unittest.TestCase):
    def test_blocked_automation_times_out_without_confirmation(self):
        with patch.object(runner.subprocess, 'run', side_effect=subprocess.TimeoutExpired('osascript', 60)) as invoke:
            with self.assertRaisesRegex(ValueError, '60 seconds'):
                runner.run_import('/tmp/base.dotm', '/tmp/module.bas')
            self.assertEqual(invoke.call_args.kwargs['timeout'], 60)
            self.assertNotIn('stderr', invoke.call_args.kwargs) # Progresso fica visivel no Terminal.

    @unittest.skipUnless(sys.platform == 'darwin' and Path('/Applications/Microsoft Word.app').exists(), 'Requer AppleScript e dicionario do Word Mac')
    def test_row_selection_requires_unique_exact_label(self):
        handlers = (module.ROOT/'scripts/import-native-vba.applescript').read_text().split('on run arguments', 1)[0]
        driver = handlers + '''on run arguments
            repeat with stateValue in {0, false, "0", "false", "off"}
                if not my collapsedState(contents of stateValue) then error "Collapsed folder not recognized"
            end repeat
            repeat with stateValue in {1, true, "1", "true", "on"}
                if my collapsedState(contents of stateValue) then error "Pasta expandida seria recolhida"
            end repeat
            set expectedLabel to "TemplateProject (BIPTemplateEditor.base)"
            set labelsByRow to {{"Normal"}, {"", "AXRow", expectedLabel, expectedLabel}, {"ThisDocument"}}
            if my uniqueRowIndex(labelsByRow, expectedLabel) is not 2 then error "Value/child label not selected"
            if my uniqueRowIndex({{"Normal"}, {"TemplateProject (outro)"}}, expectedLabel) is not 0 then error "Projeto errado aceito"
            if my uniqueRowIndex({{"TemplateProject (BIPTemplateEditor.base)"}}, "BIPTemplateEditor") is not 0 then error "Projeto confundido com modulo"
            set duplicateRejected to false
            try
                my uniqueRowIndex({{expectedLabel}, {expectedLabel}}, expectedLabel)
            on error
                set duplicateRejected to true
            end try
            if not duplicateRejected then error "Projeto ambiguo aceito"
            return "OK"
        end run'''
        result = subprocess.run(['osascript', '-e', driver], text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), 'OK')

    @unittest.skipUnless(sys.platform == 'darwin' and Path('/Applications/Microsoft Word.app').exists(), 'Requer compilador AppleScript e dicionario do Word Mac')
    def test_document_identity_accepts_hfs_and_posix_but_rejects_other_files(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            base = root/'Ação com espaços.dotm'
            other = root/'outro'/'Ação com espaços.dotm'
            other.parent.mkdir()
            base.touch()
            other.touch()
            compiled = root/'import.scpt'
            subprocess.run(['osacompile', '-o', str(compiled), str(module.ROOT/'scripts/import-native-vba.applescript')], check=True, capture_output=True)
            # Executa o handler real isolado; nao abre Word nem System Events.
            identity = (module.ROOT/'scripts/import-native-vba.applescript').read_text().split('on run arguments', 1)[0]
            driver = identity + '''on run arguments
                set expectedFile to POSIX file (item 1 of arguments) as alias
                set otherFile to POSIX file (item 2 of arguments) as alias
                if not my sameFile(POSIX path of expectedFile, expectedFile) then error "POSIX rejeitado"
                if not my sameFile(expectedFile as text, expectedFile) then error "HFS rejeitado"
                if my sameFile(POSIX path of otherFile, expectedFile) then error "Other file accepted"
                if my sameFile(otherFile as text, expectedFile) then error "Outro HFS aceito"
                if my sameFile("Document1", expectedFile) then error "Unsaved document accepted"
                return "OK"
            end run'''
            result = subprocess.run(['osascript', '-e', driver, str(base), str(other)], text=True, capture_output=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(result.stdout.strip(), 'OK')

    def test_import_failure_stops_build_and_install(self):
        for response in ('exit 1', 'echo UNCONFIRMED', 'echo BIP_VBA_IMPORTED'):
            with self.subTest(response=response), TemporaryDirectory() as directory:
                root = Path(directory)
                (root/'scripts').mkdir()
                (root/'out/native').mkdir(parents=True)
                (root/'bin').mkdir()
                shutil.copyfile(module.ROOT/'scripts/setup-native.sh', root/'scripts/setup-native.sh')
                shutil.copyfile(module.ROOT/'scripts/run-native-import.py', root/'scripts/run-native-import.py')
                base = root/'out/native/BIPTemplateEditor.base.dotm'
                Packaging().base(base, macro=False)
                original = base.read_bytes()
                for name, body in [('uname', 'echo Darwin'), ('osascript', response)]:
                    executable = root/'bin'/name
                    executable.write_text('#!/bin/sh\n'+body+'\n')
                    executable.chmod(0o755)
                for step in ('build', 'install'):
                    (root/f'scripts/{step}-native.sh').write_text(f'touch {step}.called\n')
                result = subprocess.run(['bash', str(root/'scripts/setup-native.sh')], input='\n', text=True, capture_output=True, env={**os.environ, 'PATH': str(root/'bin')+os.pathsep+os.environ['PATH']})
                confirmed = response == 'echo BIP_VBA_IMPORTED'
                self.assertEqual(result.returncode == 0, confirmed, result.stderr)
                self.assertEqual((root/'build.called').exists(), confirmed)
                self.assertEqual((root/'install.called').exists(), confirmed)
                self.assertEqual(base.read_bytes(), original)
                backups = list(base.parent.glob('*.backup-*'))
                self.assertEqual(len(backups), 1)
                self.assertEqual(backups[0].read_bytes(), original)

if __name__ == '__main__': unittest.main()
