"""Attach our Ribbon XML to a VBA template compiled and saved by Word."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
UI_REL = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
ET.register_namespace('', REL)
# xsi:type stores this prefix as text; ElementTree cannot rewrite it.
ET.register_namespace('dcterms', 'http://purl.org/dc/terms/')

def package(source, destination, icons=ROOT / 'out/native/icons'):
    source, destination = Path(source), Path(destination)
    if source.resolve() == destination.resolve():
        raise ValueError('The output must be a separate copy of the base template.')
    with ZipFile(source) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)) or any(n.startswith('/') or '..' in Path(n).parts for n in names):
            raise ValueError('Ambiguous or invalid Word package.')
        if 'word/vbaProject.bin' not in names or not archive.read('word/vbaProject.bin').startswith(bytes.fromhex('d0cf11e0a1b11ae1')):
            raise ValueError('The template must contain VBA compiled by Word; import native/BIPTemplateEditor.bas first.')
        if any('signature' in n.lower() for n in names):
            raise ValueError('Use the unsigned development template; a signature would be invalidated.')
        ct = ET.fromstring(archive.read('[Content_Types].xml'))
        if not any(n.get('PartName') == '/word/document.xml' and n.get('ContentType') == 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml' for n in ct):
            raise ValueError('Save the file as a Word Macro-Enabled Template (.dotm).')
        relationships = ET.fromstring(archive.read('_rels/.rels'))
        for relationship in list(relationships):
            if relationship.get('Type') == UI_REL:
                relationships.remove(relationship)
        ids = {n.get('Id') for n in relationships}
        identity = 'BIPNativeRibbon'
        while identity in ids:
            identity += '_'
        ET.SubElement(relationships, f'{{{REL}}}Relationship', {'Id': identity, 'Type': UI_REL, 'Target': 'customUI/customUI.xml'})
        ribbon = (ROOT / 'native/customUI.xml').read_bytes()
        tree = ET.fromstring(ribbon)
        callbacks = {node.get('onAction') for node in tree.iter() if node.get('onAction')}
        module = (ROOT / 'native/BIPTemplateEditor.bas').read_text()
        if any(f'Public Sub {name}(' not in module for name in callbacks):
            raise ValueError('Some ribbon callbacks have no VBA implementation.')
        images = sorted({node.get('image') for node in tree.iter() if node.get('image')})
        missing = [name for name in images if not (Path(icons) / f'{name}.png').is_file()]
        if missing:
            raise ValueError(f'Missing icons in {icons}: {", ".join(missing)}. Run native:build.')
        image_rels = ET.Element(f'{{{REL}}}Relationships')
        for name in images:
            ET.SubElement(image_rels, f'{{{REL}}}Relationship', {'Id': name, 'Type': IMAGE_REL, 'Target': f'images/{name}.png'})
        content_types = archive.read('[Content_Types].xml')
        if not any(n.get('Extension', '').lower() == 'png' for n in ct):
            # Insercao textual: ElementTree trocaria o namespace padrao por prefixo.
            end = content_types.index(b'>', content_types.index(b'<Types')) + 1
            content_types = content_types[:end] + b'<Default Extension="png" ContentType="image/png"/>' + content_types[end:]
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix('.tmp')
        with ZipFile(temporary, 'w', ZIP_DEFLATED) as output:
            for name in names:
                if name in ('_rels/.rels', '[Content_Types].xml') or name.startswith('customUI/'):
                    continue
                data = archive.read(name)
                if name == 'docProps/core.xml':
                    core = ET.fromstring(data)
                    for item in core:
                        if item.tag.rsplit('}', 1)[-1] in ('creator', 'lastModifiedBy'):
                            item.text = 'BIP Template Editor'
                    data = ET.tostring(core, encoding='utf-8', xml_declaration=True)
                output.writestr(name, data)
            output.writestr('_rels/.rels', ET.tostring(relationships, encoding='utf-8', xml_declaration=True))
            output.writestr('[Content_Types].xml', content_types)
            output.writestr('customUI/customUI.xml', ribbon)
            output.writestr('customUI/_rels/customUI.xml.rels', ET.tostring(image_rels, encoding='utf-8', xml_declaration=True))
            for name in images:
                output.write(Path(icons) / f'{name}.png', f'customUI/images/{name}.png')
        temporary.replace(destination)
    return destination

if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('Usage: python3 scripts/package-native.py base-template.dotm')
    try:
        print(package(sys.argv[1], ROOT / 'out/native/BIPTemplateEditor.dotm'))
    except (ValueError, OSError) as error:
        raise SystemExit(str(error))
