import AppKit

// Ribbon XML nao referencia SF Symbols; renderizamos PNGs no build. Id = atributo image do customUI.xml.
let icons = [
    "iconImportXml": "square.and.arrow.down",
    "iconInsertField": "curlybraces.square",
    "iconForEach": "repeat",
    "iconGroup": "rectangle.stack",
    "iconIf": "arrow.triangle.branch",
    "iconChoose": "list.bullet.indent",
    "iconPreview": "doc.text.magnifyingglass",
    "iconClearXml": "trash",
    "iconInspectField": "text.magnifyingglass",
    "iconAbout": "info.circle",
]
// ponytail: cor fixa; o Word nao inverte imagens customizadas no modo escuro.
let tint = NSColor(srgbRed: 0.18, green: 0.18, blue: 0.2, alpha: 1)
let side = 64, box: CGFloat = 54

guard CommandLine.arguments.count == 2 else { fatalError("Usage: ribbon-icons output-folder") }
let output = URL(fileURLWithPath: CommandLine.arguments[1])
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
for (id, name) in icons {
    guard let symbol = NSImage(systemSymbolName: name, accessibilityDescription: nil)?
        .withSymbolConfiguration(.init(pointSize: box, weight: .regular)) else { fatalError("SF Symbol ausente: \(name)") }
    let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: side, pixelsHigh: side, bitsPerSample: 8, samplesPerPixel: 4,
                               hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
    let scale = min(box / symbol.size.width, box / symbol.size.height)
    let size = NSSize(width: symbol.size.width * scale, height: symbol.size.height * scale)
    let canvas = NSRect(x: 0, y: 0, width: side, height: side)
    symbol.draw(in: NSRect(x: (canvas.width - size.width) / 2, y: (canvas.height - size.height) / 2, width: size.width, height: size.height))
    tint.set()
    canvas.fill(using: .sourceAtop)
    NSGraphicsContext.restoreGraphicsState()
    try rep.representation(using: .png, properties: [:])!.write(to: output.appendingPathComponent("\(id).png"))
}
print("Ribbon icons: \(icons.count) SF Symbols rendered.")
