import Foundation

enum CatalogError: Error { case invalid, unsupported, limit }

final class Catalog: NSObject, XMLParserDelegate {
    var stack: [String] = []
    var paths: [String] = []
    var seen = Set<String>()
    var counts: [String: Int] = [:]
    var samples: [String: String] = [:]
    var texts: [String] = []
    var nodes = 0
    var failure: CatalogError?

    func add(_ path: String, parser: XMLParser) {
        counts[path, default: 0] += 1
        if seen.contains(path) { return }
        if path.utf8.count > 4096 || paths.count >= 10_000 { failure = .limit; parser.abortParsing(); return }
        if seen.insert(path).inserted { paths.append(path) }
    }

    func parser(_ parser: XMLParser, didStartElement name: String, namespaceURI: String?, qualifiedName: String?, attributes: [String: String]) {
        nodes += 1
        if nodes > 200_000 || stack.count >= 128 { failure = .limit; parser.abortParsing(); return }
        stack.append(name)
        texts.append("")
        let path = "/" + stack.joined(separator: "/")
        add(path, parser: parser)
        for attribute in attributes.keys.sorted() where attribute != "xmlns" && !attribute.hasPrefix("xmlns:") {
            add(path + "/@" + attribute, parser: parser)
            sample(path + "/@" + attribute, attributes[attribute]!)
        }
    }

    // Primeiro valor nao vazio de cada caminho (texto direto do elemento), para a inspecao de campos.
    func sample(_ path: String, _ value: String) {
        let clean = value.split(whereSeparator: \.isWhitespace).joined(separator: " ")
        if !clean.isEmpty && samples[path] == nil { samples[path] = String(clean.prefix(80)) }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if !texts.isEmpty { texts[texts.count - 1] += string }
    }

    func parser(_ parser: XMLParser, didEndElement name: String, namespaceURI: String?, qualifiedName: String?) {
        sample("/" + stack.joined(separator: "/"), texts.popLast() ?? "")
        _ = stack.popLast()
    }

    // caminho TAB ocorrencias TAB exemplo
    var lines: [String] { paths.map { "\($0)\t\(counts[$0] ?? 0)\t\(samples[$0] ?? "")" } }
    func parser(_ parser: XMLParser, resolveExternalEntityName name: String, systemID: String?) -> Data? {
        failure = .unsupported; parser.abortParsing(); return nil
    }
}

func catalog(_ data: Data) throws -> Catalog {
    guard data.count <= 10 * 1024 * 1024 else { throw CatalogError.limit }
    guard let xml = String(data: data, encoding: .utf8), !xml.contains("\0") else { throw CatalogError.unsupported }
    let declaration = try NSRegularExpression(pattern: #"<\?xml\s[^?]*\bencoding\s*=\s*["']([^"']+)["']"#, options: [.caseInsensitive])
    if let match = declaration.firstMatch(in: xml, range: NSRange(xml.startIndex..., in: xml)), let range = Range(match.range(at: 1), in: xml), xml[range].lowercased() != "utf-8" {
        throw CatalogError.unsupported
    }
    if xml.range(of: "<!DOCTYPE", options: .caseInsensitive) != nil || xml.range(of: "<!ENTITY", options: .caseInsensitive) != nil { throw CatalogError.unsupported }
    let reader = Catalog()
    let parser = XMLParser(data: data)
    parser.shouldResolveExternalEntities = false
    parser.delegate = reader
    let valid = parser.parse()
    if let failure = reader.failure { throw failure }
    guard valid && reader.stack.isEmpty && !reader.paths.isEmpty else { throw CatalogError.invalid }
    return reader
}

func selfTest() throws {
    let people = try catalog(Data("<DATA><PERSON code='1'><NAME>  João\n da Silva </NAME></PERSON><PERSON code='2'><NAME>Ana</NAME></PERSON></DATA>".utf8))
    precondition(people.paths == ["/DATA", "/DATA/PERSON", "/DATA/PERSON/@code", "/DATA/PERSON/NAME"])
    precondition(people.lines == ["/DATA\t1\t", "/DATA/PERSON\t2\t", "/DATA/PERSON/@code\t2\t1", "/DATA/PERSON/NAME\t2\tJoão da Silva"])
    let namespaced = try catalog(Data("<r xmlns:x='urn:test'><x:NAME/></r>".utf8))
    precondition(namespaced.paths == ["/r", "/r/x:NAME"])
    let bad = ["<broken>", "<!DOCTYPE r [<!ENTITY x SYSTEM 'file:///etc/passwd'>]><r>&x;</r>", String(repeating: "<r>", count: 129) + String(repeating: "</r>", count: 129)]
    for xml in bad {
        do { _ = try catalog(Data(xml.utf8)); preconditionFailure("XML inseguro aceito") } catch is CatalogError {}
    }
    do { _ = try catalog(Data(repeating: 32, count: 10 * 1024 * 1024 + 1)); preconditionFailure("Limite ignorado") } catch is CatalogError {}
    for data in [Data("<?xml version='1.0' encoding='ISO-8859-1'?><r/>".utf8), Data("<r/>".utf16.flatMap { [UInt8($0 & 255), UInt8($0 >> 8)] })] {
        do { _ = try catalog(data); preconditionFailure("Unsupported encoding accepted") } catch is CatalogError {}
    }
    print("Native XML: paths, occurrences, examples, deduplication, attributes, namespaces, Unicode, invalid XML, DTD, and limits OK.")
}

do {
    let arguments = CommandLine.arguments
    if arguments.count == 2 && arguments[1] == "--self-test" {
        try selfTest()
    } else if arguments.count == 3 && arguments[1] == "--xml" {
        let data = try Data(contentsOf: URL(fileURLWithPath: arguments[2]), options: .mappedIfSafe)
        print("BIPXML2\n" + (try catalog(data)).lines.joined(separator: "\n"))
    } else {
        print("Usage: bip-xml --xml file.xml | --self-test")
        exit(1)
    }
} catch CatalogError.limit {
    print("ERR|XML exceeds the limits: 10 MB, 128 levels, 200,000 elements, or 10,000 paths.")
    exit(1)
} catch CatalogError.unsupported {
    print("ERR|Use XML UTF-8 sem DTD ou entidades externas.")
    exit(1)
} catch {
    print("ERR|Could not read a valid XML file. No document was changed.")
    exit(1)
}
