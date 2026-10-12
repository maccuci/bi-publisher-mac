// Preview local de template RTF com a engine do BI Publisher (jars do Oracle Home, nao redistribuidos).
// Uso: java -cp "<Oracle_Home>/bip/lib/*" native/BIPPreview.java template.rtf dados.xml saida.(pdf|html|rtf)
import oracle.xdo.template.FOProcessor;
import oracle.xdo.template.RTFProcessor;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BIPPreview {
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("uso: BIPPreview template.rtf dados.xml saida.(pdf|html|rtf)");
            System.exit(2);
        }
        String out = args[2];
        byte format = out.endsWith(".html") ? FOProcessor.FORMAT_HTML
                : out.endsWith(".rtf") ? FOProcessor.FORMAT_RTF
                : FOProcessor.FORMAT_PDF;

        File xsl = File.createTempFile("bip-preview", ".xsl");
        xsl.deleteOnExit();
        RTFProcessor rtf = new RTFProcessor(args[0]);
        rtf.setOutput(xsl.getPath());
        rtf.process();
        stubCatalogImports(xsl.toPath());

        FOProcessor fo = new FOProcessor();
        fo.setData(args[1]);
        fo.setTemplate(xsl.getPath());
        fo.setOutput(out);
        fo.setOutputFormat(format);
        fo.generate();
    }

    // Subtemplates do catalogo (xdoxsl:///...) so existem no servidor: remove o import e
    // cria templates vazios para cada call-template sem definicao local.
    // ponytail: secoes de subtemplate saem em branco; resolver localmente se o .xsb for exportado.
    static void stubCatalogImports(Path xsl) throws Exception {
        String s = Files.readString(xsl).replaceAll("<xsl:import href=\"xdoxsl:///[^\"]*\"/>", "");
        Set<String> missing = new TreeSet<>();
        Matcher m = Pattern.compile("call-template name=\"([^\"]+)\"").matcher(s);
        while (m.find()) missing.add(m.group(1));
        m = Pattern.compile("<xsl:template name=\"([^\"]+)\"").matcher(s);
        while (m.find()) missing.remove(m.group(1));
        if (missing.isEmpty()) return;
        System.err.println("subtemplates ausentes (em branco): " + missing);
        StringBuilder stubs = new StringBuilder();
        for (String n : missing) stubs.append("<xsl:template name=\"").append(n).append("\"/>");
        int end = s.lastIndexOf("</xsl:stylesheet>");
        Files.writeString(xsl, s.substring(0, end) + stubs + s.substring(end));
    }
}
