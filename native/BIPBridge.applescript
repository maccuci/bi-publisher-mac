-- AppleScriptTask volta ao VBA quando o dialogo fecha, antes de o script terminar, e o retorno
-- se perde. O resultado vai para o arquivo informado pelo VBA na 1a linha do parametro:
-- grava em .part e renomeia (atomico), e o VBA espera o arquivo final.
on resultFile(payload)
    set target to paragraph 1 of payload
    set allowed to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Word/Data/tmp/"
    if target does not start with allowed or target contains "/../" then error "Invalid result file."
    return target
end resultFile

on reply(target, value)
    do shell script "printf '%s' " & quoted form of value & " > " & quoted form of (target & ".part") & "; mv -f " & quoted form of (target & ".part") & " " & quoted form of target
    return value
end reply

on importXml(payload)
    set target to my resultFile(payload)
    -- Seletor do AppleScript: o NSOpenPanel do helper de linha de comando perdia o foco e travava.
    try
        set xmlFile to choose file with prompt "Import sample XML - BIP Template Editor" of type {"public.xml"}
    on error number -128
        return my reply(target, "CANCEL")
    end try
    set helper to (POSIX path of (path to home folder)) & "Library/Application Scripts/com.microsoft.Word/bip-xml"
    -- Em XML invalido o helper sai com 1; a mensagem ERR| gravada chega ao VBA pelo mv.
    -- Com sucesso, a linha FILE| informa o arquivo ao VBA (usado pelo preview).
    set xmlPath to POSIX path of xmlFile
    do shell script quoted form of helper & " --xml " & quoted form of xmlPath & " > " & quoted form of (target & ".part") & " && printf '\\nFILE|%s\\n' " & quoted form of xmlPath & " >> " & quoted form of (target & ".part") & "; mv -f " & quoted form of (target & ".part") & " " & quoted form of target
    return "OK"
end importXml

on choosePath(payload)
    set target to my resultFile(payload)
    set options to {}
    repeat with candidate in items 3 thru -1 of paragraphs of payload
        if (contents of candidate) is not "" then set end of options to (contents of candidate)
    end repeat
    set chosen to choose from list options with title "BIP Template Editor" with prompt (paragraph 2 of payload) without multiple selections allowed
    if chosen is false then return my reply(target, "CANCEL")
    return my reply(target, item 1 of chosen)
end choosePath

-- Preview com a engine Oracle local. Roda em segundo plano (o VBA espera so 10 s e a engine leva alguns
-- segundos): abre o PDF ao terminar, ou o log de erro.
-- Os jars da engine ficam numa pasta local, copiados uma vez por `bun run oracle:import` (nao dependem do Oracle Home).
-- Java 21: Homebrew no Apple Silicon (/opt/homebrew) ou Intel (/usr/local), ou um JDK registrado no macOS.
on preview(payload)
    set target to my resultFile(payload)
    set rtfPath to paragraph 2 of payload
    set xmlPath to paragraph 3 of payload
    set oracleLib to (POSIX path of (path to home folder)) & "Library/Application Support/BIPTemplateEditor/oracle-lib"
    set engine to (POSIX path of (path to home folder)) & "Library/Application Scripts/com.microsoft.Word/BIPPreview.java"
    try
        set javaBin to do shell script "for j in /opt/homebrew/opt/openjdk@21/bin/java /usr/local/opt/openjdk@21/bin/java \"$(/usr/libexec/java_home -v 21 2>/dev/null)/bin/java\"; do [ -x \"$j\" ] && { echo \"$j\"; exit 0; }; done; exit 1"
    on error
        return my reply(target, "ERR|Java 21 not found. Install it with: brew install openjdk@21")
    end try
    try
        do shell script "test -f " & quoted form of (oracleLib & "/xdo-core.jar")
    on error
        return my reply(target, "ERR|Oracle engine not found in " & oracleLib & ". Run \"Import Oracle Engine.command\" once and choose your Oracle_Home folder.")
    end try
    set outDir to do shell script "mktemp -d -t bip-preview"
    set pdf to outDir & "/preview.pdf"
    set logFile to outDir & "/preview.log"
    set javaCommand to quoted form of javaBin & " -cp " & quoted form of (oracleLib & "/*") & " " & quoted form of engine & " " & quoted form of rtfPath & " " & quoted form of xmlPath & " " & quoted form of pdf
    do shell script "(" & javaCommand & " > " & quoted form of logFile & " 2>&1 && open " & quoted form of pdf & " || open -e " & quoted form of logFile & ") > /dev/null 2>&1 &"
    return my reply(target, "OK")
end preview
