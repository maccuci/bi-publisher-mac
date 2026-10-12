Attribute VB_Name = "BIPTemplateEditor"
Option Explicit

' Pausa sem DoEvents: DoEvents durante o fim de um AppleScriptTask derrubou o Word.
Private Declare PtrSafe Function usleep Lib "/usr/lib/libc.dylib" (ByVal microseconds As Long) As Long

Private xmlPaths As Collection
Private xmlDocument As Document
Private xmlFile As String
' Linhas do leitor XML: caminho TAB ocorrencias TAB exemplo.
Private xmlDetails As Collection

Public Sub BIPImportXml(ByVal control As Object)
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    Dim target As Document
    Set target = ActiveDocument
    Dim response As String
    response = BridgeCall("importXml", "")
    If response = "CANCEL" Then Exit Sub
    If Len(response) = 0 Then response = "ERR|Could not start the native XML reader. Check the installation."
    response = Replace(Replace(response, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String
    lines = Split(response, vbLf)
    If lines(0) <> "BIPXML2" Then
        MsgBox Replace(response, "ERR|", ""), vbExclamation, "Import XML"
        Exit Sub
    End If
    Dim nextPaths As New Collection
    Dim nextDetails As New Collection
    Dim i As Long
    Dim nextFile As String
    For i = 1 To UBound(lines)
        If Left$(lines(i), 5) = "FILE|" Then
            nextFile = Mid$(lines(i), 6)
        ElseIf Len(lines(i)) > 0 Then
            nextDetails.Add lines(i)
            nextPaths.Add Split(lines(i), vbTab)(0)
        End If
    Next i
    If nextPaths.Count = 0 Then Err.Raise 5
    Set xmlPaths = nextPaths
    Set xmlDetails = nextDetails
    Set xmlDocument = target
    xmlFile = nextFile
    MsgBox CStr(xmlPaths.Count) & " XML paths loaded for this document." & vbCr & _
        "Use Insert Field to add a field at the cursor." & vbCr & _
        "The XML is kept for this session only.", vbInformation, "BIP Template Editor"
    Exit Sub
Failed:
    MsgBox "Could not import the XML. Check the BIPBridge and native XML reader installation.", vbExclamation, "BIP Template Editor"
End Sub

Public Sub BIPBrowseXml(ByVal control As Object)
    On Error GoTo Failed
    Dim chosen As String
    chosen = ChooseXmlPath("Select an XML field:")
    If Len(chosen) = 0 Then Exit Sub
    ' Word ignora sem erro um campo dentro de outro: sobre um campo existente, insere logo depois dele.
    Dim target As Range
    Set target = Selection.Range
    Dim existing As Field
    For Each existing In ActiveDocument.Fields
        If target.Start <= existing.Result.End + 1 And target.End >= existing.Code.Start - 1 Then
            target.SetRange existing.Result.End + 1, existing.Result.End + 1
        End If
    Next existing
    ' Campo legado BI Publisher: texto exibido = nome do elemento, expressao no Help Text (F1, max. 255).
    InsertBipField target, "<?" & RelativePath(chosen, target.Start) & "?>", Mid$(chosen, InStrRev(chosen, "/") + 1)
    Exit Sub
Failed:
    MsgBox "Could not insert the field (the document may be protected).", vbExclamation, "Insert Field"
End Sub

' Repeticao: F antes e E depois da selecao. Em tabela, repete as linhas selecionadas (F na 1a celula, E na ultima).
' Dentro de um for-each-group, oferece repetir os itens do grupo (current-group()).
Public Sub BIPInsertForEach(ByVal control As Object)
    On Error GoTo Failed
    Dim expression As String
    If InsideBlock(Selection.Start, "for-each-group") Then
        Select Case MsgBox("Repeat the items of the current group (current-group())?" & vbCr & "No = choose another XML element.", vbYesNoCancel + vbQuestion, "Repeat")
            Case vbYes: expression = "current-group()"
            Case vbCancel: Exit Sub
        End Select
    End If
    If Len(expression) = 0 Then
        Dim chosen As String
        chosen = ChooseXmlPath("Select the repeating element:")
        If Len(chosen) = 0 Then Exit Sub
        expression = RelativePath(chosen, Selection.Start)
    End If
    WrapSelection Array("<?for-each:" & expression & "?>", "F"), Array("<?end for-each?>", "E"), True
    Exit Sub
Failed:
    MsgBox "Could not insert the loop (the document may be protected).", vbExclamation, "Repeat"
End Sub

' Agrupamento: G antes e EG depois; dentro dele, Repeat oferece os itens do grupo.
Public Sub BIPInsertGroup(ByVal control As Object)
    On Error GoTo Failed
    Dim element As String
    element = ChooseXmlPath("Select the element to group:")
    If Len(element) = 0 Then Exit Sub
    Dim keys As New Collection
    Dim item As Variant
    For Each item In xmlPaths
        If Left$(CStr(item), Len(element) + 1) = element & "/" Then keys.Add CStr(item)
    Next item
    If keys.Count = 0 Then
        MsgBox "The selected element has no fields to use as a grouping key.", vbInformation, "Group"
        Exit Sub
    End If
    Dim key As String
    key = ChooseOption("Group by which field?", keys)
    If Len(key) = 0 Then Exit Sub
    WrapSelection Array("<?for-each-group:" & RelativePath(element, Selection.Start) & ";./" & Mid$(key, Len(element) + 2) & "?>", "G"), _
        Array("<?end for-each-group?>", "EG"), True
    Exit Sub
Failed:
    MsgBox "Could not insert the group (the document may be protected).", vbExclamation, "Group"
End Sub

' Condicao: C antes e EC depois da selecao; se a selecao cobre mais de uma celula, condiciona as linhas.
Public Sub BIPInsertIf(ByVal control As Object)
    On Error GoTo Failed
    Dim condition As String
    condition = AskCondition("Condition")
    If Len(condition) = 0 Then Exit Sub
    WrapSelection Array("<?if:" & condition & "?>", "C"), Array("<?end if?>", "EC"), SeveralCells()
    Exit Sub
Failed:
    MsgBox "Could not insert the condition (the document may be protected).", vbExclamation, "Condition"
End Sub

' Escolha: fora de um choose (ou dentro de um caso), cria CH + 1o caso W em volta da selecao.
' Entre os casos de um choose aberto, adiciona outro when ou o otherwise.
Public Sub BIPInsertChoose(ByVal control As Object)
    On Error GoTo Failed
    Dim position As Long, condition As String
    position = Selection.Start
    Dim outer As Collection
    Set outer = Blocks(position)
    Dim betweenCases As Boolean
    If outer.Count > 0 Then betweenCases = TagKind(outer(outer.Count)) = "choose"
    If betweenCases Then
        Dim actions As New Collection
        actions.Add "Add case (when)"
        actions.Add "Otherwise (otherwise)"
        Select Case ChooseOption("Add to choose:", actions)
            Case "Add case (when)"
                condition = AskCondition("Case (when)")
                If Len(condition) = 0 Then Exit Sub
                WrapSelection Array("<?when:" & condition & "?>", "W"), Array("<?end when?>", "EW"), SeveralCells()
            Case "Otherwise (otherwise)"
                WrapSelection Array("<?otherwise:?>", "O"), Array("<?end otherwise?>", "EO"), SeveralCells()
        End Select
        Exit Sub
    End If
    condition = AskCondition("First case (when)")
    If Len(condition) = 0 Then Exit Sub
    If Not WrapSelection(Array("<?choose:?>", "CH", "<?when:" & condition & "?>", "W"), _
        Array("<?end when?>", "EW", "<?end choose?>", "ECH"), SeveralCells()) Then Exit Sub
    MsgBox "To add another case or the otherwise, select the text right after EW" & vbCr & _
        "(before ECH) and click Choose again.", vbInformation, "Choose"
    Exit Sub
Failed:
    MsgBox "Could not insert the choose block (the document may be protected).", vbExclamation, "Choose"
End Sub

' Campo do XML + condicao XPath editavel; vazio se cancelado.
Private Function AskCondition(ByVal title As String) As String
    Dim chosen As String
    chosen = ChooseXmlPath("Select the condition field:")
    If Len(chosen) = 0 Then Exit Function
    AskCondition = Trim$(InputBox("XPath condition (e.g. STATUS='ACTIVE', number(AMOUNT)>0, STATUS!=''):", _
        title, RelativePath(chosen, Selection.Start) & "=''"))
End Function

' Selection.Cells falha fora de tabela.
Private Function SeveralCells() As Boolean
    If Selection.Information(wdWithInTable) Then SeveralCells = Selection.Cells.Count > 1
End Function

' Preview com a engine Oracle local (native/BIPPreview.java); o bridge abre o PDF quando termina.
Public Sub BIPPreview(ByVal control As Object)
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    If xmlDocument Is Nothing Or Len(xmlFile) = 0 Then GoTo Missing
    If Not (xmlDocument Is ActiveDocument) Then GoTo Missing
    If ActiveDocument.SaveFormat <> wdFormatRTF Then
        MsgBox "Preview uses the file saved as RTF. Save the template as RTF (File > Save As) and try again.", vbInformation, "Preview"
        Exit Sub
    End If
    If Not ActiveDocument.Saved Then
        If MsgBox("Preview uses the saved file. Save the document now?", vbOKCancel + vbQuestion, "Preview") <> vbOK Then Exit Sub
        ActiveDocument.Save
    End If
    Dim response As String
    response = BridgeCall("preview", ActiveDocument.FullName & vbLf & xmlFile)
    If response <> "OK" Then MsgBox Replace(response, "ERR|", ""), vbExclamation, "Preview"
    Exit Sub
Missing:
    MsgBox "Import an XML file for this document first.", vbInformation, "BIP Template Editor"
    Exit Sub
Failed:
    MsgBox "Could not start the preview. Check the BIPBridge installation.", vbExclamation, "Preview"
End Sub

' Caminho escolhido no catalogo do documento ativo; vazio se cancelado ou sem XML.
Private Function ChooseXmlPath(ByVal prompt As String) As String
    If Documents.Count = 0 Or xmlDocument Is Nothing Then GoTo Missing
    If Not (xmlDocument Is ActiveDocument) Then GoTo Missing
    ChooseXmlPath = ChooseOption(prompt, xmlPaths)
    Exit Function
Missing:
    MsgBox "Import an XML file for this document first.", vbInformation, "BIP Template Editor"
End Function

' Lista nativa de escolha (bridge); vazio se cancelado. Recusa resposta fora da lista.
Private Function ChooseOption(ByVal prompt As String, ByVal options As Collection) As String
    Dim lines As String
    Dim item As Variant
    For Each item In options
        lines = lines & CStr(item) & vbLf
    Next item
    Dim chosen As String
    chosen = BridgeCall("choosePath", prompt & vbLf & lines)
    If chosen = "CANCEL" Or Len(chosen) = 0 Then Exit Function
    For Each item In options
        If CStr(item) = chosen Then ChooseOption = chosen
    Next item
    If Len(ChooseOption) = 0 Then Err.Raise 5
End Function

' Expressao BI Publisher do campo com referencias <?ref:...?> expandidas.
Private Function FieldExpression(ByVal field As FormField) As String
    FieldExpression = ExpandRefs(RawExpression(field))
End Function

' Help Text proprio ou, como o Template Builder grava, texto de status proprio.
Private Function RawExpression(ByVal field As FormField) As String
    If field.OwnHelp Then RawExpression = field.HelpText
    If Len(RawExpression) = 0 And field.OwnStatus Then RawExpression = field.StatusText
End Function

' O Template Builder grava <?ref:xdoNNNN?> no campo e a expressao em base64 numa variavel do documento.
Private Function ExpandRefs(ByVal expression As String) As String
    Dim startAt As Long, finish As Long, value As String
    ExpandRefs = expression
    startAt = InStr(ExpandRefs, "<?ref:")
    Do While startAt > 0
        finish = InStr(startAt, ExpandRefs, "?>")
        If finish = 0 Then Exit Do
        value = DocumentVariable(Mid$(ExpandRefs, startAt + 6, finish - startAt - 6))
        If Len(value) = 0 Then
            startAt = InStr(finish, ExpandRefs, "<?ref:")
        Else
            ExpandRefs = Left$(ExpandRefs, startAt - 1) & value & Mid$(ExpandRefs, finish + 2)
            startAt = InStr(startAt + Len(value), ExpandRefs, "<?ref:")
        End If
    Loop
End Function

Private Function DocumentVariable(ByVal name As String) As String
    On Error Resume Next
    DocumentVariable = Utf8Decode(Base64Decode(ActiveDocument.Variables(name).Value))
End Function

' Instrucoes <?...?> de uma expressao (o Template Builder pode juntar varias num campo).
Private Function Tags(ByVal expression As String) As Collection
    Set Tags = New Collection
    Dim startAt As Long, finish As Long
    startAt = InStr(expression, "<?")
    Do While startAt > 0
        finish = InStr(startAt, expression, "?>")
        If finish = 0 Then Exit Do
        Tags.Add Mid$(expression, startAt, finish - startAt + 2)
        startAt = InStr(finish + 2, expression, "<?")
    Loop
End Function

' Palavra-chave da instrucao ("for-each", "if", "end if"...), sem sufixo @section/@row; "" para campo de dados.
Private Function TagKind(ByVal tag As String) As String
    Dim body As String
    body = Trim$(Mid$(tag, 3, Len(tag) - 4))
    If InStr(body, ":") = 0 Then
        If Left$(body, 4) = "end " Then TagKind = body
        Exit Function
    End If
    TagKind = Left$(body, InStr(body, ":") - 1)
    If InStr(TagKind, "@") > 0 Then TagKind = Left$(TagKind, InStr(TagKind, "@") - 1)
End Function

Private Function TagArgument(ByVal tag As String) As String
    Dim body As String
    body = Mid$(tag, 3, Len(tag) - 4)
    If InStr(body, ":") > 0 Then TagArgument = Trim$(Mid$(body, InStr(body, ":") + 1))
End Function

' Instrucoes de bloco abertas antes de position, da mais externa para a mais interna.
Private Function Blocks(ByVal position As Long) As Collection
    Set Blocks = New Collection
    Dim field As FormField, tag As Variant, kind As String
    For Each field In ActiveDocument.FormFields
        If field.Range.End > position Then Exit For
        For Each tag In Tags(FieldExpression(field))
            kind = TagKind(CStr(tag))
            Select Case kind
                Case "for-each", "for-each-group", "if", "choose", "when", "otherwise"
                    Blocks.Add CStr(tag)
                Case Else
                    If Left$(kind, 4) = "end " And Blocks.Count > 0 Then Blocks.Remove Blocks.Count
            End Select
        Next tag
    Next field
End Function

Private Function InsideBlock(ByVal position As Long, ByVal kind As String) As Boolean
    Dim tag As Variant
    For Each tag In Blocks(position)
        If TagKind(CStr(tag)) = kind Then InsideBlock = True
    Next tag
End Function

' Caminho XML absoluto do contexto dos lacos abertos (for-each/for-each-group).
Private Function ContextPath(ByVal outer As Collection) As String
    Dim tag As Variant, loopPath As String
    For Each tag In outer
        Select Case TagKind(CStr(tag))
            Case "for-each", "for-each-group"
                loopPath = TagArgument(CStr(tag))
                If InStr(loopPath, ";") > 0 Then loopPath = Left$(loopPath, InStr(loopPath, ";") - 1)
                ' Os itens de current-group() sao o mesmo elemento do grupo.
                If loopPath <> "current-group()" Then ContextPath = ResolveXmlPath(loopPath, ContextPath)
        End Select
    Next tag
End Function

' Caminho absoluto de uma referencia: absoluta, ou como o Oracle (.//NOME) abaixo do contexto, pelo XML importado.
' Sem correspondencia, junta o texto ao contexto.
Private Function ResolveXmlPath(ByVal path As String, ByVal context As String) As String
    If Left$(path, 1) = "/" Then ResolveXmlPath = path: Exit Function
    Dim matches As Collection
    Set matches = XmlMatches(path, context)
    If matches.Count > 0 Then
        ResolveXmlPath = matches(1)
    ElseIf Len(context) > 0 Then
        ResolveXmlPath = context & "/" & path
    Else
        ResolveXmlPath = path
    End If
End Function

' Caminhos do XML importado que terminam em path, abaixo do contexto.
Private Function XmlMatches(ByVal path As String, ByVal context As String) As Collection
    Set XmlMatches = New Collection
    If Not CatalogReady() Then Exit Function
    If Left$(path, 2) = "./" Then path = Mid$(path, 3)
    Dim item As Variant
    For Each item In xmlPaths
        If Right$(item, Len(path) + 1) = "/" & path Then
            If Len(context) = 0 Or Left$(item, Len(context) + 1) = context & "/" Then XmlMatches.Add CStr(item)
        End If
    Next item
End Function

Private Function CatalogReady() As Boolean
    If xmlPaths Is Nothing Or xmlDocument Is Nothing Then Exit Function
    CatalogReady = xmlDocument Is ActiveDocument
End Function

' Caminho relativo ao laco aberto antes de position (o Oracle avalia campos no contexto do laco).
Private Function RelativePath(ByVal path As String, ByVal position As Long) As String
    Dim context As String
    context = ContextPath(Blocks(position))
    RelativePath = path
    If Len(context) > 0 And Left$(path, Len(context) + 1) = context & "/" Then RelativePath = Mid$(path, Len(context) + 2)
End Function

' Envolve a selecao com campos: opens/closes = Array(expressao, rotulo, ...), na ordem do documento.
' Insere do fim para o inicio, cada um na mesma posicao, para nao deslocar as posicoes calculadas.
Private Function WrapSelection(ByVal opens As Variant, ByVal closes As Variant, ByVal wholeRows As Boolean) As Boolean
    Dim startAt As Long, endAt As Long, i As Long
    If wholeRows And Selection.Information(wdWithInTable) Then
        startAt = Selection.Rows(1).Range.Cells(1).Range.Start
        With Selection.Rows(Selection.Rows.Count).Range.Cells
            endAt = .Item(.Count).Range.End - 1
        End With
    Else
        startAt = OutsideField(Selection.Start, False)
        endAt = OutsideField(Selection.End, True)
    End If
    For i = UBound(closes) - 1 To 0 Step -2
        If Not InsertBipField(ActiveDocument.Range(endAt, endAt), closes(i), closes(i + 1)) Then Exit Function
    Next i
    For i = UBound(opens) - 1 To 0 Step -2
        If Not InsertBipField(ActiveDocument.Range(startAt, startAt), opens(i), opens(i + 1)) Then Exit Function
    Next i
    WrapSelection = True
End Function

' Posicao fora de qualquer campo: antes dele (inicio) ou depois dele (fim).
Private Function OutsideField(ByVal position As Long, ByVal after As Boolean) As Long
    OutsideField = position
    Dim existing As Field
    For Each existing In ActiveDocument.Fields
        If position > existing.Code.Start - 1 And position <= existing.Result.End Then
            If after Then OutsideField = existing.Result.End + 1 Else OutsideField = existing.Code.Start - 1
        End If
    Next existing
End Function

Private Function InsertBipField(ByVal target As Range, ByVal expression As String, ByVal label As String) As Boolean
    If Len(expression) > 255 Then
        MsgBox "Expression too long for Help Text (255 characters). Nothing was inserted.", vbExclamation, "BIP Template Editor"
        Exit Function
    End If
    ' FormFields.Add no Word Mac cria um campo vazio que nao aceita texto; um trecho RTF
    ' (mesmo formato do Word/Oracle) grava texto exibido e Help Text de uma vez.
    Dim snippet As String
    snippet = Environ("TMPDIR") & "/bip-field.rtf"
    Dim handle As Integer
    handle = FreeFile
    Open snippet For Output As #handle
    Print #handle, "{\rtf1\ansi\deff0{\field{\*\fldinst FORMTEXT {\*\formfield\fftype0\ffownhelp1{\*\ffhelptext " & _
        RtfText(expression) & "}{\*\ffdeftext " & RtfText(label) & "}}}{\fldrslt " & RtfText(label) & "}}}"
    Close #handle
    Dim before As Long
    before = ActiveDocument.FormFields.Count
    target.InsertFile snippet
    Kill snippet
    If ActiveDocument.FormFields.Count <> before + 1 Then
        MsgBox "Word did not insert the field at this position. Place the cursor outside other fields and try again.", vbExclamation, "BIP Template Editor"
        Exit Function
    End If
    ' Destaque cinza dos campos, como no plugin Oracle (configuracao de exibicao do documento).
    ActiveDocument.FormFields.Shaded = True
    InsertBipField = True
End Function

' AppleScriptTask volta quando o dialogo do script fecha, antes de o script terminar, e o retorno
' chega vazio. O bridge grava o resultado em arquivo (renomeado ao final) e esperamos por ele.
Private Function BridgeCall(ByVal handler As String, ByVal payload As String) As String
    Dim folder As String
    folder = Environ("TMPDIR")
    If Right$(folder, 1) <> "/" Then folder = folder & "/"
    Dim resultPath As String
    resultPath = folder & "bip-result.txt"
    On Error Resume Next
    Kill resultPath
    Kill resultPath & ".part"
    On Error GoTo 0
    Dim ignored As String
    ignored = AppleScriptTask("BIPBridge.applescript", handler, resultPath & vbLf & payload)
    Dim started As Single
    started = Timer
    ' Dir, nao FileLen: no Word Mac FileLen de arquivo inexistente devolve 0 sem erro.
    Do While Len(Dir(resultPath)) = 0
        ' ponytail: espera fixa de 10 s (Timer zera a meia-noite); o script so grava e renomeia apos o dialogo.
        If Timer - started > 10 Or Timer < started Then Err.Raise 53
        usleep 50000
    Loop
    BridgeCall = ReadUtf8(resultPath)
    Kill resultPath
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim handle As Integer
    handle = FreeFile
    Open path For Binary Access Read As #handle
    If LOF(handle) = 0 Then
        Close #handle
        Exit Function
    End If
    Dim bytes() As Byte
    ReDim bytes(0 To LOF(handle) - 1)
    Get #handle, , bytes
    Close #handle
    ReadUtf8 = Utf8Decode(bytes)
End Function

Private Function Utf8Decode(ByRef bytes() As Byte) As String
    Dim buffer As String
    buffer = Space$(UBound(bytes) + 1)
    Dim i As Long, used As Long, code As Long, extra As Long, k As Long
    Do While i <= UBound(bytes)
        code = bytes(i)
        If code >= &HF0 Then
            code = code And &H7: extra = 3
        ElseIf code >= &HE0 Then
            code = code And &HF: extra = 2
        ElseIf code >= &HC0 Then
            code = code And &H1F: extra = 1
        Else
            extra = 0
        End If
        For k = 1 To extra
            code = code * 64 + (bytes(i + k) And &H3F)
        Next k
        i = i + extra + 1
        If code > &HFFFF& Then
            code = code - &H10000
            used = used + 1: Mid$(buffer, used, 1) = ChrW(&HD800& + code \ 1024)
            used = used + 1: Mid$(buffer, used, 1) = ChrW(&HDC00& + (code And &H3FF))
        Else
            used = used + 1: Mid$(buffer, used, 1) = ChrW(code)
        End If
    Loop
    Utf8Decode = Left$(buffer, used)
End Function

Private Function Base64Decode(ByVal text As String) As Byte()
    Const alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim bytes() As Byte, used As Long, buffer As Long, bits As Long, i As Long, value As Long
    ReDim bytes(0 To Len(text))
    For i = 1 To Len(text)
        value = InStr(1, alphabet, Mid$(text, i, 1), vbBinaryCompare) - 1
        If value >= 0 Then
            buffer = (buffer * 64 + value) And &HFFFFFF
            bits = bits + 6
            If bits >= 8 Then
                bits = bits - 8
                bytes(used) = (buffer \ (2 ^ bits)) And 255
                used = used + 1
            End If
        End If
    Next i
    If used = 0 Then Err.Raise 5
    ReDim Preserve bytes(0 To used - 1)
    Base64Decode = bytes
End Function

' Texto RTF em ASCII: escapa \ { } e usa \uN? para caracteres fora do ASCII.
Private Function RtfText(ByVal value As String) As String
    Dim i As Long, code As Long, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)
        If ch = "\" Or ch = "{" Or ch = "}" Then
            RtfText = RtfText & "\" & ch
        ElseIf code < 32 Or code > 126 Then
            RtfText = RtfText & "\u" & CStr(code) & "?"
        Else
            RtfText = RtfText & ch
        End If
    Next i
End Function

Public Sub BIPClearXml(ByVal control As Object)
    Set xmlPaths = Nothing
    Set xmlDetails = Nothing
    Set xmlDocument = Nothing
    xmlFile = ""
    MsgBox "Sample XML cleared from the session.", vbInformation, "BIP Template Editor"
End Sub

' Como as propriedades do Template Builder no Windows: com o cursor num campo, mostra tipo, expressao
' (expandindo <?ref:...?>), origem no XML importado com ocorrencias e exemplo, e os blocos que o envolvem.
Public Sub BIPInspectField(ByVal control As Object)
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    Dim field As FormField
    Set field = FieldAtSelection()
    If field Is Nothing Then
        MsgBox "Place the cursor inside a template field (or select it) and click Inspect Field.", vbInformation, "Inspect Field"
        Exit Sub
    End If
    Dim raw As String
    raw = RawExpression(field)
    If Len(raw) = 0 Then
        MsgBox "Field: " & field.Name & vbCr & "Display text: " & field.Result & vbCr & _
            "No own Help Text or status text: BI Publisher does not read this field.", vbInformation, "Inspect Field"
        Exit Sub
    End If
    Dim expression As String
    expression = ExpandRefs(raw)
    Dim outer As Collection
    Set outer = Blocks(field.Range.Start)
    Dim report As String
    report = "Field: " & field.Name & "    Display text: " & field.Result & vbCr & _
        "Type: " & TagDescription(expression) & vbCr & vbCr & _
        "Expression: " & expression & vbCr
    If raw <> expression Then report = report & "Stored as: " & raw & " (document variable)" & vbCr
    report = report & vbCr & "XML source: " & OriginReport(expression, ContextPath(outer)) & vbCr & vbCr & _
        "Inside: " & ContextReport(outer)
    MsgBox report, vbInformation, "Inspect Field"
    Exit Sub
Failed:
    MsgBox "Could not read this field. No changes were made.", vbExclamation, "Inspect Field"
End Sub

' Campo selecionado ou sob o cursor.
Private Function FieldAtSelection() As FormField
    If Selection.FormFields.Count > 0 Then Set FieldAtSelection = Selection.FormFields(1): Exit Function
    Dim field As FormField
    For Each field In ActiveDocument.FormFields
        If field.Range.Start > Selection.End Then Exit For
        If Selection.Start >= field.Range.Start And Selection.Start <= field.Range.End Then Set FieldAtSelection = field: Exit Function
    Next field
End Function

Private Function TagDescription(ByVal expression As String) As String
    Dim all As Collection
    Set all = Tags(expression)
    If all.Count = 0 Then TagDescription = "Text without a <?...?> instruction": Exit Function
    Dim kind As String
    kind = TagKind(all(1))
    Select Case kind
        Case "": TagDescription = "Data field"
        Case "for-each": TagDescription = "Loop (for-each)"
        Case "for-each-group": TagDescription = "Group (for-each-group)"
        Case "if": TagDescription = "Condition (if)"
        Case "choose": TagDescription = "Choose (choose)"
        Case "when": TagDescription = "Case (when)"
        Case "otherwise": TagDescription = "Otherwise (otherwise)"
        Case "call-template", "call": TagDescription = "Subtemplate call"
        Case "ref": TagDescription = "Reference without a document variable"
        Case Else
            If Left$(kind, 4) = "end " Then
                TagDescription = "End of block (" & Mid$(kind, 5) & ")"
            Else
                TagDescription = "BI Publisher instruction (" & kind & ")"
            End If
    End Select
    If all.Count > 1 Then TagDescription = TagDescription & " + " & CStr(all.Count - 1) & " more instruction(s)"
End Function

' Caminho no XML importado, ocorrencias e exemplo; so para campos de dados e lacos com caminho direto.
Private Function OriginReport(ByVal expression As String, ByVal context As String) As String
    Dim all As Collection, path As String
    Set all = Tags(expression)
    If all.Count = 1 Then
        Select Case TagKind(all(1))
            Case "": path = Trim$(Mid$(all(1), 3, Len(all(1)) - 4))
            Case "for-each", "for-each-group": path = TagArgument(all(1))
        End Select
        If InStr(path, ";") > 0 Then path = Left$(path, InStr(path, ";") - 1)
    End If
    If path = "current-group()" Then OriginReport = "items of the current group (" & context & ")": Exit Function
    If Len(path) = 0 Or path Like "*[ ()'=<>!$,""]*" Or InStr(path, "..") > 0 Then
        OriginReport = "instruction or calculated expression, no direct XML path": Exit Function
    End If
    If Not CatalogReady() Then OriginReport = path & " (import an XML file to see occurrences and an example)": Exit Function
    Dim matches As Collection
    If Left$(path, 1) = "/" Then
        Set matches = New Collection
        If Len(PathDetails(path)) > 0 Then matches.Add path
    Else
        Set matches = XmlMatches(path, context)
    End If
    If matches.Count = 0 Then OriginReport = path & " - NOT found in the imported XML": Exit Function
    Dim parts() As String
    parts = Split(PathDetails(matches(1)), vbTab)
    OriginReport = matches(1) & vbCr & "    " & parts(1) & " occurrence(s) in the XML"
    If UBound(parts) >= 2 Then
        If Len(parts(2)) > 0 Then OriginReport = OriginReport & ", example: " & parts(2)
    End If
    If matches.Count > 1 Then OriginReport = OriginReport & vbCr & "    (+" & CStr(matches.Count - 1) & " other path(s) with this name; Oracle uses the first)"
End Function

Private Function PathDetails(ByVal path As String) As String
    Dim item As Variant
    For Each item In xmlDetails
        If Left$(item, Len(path) + 1) = path & vbTab Then PathDetails = item: Exit Function
    Next item
End Function

Private Function ContextReport(ByVal outer As Collection) As String
    Dim tag As Variant
    For Each tag In outer
        If Len(ContextReport) > 0 Then ContextReport = ContextReport & " > "
        ContextReport = ContextReport & Trim$(Mid$(tag, 3, Len(tag) - 4))
    Next tag
    If Len(ContextReport) = 0 Then ContextReport = "document top level (outside loops and conditions)"
End Function

Public Sub BIPAbout(ByVal control As Object)
    MsgBox "BIP Template Editor - native prototype for Word for Mac." & vbCr & _
        "Ribbon tab, local XML import, fields, for-each, for-each-group, if, choose, and preview" & vbCr & _
        "with the locally installed Oracle engine (Oracle_Home).", vbInformation, "BIP Template Editor"
End Sub
