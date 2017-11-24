Attribute VB_Name = "execucaoLeituraDocumento"
Public arrayArquivo() As String

Dim appWord As Word.Application
Dim doc As Word.Document
Dim prg As Paragraph
Dim wdrange As Word.Range

Dim paragrafoLocalizado As String
Dim codigoNCM As String

Dim i As Integer


Private Function trataCaracterTextoLidoDoc(texto As String) As String
    
    texto = Replace(texto, Chr(13), Chr(32))
    texto = Replace(texto, Chr(9), "")
    texto = Replace(texto, Chr(10), "")
    texto = Replace(texto, Chr(16), "")
    texto = Replace(texto, Chr(Asc(Mid(texto, Len(texto), 1))), "")
    
    trataCaracterTextoLidoDoc = texto
    
End Function

Public Sub LerPorDocumento(strCaminho As String)

    Dim posicaoNumeroLinha As Integer

    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set doc = appWord.Documents.Open(strCaminho)

    Set wdrange = doc.Range
    
    wdrange.ParagraphFormat.SpaceAfter = 0
    wdrange.ParagraphFormat.SpaceBefore = 0
    wdrange.ParagraphFormat.SpaceAfterAuto = 0
    wdrange.ParagraphFormat.SpaceBeforeAuto = 0
    
    wdrange.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast
    wdrange.ParagraphFormat.LineSpacing = 0.7
    
    i = 3
    
    '---------------------------------------------------------------------
    'Planilha NCM
    '---------------------------------------------------------------------
    Worksheets("NCM").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    
    'Descrição do Segmento
    i = i + 2
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C" + CStr(posicaoNumeroLinha)).Value = paragrafoLocalizado
    
    'NCM
    i = i + 6
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("A" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    codigoNCM = Trim(paragrafoLocalizado)
    
    'Descrição NCM
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
    Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Código do CEST
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
    Range("D" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)

    '---------------------------------------------------------------------
    'Planilha Base Legal
    '---------------------------------------------------------------------
    Worksheets("BASE_LEGAL").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
    
    'UF
    i = i + 3
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Descrição da Base Legal
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("BASE_LEGAL").Activate
    Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Base de Cálculo
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D" + CStr(posicaoNumeroLinha)).Value = IIf(Trim(paragrafoLocalizado) = "-", 0, Trim(paragrafoLocalizado))
    
    'Início da Vigência
    i = i + 5
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Fim da Vigência
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("F" + CStr(posicaoNumeroLinha)).Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
    
    '---------------------------------------------------------------------
    'Planilha MVA Original
    '---------------------------------------------------------------------
    Worksheets("ALIQUOTAS_MVA").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
    
    'MVA Original
    i = i + 8
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'MVA Ajustada 4%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'MVA Ajustada 12%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Alíquota Interna
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    '---------------------------------------------------------------------
    'IPI
    '---------------------------------------------------------------------
    Worksheets("IPI").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    
    'Descrição
    i = i + 6
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    
    If Trim(paragrafoLocalizado) = "NCM" Then
        
        Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
        
        i = i + 8
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
                
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("D" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("E" + CStr(posicaoNumeroLinha)).Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("F" + CStr(posicaoNumeroLinha)).Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
        
        i = i + 4
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("G" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
    End If
    
    Worksheets("NCM").Activate
    Range("A1").Select
        
    doc.Close
    Set doc = Nothing
    Set appWord = Nothing

End Sub

Private Function posicaoDeLinha() As Integer

    Dim contador As Integer
    
    contador = 1
    
    Do While Range("A" + CStr(contador)).Text <> ""
        contador = contador + 1
    Loop
    
    posicaoDeLinha = contador
    
End Function

Private Function ListaArquivos(ByVal caminho As String) As String()

'Atenção: Faça referência à biblioteca Micrsoft Scripting Runtime
Dim FSO As New FileSystemObject
Dim result() As String
Dim Pasta As Folder
Dim Arquivo As File
Dim indice As Long
  
    ReDim result(0) As String
    If FSO.FolderExists(caminho) Then
        Set Pasta = FSO.GetFolder(caminho)
 
        For Each Arquivo In Pasta.Files
            indice = IIf(result(0) = "", 0, indice + 1)
            ReDim Preserve result(indice) As String
            result(indice) = Arquivo.Name
        Next
    End If
 
    ListaArquivos = result
ErrHandler:
    Set FSO = Nothing
    Set Pasta = Nothing
    Set Arquivo = Nothing
End Function

Public Sub ListaArquivosNoDiretorio(caminho As String)
    
    Dim arquivos() As String
    Dim lCtr As Integer
    
    arquivos = ListaArquivos(caminho)
    ReDim arrayArquivo(UBound(arquivos))
    
    For lCtr = 0 To UBound(arquivos)
        
        arrayArquivo(lCtr) = arquivos(lCtr)
    Next
    
End Sub

Public Sub abrirForm()
    
    frmArquivosImportados.Show
    
End Sub
