Attribute VB_Name = "execucaoLeituraDocumentoFederal"
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
    
    i = 8
    
    '---------------------------------------------------------------------
    'IPI
    '---------------------------------------------------------------------
    Worksheets("IPI").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    
    'NCM
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("A" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    codigoNCM = Trim(paragrafoLocalizado)
    
    'Descrição
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("B" + CStr(posicaoNumeroLinha)).Value = paragrafoLocalizado
    
    'Aliquota
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = paragrafoLocalizado + trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
    
    'Descrição Completa
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)

    '---------------------------------------------------------------------
    'PIS-COFINS
    '---------------------------------------------------------------------
    Worksheets("PIS-COFINS").Activate
    posicaoNumeroLinha = posicaoDeLinha()
        
    Do While Trim(paragrafoLocalizado) <> "PIS/COFINS"
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        
    Loop
    
    i = i + 6
    
    Do While Trim(paragrafoLocalizado) <> "DESONERAÇÃO"
        
        If paragrafoLocalizado <> "" Then
        
            Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
        
            'APLICACAO
            Set prg = doc.Paragraphs(i)
            paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
            Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
            'REGIME
            i = i + 1
            Set prg = doc.Paragraphs(i)
            paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
            Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
            
            'PIS
            i = i + 1
            Set prg = doc.Paragraphs(i)
            paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
            Range("D" + CStr(posicaoNumeroLinha)).Value = IIf(Trim(paragrafoLocalizado) = "-", 0, Trim(paragrafoLocalizado))
            
            'COFINS
            i = i + 1
            Set prg = doc.Paragraphs(i)
            paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
            Range("E" + CStr(posicaoNumeroLinha)).Value = IIf(Trim(paragrafoLocalizado) = "-", 0, Trim(paragrafoLocalizado))
            
        End If
        
        i = i + 1
        
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        
        If Trim(paragrafoLocalizado) = "DESONERAÇÃO" Then
            Exit Do
        End If
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
           
        posicaoNumeroLinha = posicaoNumeroLinha + 1
        
    Loop
     
    Do While Trim(paragrafoLocalizado) <> "ALÍQUOTA CPRB"
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        
    Loop
    
    Worksheets("DESONERACAO").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    
    Do While Trim(paragrafoLocalizado) <> "Base legal"
    
        Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
        
        'SERVIÇO
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
         
        'ALÍQUOTA CPRB
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        posicaoNumeroLinha = posicaoNumeroLinha + 1
        
    Loop
    
    Worksheets("CARGA TRIBUTARIA").Activate
    posicaoNumeroLinha = posicaoDeLinha()
    
    Do While Trim(paragrafoLocalizado) = "ST INTERNA"
        
        Range("A" + CStr(posicaoNumeroLinha)).Value = codigoNCM
        
        'Estado
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("B" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
         
        'Média Nacional
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("C" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        'Média Estadual
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("D" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        'Média Nacional + Estadual
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("E" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        'Média Importação
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("F" + CStr(posicaoNumeroLinha)).Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        posicaoNumeroLinha = posicaoNumeroLinha + 1
        
    Loop
    
    Worksheets("IPI").Activate
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
