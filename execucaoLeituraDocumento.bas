Attribute VB_Name = "execucaoLeituraDocumento"
Dim appWord As Word.Application
Dim doc As Word.Document
Dim prg As Paragraph
Dim wdrange As Word.Range

Dim paragrafoLocalizado As String
Dim codigoNCM As String

Dim i As Integer
Sub LerDocumento()
     
    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set doc = appWord.Documents.Open("C:\Users\evald\iCloudDrive\Pessoal\Concept\Projeto Contabilidade\Leitura de Documentos\legisweb-consulta-82031010.doc")

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
    
    'Descrição do Segmento
    i = i + 2
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C2").Value = paragrafoLocalizado
    
    'NCM
    i = i + 6
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("A2").Value = Trim(paragrafoLocalizado)
    
    codigoNCM = Trim(paragrafoLocalizado)
    
    'Descrição NCM
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
    Range("B2").Value = Trim(paragrafoLocalizado)
    
    'Código do CEST
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
    Range("D2").Value = Trim(paragrafoLocalizado)

    '---------------------------------------------------------------------
    'Planilha Base Legal
    '---------------------------------------------------------------------
    Worksheets("BASE_LEGAL").Activate
    Range("A2").Value = codigoNCM
    
    'UF
    i = i + 3
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("B2").Value = Trim(paragrafoLocalizado)
    
    'Descrição da Base Legal
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("BASE_LEGAL").Activate
    Range("C2").Value = Trim(paragrafoLocalizado)
    
    'Base de Cálculo
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D2").Value = IIf(Trim(paragrafoLocalizado) = "-", 0, Trim(paragrafoLocalizado))
    
    'Início da Vigência
    i = i + 5
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E2").Value = Trim(paragrafoLocalizado)
    
    'Fim da Vigência
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("F2").Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
    
    '---------------------------------------------------------------------
    'Planilha MVA Original
    '---------------------------------------------------------------------
    Worksheets("ALIQUOTAS_MVA").Activate
    Range("A2").Value = codigoNCM
    
    'MVA Original
    i = i + 8
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("B2").Value = Trim(paragrafoLocalizado)
    
    'MVA Ajustada 4%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C2").Value = Trim(paragrafoLocalizado)
    
    'MVA Ajustada 12%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D2").Value = Trim(paragrafoLocalizado)
    
    'Alíquota Interna
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E2").Value = Trim(paragrafoLocalizado)
    
    '---------------------------------------------------------------------
    'IPI
    '---------------------------------------------------------------------
    Worksheets("IPI").Activate
    
    'Descrição
    i = i + 6
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    
    If Trim(paragrafoLocalizado) = "NCM" Then
        
        Range("A2").Value = codigoNCM
        
        i = i + 8
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("B2").Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("C2").Value = Trim(paragrafoLocalizado)
                
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("D2").Value = Trim(paragrafoLocalizado)
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("E2").Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
        
        i = i + 1
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("F2").Value = IIf(Not IsDate(Trim(paragrafoLocalizado)), "31/12/2100", Trim(paragrafoLocalizado))
        
        i = i + 4
        Set prg = doc.Paragraphs(i)
        paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
        Range("G2").Value = Trim(paragrafoLocalizado)
        
    End If
    
    Worksheets("NCM").Activate
    Range("A1").Select
        
    doc.Close
    Set doc = Nothing
    Set appWord = Nothing
     
     
End Sub

Function trataCaracterTextoLidoDoc(texto As String) As String
    
    texto = Replace(texto, Chr(13), Chr(32))
    texto = Replace(texto, Chr(9), "")
    texto = Replace(texto, Chr(10), "")
    texto = Replace(texto, Chr(16), "")
    texto = Replace(texto, Chr(Asc(Mid(texto, Len(texto), 1))), "")
    
    trataCaracterTextoLidoDoc = texto
    
End Function


