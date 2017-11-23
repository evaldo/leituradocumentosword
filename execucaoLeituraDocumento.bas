Attribute VB_Name = "execucaoLeituraDocumento"
Sub LerDocumento()
  
    Dim appWord As Word.Application
    Dim doc As Word.Document
    Dim prg As Paragraph
    Dim wdrange As Word.Range
    
    Dim paragrafoLocalizado As String
    Dim codigoNCM As String
    
    Dim i As Integer
    
    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set doc = appWord.Documents.Open("C:\Users\evald\iCloudDrive\Pessoal\Concept\Projeto Contabilidade\Leitura de Documentos\legisweb-consulta-72172090.docx")

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
    'Descrição do Segmento
    '---------------------------------------------------------------------
    i = i + 2
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
    Range("C2").Value = paragrafoLocalizado
    
    i = i + 6
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("NCM").Activate
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
    'UF
    '---------------------------------------------------------------------
    i = i + 3
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("BASE_LEGAL").Activate
    Range("B2").Value = paragrafoLocalizado
    Range("A2").Value = codigoNCM
    
    'Descrição da Base Legal
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("BASE_LEGAL").Activate
    Range("C2").Value = paragrafoLocalizado
    
    'Base de Cálculo
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D2").Value = paragrafoLocalizado
    
    'Início da Vigência
    i = i + 5
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E2").Value = paragrafoLocalizado
    
    'Fim da Vigência
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("F2").Value = paragrafoLocalizado
    
    '---------------------------------------------------------------------
    'Planilha MVA Original
    'MVA Original
    '---------------------------------------------------------------------
    i = i + 8
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Worksheets("ALIQUOTAS_MVA").Activate
    Range("B2").Value = paragrafoLocalizado
    Range("A2").Value = codigoNCM
    
    'MVA Ajustada 4%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("C2").Value = paragrafoLocalizado
    
    'MVA Ajustada 12%
    i = i + 1
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("D2").Value = paragrafoLocalizado
    
    'Alíquota Interna
    i = i + 4
    Set prg = doc.Paragraphs(i)
    paragrafoLocalizado = trataCaracterTextoLidoDoc(prg.Range.Text)
    Range("E2").Value = paragrafoLocalizado
     
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


