VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmArquivosImportados 
   Caption         =   "Importação de NCMs via arquivos no formato Microsoft Word"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   OleObjectBlob   =   "frmArquivosImportados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmArquivosImportados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFechar_Click()

    Unload Me

End Sub

Private Sub btnImportar_Click()

    Dim i As Integer

    For i = 0 To lstArquivo.ListCount - 1
        If lstArquivo.Selected(i) = True Then
           Call LerPorDocumento(Trim(txtCaminho + "\" + lstArquivo.List(i)))
           lstArquivo.Selected(i) = False
        End If
    Next i
    
    MsgBox "Processamento realizado com sucesso!", vbOKOnly, "Importação de Arquivos no formato Microsoft Word"
    Exit Sub
    
End Sub

Private Sub btnListarArquivo_Click()
    
    ListaArquivosNoDiretorio (txtCaminho.Text)
    lstArquivo.List = arrayArquivo
    
End Sub

