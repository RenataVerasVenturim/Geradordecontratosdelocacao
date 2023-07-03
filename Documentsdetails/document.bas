Just use this document to GITHUB undestand the VBA language


VBA´s Codes of this project:

Sub criar_novo_contrato()

Set objWord = CreateObject("Word.Application")

objWord.Visible = True

Set arqContrato = objWord.Documents.Open(ThisWorkbook.Path & "\Modelo Contrato" & ".docx")
Set conteudoDoc = arqContrato.Application.Selection

For colTab = 1 To 24
    conteudoDoc.Find.Text = Cells(1, colTab).Value
    conteudoDoc.Find.Execute
    conteudoDoc.Range = Cells(2, colTab).Value
    
Next
    
arqContrato.saveas2 (ThisWorkbook.Path & "/Contratos" & "/Contrato - " & Cells(2, 1).Value & ".docx")

arqContrato.Close
objWord.Quit

Set arqContrato = Nothing
Set conteudoDoc = Nothing
Set objWord = Nothing

MsgBox ("Contrato gerado e salvo na Pasta 'Contratos' com sucesso!")

End Sub

Sub abrir_formulario()

inserir_dados_contrato.Show

End Sub



