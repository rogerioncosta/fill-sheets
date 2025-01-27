Attribute VB_Name = "Nomear_coletas"
Sub Nomear_ordens()
  
    Dim caminho As String
    Dim wb As Workbook
    Dim nomeArquivoOriginal As String
    Dim nomeArquivoNovo As String
    Dim novoCaminho As String
    
    Dim id As Integer
    
    id = ThisWorkbook.Sheets("ID").Range("A1").Value

    ' Seleciona o arquivo a ser aberto
    caminho = Application.GetOpenFilename()
    
    If caminho = "Falso" Then
        MsgBox "Nenhum arquivo selecionado.", vbExclamation
        Exit Sub
    End If
    
    ' Abre o workbook e define a referência para ele
    Set wb = Workbooks.Open(caminho)
    
    ' Armazena o nome do arquivo original (sem extensão)
    nomeArquivoOriginal = Left(wb.Name, InStrRev(wb.Name, ".") - 1)

    ' Configuração de validação na célula I7
'    With wb.Sheets("21.02.2014").Range("I7").Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
'             Formula1:="BYE4744, MEG5943, MEG5963, MEG6003, MEZ4496, MFG4893, MFS6889, MFS9332"
'    End With
    
    ' Configuração de validação na célula G9
'    With wb.Sheets("21.02.2014").Range("G9").Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
'             Formula1:="ALEX TAVARES, AGNALDO JOSE, JOSE APARECIDO, MARCOS AURELIO, PAULO MAMORU, VALDECIR APARECIDO, VALMIR JOSUE"
'    End With

    ' Mostra o UserForm para selecionar a opção
    frmSelectOption.Show

    ' Acesse a variável diretamente do UserForm
    If frmSelectOption.selectedOption = "" And frmSelectOption.selectedOption2 = "" Then
        MsgBox "Nenhuma opção selecionada.", vbExclamation
        Exit Sub
    End If

    ' Atribui a seleção à célula I7 e G9
    ' wb.Sheets("21.02.2014").Range("I7").Value = UCase(frmSelectOption.selectedOption2)
    wb.Sheets(2).Range("I7").Value = UCase(frmSelectOption.selectedOption2)
    wb.Sheets(2).Range("G9").Value = UCase(frmSelectOption.selectedOption)

    ' Espera opcional de 5 segundos
'    Application.Wait Now + TimeValue("00:00:05")
    
    ' Novo nome concatenado com o nome original e o valor da célula G5
    'nomeArquivoNovo = nomeArquivoOriginal & "_" & wb.Sheets("21.02.2014").Range("G9").Value & "_" & Format(wb.Sheets("21.02.2014").Range("I3").Value, "H\H") & "_" & Format(Date, "dd-mm-yyyy")
    nomeArquivoNovo = nomeArquivoOriginal & "_" & wb.Sheets(2).Range("G9").Value & "_" & wb.Sheets(2).Range("I7").Value & "_" & Format(Date, "dd-mm-yyyy")
    If nomeArquivoNovo = "" Then
        MsgBox "A célula G9 está vazia. Por favor, insira um nome de arquivo.", vbExclamation
        Exit Sub
    End If
    
    If wb.Sheets(2).Range("A1").Value = "" Then
    
        wb.Sheets(2).Range("A1").Value = id
        
        id = id + 1
        
        ThisWorkbook.Sheets("ID").Range("A1").Value = id
    End If
    
    ' Define o caminho completo para salvar o arquivo
    novoCaminho = wb.Path & "\" & nomeArquivoNovo & ".xlsx"
        
    ' Salva o arquivo com o novo nome
    wb.SaveAs Filename:=novoCaminho, FileFormat:=xlOpenXMLWorkbook
    MsgBox "Arquivo salvo como " & novoCaminho, vbInformation
    
    ' Fecha o workbook
    wb.Close False
    
    Kill caminho

End Sub


