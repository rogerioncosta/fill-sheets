Attribute VB_Name = "puxar_horario"
Sub PuxarHora()
    Dim pasta As String
    Dim fso As Object
    Dim arquivo As Object
    Dim wb As Workbook
    Dim nomeArquivoOriginal As String
    
    Dim id As Integer
    
    id = ThisWorkbook.Sheets("ID").Range("A1").Value

    ' Especifique o caminho da pasta
    'pasta = Environ("USERPROFILE") & "\Desktop\NOROESTE\"
    
    pasta = "\\ESCRITORIO-PC1\Scan\01 Planilhas Consulta\_treinamentos\NOROESTE\"

    ' Crie um FileSystemObject para acessar os arquivos
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Verifique se a pasta existe
    If Not fso.FolderExists(pasta) Then
        MsgBox "A pasta especificada não existe.", vbExclamation
        Exit Sub
    End If

    ' Percorra cada arquivo na pasta
    For Each arquivo In fso.GetFolder(pasta).Files
        ' Verifique se o arquivo é uma planilha do Excel (pode ajustar a extensão conforme necessário)
        If LCase(Right(arquivo.Name, 4)) = ".xls" Or LCase(Right(arquivo.Name, 5)) = ".xlsx" Then
            ' Abra o arquivo
            Set wb = Workbooks.Open(arquivo.Path)
            
            ' Armazene o nome do arquivo original (sem extensão)
            nomeArquivoOriginal = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
            
            ' Coloque aqui o que você precisa fazer com o arquivo aberto
            ' Exemplo:
            'MsgBox "Processando arquivo: " & nomeArquivoOriginal
            
            If Range("I3") = "__:__" Then
                Range("I3").Value = 0
            End If
            
            ' Novo nome concatenado com o nome original e o valor da célula G5
            nomeArquivoNovo = Format(wb.Sheets(2).Range("I3").Value, "H\H") & "_" & nomeArquivoOriginal
                        
            If nomeArquivoNovo = "" Then
                MsgBox "A célula I3 está vazia. Por favor, insira um nome de arquivo.", vbExclamation
                Exit Sub
            End If
                
            
            'id = Int(1 + Rnd * (100 - 2 + 1))
            
            'wb.Sheets(2).Range("A1").Value = id
            
            If wb.Sheets(2).Range("A1").Value = "" Then
    
                wb.Sheets(2).Range("A1").Value = id
        
                id = id + 1
        
                ThisWorkbook.Sheets("ID").Range("A1").Value = id
            End If
            
            ' Define o caminho completo para salvar o arquivo
            novoCaminho = wb.Path & "\" & nomeArquivoNovo & ".xlsx"
            
            ' Salva o arquivo com o novo nome
            wb.SaveAs Filename:=novoCaminho, FileFormat:=xlOpenXMLWorkbook
            'MsgBox "Arquivo salvo como " & novoCaminho, vbInformation
            
            ' Feche o arquivo sem salvar
            wb.Close SaveChanges:=False
            
            Kill arquivo.Path
        End If
        
        'id = id + 1
        
        'ThisWorkbook.Sheets("ID").Range("A1").Value = id
        
    Next arquivo
    
    MsgBox "Os arquivos foram salvos com os horários.", vbInformation

    ' Limpeza
    Set fso = Nothing
End Sub


