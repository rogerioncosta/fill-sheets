Attribute VB_Name = "Atualizar_lista_coletas"
Function preencherDados(wsDestino As Worksheet, wsOrigem As Worksheet, linha As Integer)
        'id
        wsDestino.Range("C" & linha).Value = wsOrigem.Range("A1").Value

        ' Hora
        wsDestino.Range("B" & linha).Value = wsOrigem.Range("I3").Value
        
        ' Copiar o valor da célula A11 da planilha de origem para B3 da planilha central
        ' Local
        wsDestino.Range("D" & linha).Value = wsOrigem.Range("A11").Value
        
        ' Planta
        wsDestino.Range("E" & linha).Value = wsOrigem.Range("I12").Value
        
        ' Copiar o valor da célula C12 da planilha de origem para E3 da planilha central
        ' Produto
        'wsDestino.Range("E" & linhaDestino).Value = wsOrigem.Range("C12").Value
        
        produto = ""
        
        linhaProduto = 12

        wsOrigem.Range("C1048576").Select
        ultimaLinhaProduto = ActiveCell.End(xlUp).Row + 1
        
        wsOrigem.Range("C12").Select
        
        produtoAtual = ""

        While linhaProduto < ultimaLinhaProduto
        
            If wsOrigem.Range("C" & linhaProduto).Value = "" Then
                GoTo continue
            End If
            
            If wsOrigem.Range("C" & linhaProduto).Value <> produtoAtual Then
                produtoAtual = wsOrigem.Range("C" & linhaProduto).Value
                produto = produto & wsOrigem.Range("C" & linhaProduto).Value & "// "
                produtoAtual = Range("C" & linhaProduto).Value
            End If
continue:
            wsOrigem.Range("C" & linhaProduto).Offset(1, 0).Select
            linhaProduto = linhaProduto + 1
        Wend

        wsDestino.Range("F" & linha).Value = produto
    
        If wsDestino.Range("G" & linha).Value <> wsOrigem.Range("G9").Value Then
            ' Motorista
            wsDestino.Range("G" & linha).Value = wsOrigem.Range("G9").Value
        
            ' Veículo
            wsDestino.Range("H" & linha).Value = wsOrigem.Range("I7").Value
        
        End If
    
    End Function

Sub Atualizar_listar_coletas()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim pastaTrabalho As Workbook
    Dim caminhoPasta As String
    Dim arquivo As String
    'Dim linhaDestino As Integer
    'Dim linhaNaoPreenchida
    Dim produto
    Dim linhaProduto
    Dim produtoAtual
    Dim linha As Integer
    
    ' Define a planilha atual como a de destino, onde os valores serão colados
    Set wsDestino = ThisWorkbook.Sheets("Lista Coletas") ' Substitua pelo nome da aba de destino
    
    
    ' Define a pasta onde estão as planilhas de origem
    ' caminhoPasta = "C:\Users\Win10\Desktop\vaga_push\" ' Substitua pelo caminho da sua pasta
    ' arquivo = Dir(caminhoPasta & "*.xlsx*")
    
    ' Construir o caminho dinâmico para o arquivo no desktop do usuário
    'caminhoPasta = Environ("USERPROFILE") & "\NOROESTE\"
    caminhoPasta = "\\ESCRITORIO-PC1\Scan\01 Planilhas Consulta\_treinamentos\NOROESTE\"
    arquivo = Dir(caminhoPasta & "*.xlsx*")

    'linhaDestino = 3 ' Definindo a linha de destino inicial
    
    wsDestino.Range("B1048576").Select
    'linhaDestino
    ultimaLinha = ActiveCell.End(xlUp).Row + 1
    '**
    
    linha = 3
    
    wsDestino.Range("C" & ultimaLinha).Select
    
    'linhaId = 3
    
    wsDestino.Range("C3").Select
    
    If arquivo = "" Then
        MsgBox "Não há nenhuma planilha .xlsx para puxar os dados!", vbInformation
        Exit Sub
    End If
            
    Do While arquivo <> ""
        
        Set pastaTrabalho = Workbooks.Open(caminhoPasta & arquivo)
        Set wsOrigem = pastaTrabalho.Sheets(2)

        ' LÓGICA PARA PREENCHER A LISTA PELA PRIMEIRA VEZ
        'data
        If wsDestino.Range("A1").Value <> wsDestino.Range("H1").Value Then
            wsDestino.Range("B3:I27").ClearContents
            wsDestino.Range("A1").Value = wsDestino.Range("H1").Value
        End If
'continue2:
        If wsDestino.Range("C" & linha).Value = "" Then
            
            ' *****Chamada para a função de preencher os dados
            Call preencherDados(wsDestino, wsOrigem, linha)

            ' Fechar a planilha de origem sem salvar alterações
            pastaTrabalho.Close False
        
            ' Próximo arquivo
            arquivo = Dir
            
            'wsDestino.Range("B" & linha).Offset(1, 0).Select
            linha = linha + 1
            
            'LÓGICA PARA ATUALIZAR
        Else
            conferirId = 0
            ultimaLinhaMenosUm = ultimaLinha - 1
            
            While linha <= ultimaLinha
            
                If wsOrigem.Range("A1").Value = wsDestino.Range("C" & linha).Value Then
                
                    conferirId = wsDestino.Range("C" & linha).Value
                    
                    If wsDestino.Range("G" & linha).Value <> wsOrigem.Range("G9").Value Then
                        ' Motorista
                        wsDestino.Range("G" & linha).Value = wsOrigem.Range("G9").Value
            
                        ' Veículo
                        wsDestino.Range("H" & linha).Value = wsOrigem.Range("I7").Value
                    End If
                    
                End If
                
                If linha >= ultimaLinha Then
                    If conferirId = 0 And wsDestino.Range("B" & ultimaLinhaMenosUm).Value <> "" Then
                    
                        ultimaLinha = wsDestino.Range("B1048576").End(xlUp).Row + 1
                        'linhaDestino
                        'ultimaLinha = ActiveCell.End(xlUp).Row + 1
                        
                        'wsDestino.Range("C" & ultimaLinha).Select
                        linha = ultimaLinha
                        ' *****Chamada para a função de preencher os dados
                        Call preencherDados(wsDestino, wsOrigem, linha)
                        
                        
                        
                    End If
                
                End If
                
                'wsDestino.Range("B" & linha).Offset(1, 0).Select
                linha = linha + 1
            
            Wend
            
            linha = 3
            ' Fechar a planilha de origem sem salvar alterações
           
            pastaTrabalho.Close False
        
            ' Próximo arquivo
            arquivo = Dir
        
        End If

    Loop
    
    ' Ordenar por horário
    Range("B3:I27").Select
    
    Call OrdenarAZ

    'MsgBox "Valores copiados com sucesso!"
    
    ThisWorkbook.Save
    
    MsgBox "Valores copiados, e arquivo salvo com sucesso!", vbInformation
    
    
End Sub
