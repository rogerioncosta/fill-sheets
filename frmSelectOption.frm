VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectOption 
   Caption         =   "frmSelectOption "
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectOption.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare a variável como pública no UserForm
Public selectedOption As String
Public selectedOption2 As String


Private Sub ComboBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    
    ' Defina a planilha de onde os dados serão extraídos
    Set ws = ThisWorkbook.Sheets("Dados") ' Altere para o nome da sua planilha

    ' Identificar o último valor da primeira coluna (para ComboBox1)
    lastRow1 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' Preencher ComboBox1 com os valores da coluna A
    Me.ComboBox1.List = ws.Range("A1:A" & lastRow1).Value
    
    ' Identificar o último valor da segunda coluna (para ComboBox2)
    lastRow2 = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ' Preencher ComboBox2 com os valores da coluna B
    Me.ComboBox2.List = ws.Range("B1:B" & lastRow2).Value


    ' Adiciona opções ao ComboBox
'    With Me.ComboBox1
'        .AddItem "ALEX TAVARES"
'        .AddItem "AGNALDO JOSE"
'        .AddItem "JOSE APARECIDO"
'        .AddItem "LEONARDO VICTOR"
'        .AddItem "MARCELO GARCIA"
'        .AddItem "MARCOS AURELIO"
'        .AddItem "MIGUEL SAGRADO"
'        .AddItem "PAULO MAMORU"
'        .AddItem "REGINALDO COSTA"
'        .AddItem "SIDNEI SANTOS"
'        .AddItem "VALDECIR APARECIDO"
'        .AddItem "VALDINEI MENDES"
'    End With
'
'    With Me.ComboBox2
'        .AddItem "BYE4744"
'        .AddItem "MDQ1278"
'        .AddItem "MDQ1428"
'        .AddItem "MEG5943"
'        .AddItem "MEG5963"
'        .AddItem "MEG6003"
'        .AddItem "MGY3379"
'        .AddItem "MEZ4496"
'        .AddItem "MFG4893"
'        .AddItem "MFS6889"
'        .AddItem "MFS9332"
'
'    End With
End Sub

Private Sub CommandButtonOK_Click()
    ' Armazena a seleção no módulo principal
     ' Verifica se o usuário selecionou ou digitou o nome
    If Me.ComboBox1.ListIndex <> -1 Then
        selectedOption = Me.ComboBox1.Value
        
    ElseIf Me.ComboBox1.Text <> "" Then
        selectedOption = Me.ComboBox1.Text
        ' Me.Hide
    End If
    
    ' Armazena a seleção no módulo principal
    ' Verifica se o usuário selecionou ou digitou a placa
    If Me.ComboBox2.ListIndex <> -1 Then
        selectedOption2 = Me.ComboBox2.Value
        
    ElseIf Me.ComboBox2.Text <> "" Then
        selectedOption2 = Me.ComboBox2.Text
        ' Me.Hide
    End If
    
    ' Verifica se as duas opções foram preenchidas
    If selectedOption <> "" And selectedOption2 <> "" Then
        Me.Hide
    Else
        MsgBox "Por favor, selecione as duas opções, Motorista e Placa.", vbExclamation
    End If
    
    
End Sub

Private Sub CommandButtonCancel_Click()
    ' Fecha o formulário sem selecionar
    selectedOption = ""
    selectedOption2 = ""
    Me.Hide
End Sub


