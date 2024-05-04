Attribute VB_Name = "Módulo1"
Sub registrar_vendas()

Dim moto, disp As String
Dim valor As Double
Dim qtd, ult_lin As Integer

Application.ScreenUpdating = False
'Perguntar qual moto será vendida
moto = InputBox("Qual a marca da moto?", "Registro de venda")

'Pegar na aba qual o valor da moto
Sheets("Dados").Activate
valor = WorksheetFunction.VLookup(moto, Range("A2:B10"), 2, 0)

'Pegar no outro arquivo de estoque
caminho = ThisWorkbook.Path

Workbooks.Open (caminho & "\Estoque.xlsm")
qtd = WorksheetFunction.VLookup(moto, Range("A2:B10"), 2, 0)

If qtd <> 0 Then
    disp = "Disponível"
Else
    disp = "Indisponível"
End If

ActiveWorkbook.Close
'Preencher a aba Vendas diarias
Sheets("Vendas Diárias").Activate

ult_lin = Range("A1").End(xlDown).Row + 1

Cells(ult_lin, 1).Value = Cells(ult_lin - 1, 1).Value + 1
Cells(ult_lin, 2).Value = Date
Cells(ult_lin, 3).Value = moto
Cells(ult_lin, 4).Value = valor
Cells(ult_lin, 5).Value = qtd
Cells(ult_lin, 6).Value = disp

'notificar que acabou a macro

Application.ScreenUpdating = True
MsgBox ("Cadastro feito com sucesso")

End Sub
