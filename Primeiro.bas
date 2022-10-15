Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando DIM(Dimension) é utilizado para declarar variável.
'A variável nome foi tipada como String(Texto)

Dim nome As String

'O comando InputBox abre uma caixa de entrada de dados
'Assim o usuário digita o nome e aloca na
'variável nome

nome = InputBox("Digite o seu nome")

'O comando range permite selecionar uma célula na planilha do excel,
'Assim selecionamos a célula A1 e adicionamos o valor que foi digitado na caixa de entrada
'Usando a variável nome

Range("A1").Value = nome
End Sub

