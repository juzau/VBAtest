Attribute VB_Name = "M�dulo1"
Sub primeiro()
'O comando DIM(Dimension) � utilizado para declarar vari�vel.
'A vari�vel nome foi tipada como String(Texto)

Dim nome As String

'O comando InputBox abre uma caixa de entrada de dados
'Assim o usu�rio digita o nome e aloca na
'vari�vel nome

nome = InputBox("Digite o seu nome")

'O comando range permite selecionar uma c�lula na planilha do excel,
'Assim selecionamos a c�lula A1 e adicionamos o valor que foi digitado na caixa de entrada
'Usando a vari�vel nome

Range("A1").Value = nome
End Sub

