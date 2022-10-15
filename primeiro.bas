Attribute VB_Name = "Módulo1"
Sub primeiro()
'o comando Dim (Dimension) é utilizado para declarar variável'
'a variável nome foi tipada como String (Texto'

Dim nome As String
'O comando InputBox abre uma caixa de entrada de dados, assim o usuário digita o nome e aloca na variável nome'

nome = InputBox("Digite Seu Nome")
'O comando Range permite selecioar uma célula a planilha do excel, assim selecionamos a célula a1 e adicionamos o valor que foi digitado na caixa de entrada, usando a variável nome.'

Range("A1").Value = nome
End Sub
