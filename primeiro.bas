Attribute VB_Name = "M�dulo1"
Sub primeiro()
'o comando Dim (Dimension) � utilizado para declarar vari�vel'
'a vari�vel nome foi tipada como String (Texto'

Dim nome As String
'O comando InputBox abre uma caixa de entrada de dados, assim o usu�rio digita o nome e aloca na vari�vel nome'

nome = InputBox("Digite Seu Nome")
'O comando Range permite selecioar uma c�lula a planilha do excel, assim selecionamos a c�lula a1 e adicionamos o valor que foi digitado na caixa de entrada, usando a vari�vel nome.'

Range("A1").Value = nome
End Sub
