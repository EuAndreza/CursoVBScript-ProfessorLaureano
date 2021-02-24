' InStr possui ate 3 parametros
'  caracter inicial
'  texto original
'  texto procurado

Z = InputBox("informe o indice do caracter inicial") + 0
A = InputBox("informe qualquer valor")
B = InputBox("caracter de procura")

C = InStr(Z,A,B)

MsgBox "o(os) caracteres " & B & "esta na posicao " & C