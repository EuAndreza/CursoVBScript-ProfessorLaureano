' IsNumeric -> Que possui um parametro que é o valor a ser testado
'   validar se o valor informado é um numero
'IsDate -> que possui um parametro que é o valor a ser testado
'   validar se o valor informado é uma data/hora
'IsNull -> que possui um parametro que é o valor a ser testado
'   validar se o valor informado é nulo
'IsEmpty -> que possui um parametro que é o valor a ser testado
'   validar se a variavel esta vazia


A = InputBox("digite um valor qualquer")
B = Null

if IsNumeric(A) Then
    MsgBox "valor informado e numerico"
ElseIf  isDate(A) Then
    MsgBox "valor informado e data"
Else
    MsgBox "valor informado e caracter"
End if

If IsNull(B) Then
    MsgBox "valor B igual a nulo"
End If

If IsEmpty(C) Then
    MsgBox "valor C e vazio"
End if
    