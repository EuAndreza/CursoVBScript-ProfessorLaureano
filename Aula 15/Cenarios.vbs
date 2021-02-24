' operadores da estrutura de decisao
' = -> que verifica e valida se o primeiro valor é igual ao segundo valor 
' < -> que verifica e valida se o primeiro valor é menor que segundo valor
' > -> que verifica e valida se o primeiro valor é maior que segundo valor
'<= -> que verifica e valida se o primeiro valor é menor ou igual ao segundo valor
'>= -> que verifica e valida se o primeiro valor é maior ou igual ao segundo valor
'<> -> que verifica e valida se o primeiro valor é diferente do segundo valor



'Evento -> informar um numero qualquer

'Cenario 1
'condições:
'- Numero informado ser igual a zero
'Resultado:
'- o script exibirá a mensagem: " o campo informado é igual a zero"

'Cenario 2
'condições:
'- Numero informado ser menor que zero
'Resultado:
'- o script exibirá a mensagem: " o campo informado é menior que zero"

'Cenario 3
'condições:
'- Numero informado ser maior que zero
'Resultado:
'- o script exibirá a mensagem: " o campo informado é maior que zero"

A = InputBox("informe um numero")

If CDbl(A) = 0 Then
    MsgBox "o campo informado é igual a zero"
ElseIf Cdbl(A) < 0 Then
    MsgBox "o campo informado é menor que zero"
Else
    MsgBox "o campo informado é maior que zero"
End If
    