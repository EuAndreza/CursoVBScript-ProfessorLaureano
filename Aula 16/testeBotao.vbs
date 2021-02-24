' Retornos do MsgBox
'   1 -> OK
'   2 -> Cancelar
'   3 -> Anular
'   4 -> Repetir
'   5 -> ignorar
'   6 -> Sim
'   7 -> não

'A = MsgBox("teste o botao",1)
'MsgBox A

'A = MsgBox("teste o botao",2)
'MsgBox A

'A = MsgBox("teste o botao",3)
'MsgBox A

'A = MsgBox("teste o botao",4)
'MsgBox A

'A = MsgBox("teste o botao",5)
'MsgBox A

'A = MsgBox("teste o botao",6)
'MsgBox A

'A = MsgBox("teste o botao",7)
'MsgBox A

A = MsgBox("teste o botao",4)

'cenario 1
'condições:
'- o operador escolhe a opção "sim"
'resultado:
'-exibe a mensagem "foi pressionado o botao sim"

'cenario 2
'condições:
'- o operador escolhe a opção "não"
'resultado:
'-exibe a mensagem "foi pressionado o botão não"

if A = 6 Then 'pressionado o botão "sim"
    MsgBox "foi pressionado o botão sim"
Else
    MsgBox "foi pressionado o botão não"
End If
