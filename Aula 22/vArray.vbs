Dim vArray()

contador = 1
 On Error Resume Next
 Do While True
    Redim Preserve vArray(contador)
    vArray(contador) = WeekdayName(contador)
    if Err.Number <> 0 Then
    Exit Do
    End If

    contador = contador + 1
Loop
On Error Goto 0

'MsgBox "a variavel contador é array: " & IsArray(contador)
'MsgBox "a variavel  varray é array: " & IsArray(contador)

'MsgBox "o resultado da função join(vArray) é array: " & IsArray(Join(vArray))


'Join -> função que transforma um array em um valor caracter
'   Array -> que o array que será transformado em texto
'   Delimiter -> que é um delimitador dos valores do array. Default " "

'MsgBox Join(vArray)

texto = Join(vArray, "###")

'Split -> Transforma um texto em um array
'   value -> que é o texto a ser transformado em array
'   delimitador -> que é o caracter(es) que delimitam os valores a serem transformados
'   count -> que define a quantidade de substrings a serem retornadas
'   compare -> que é o tipo de comparação 0 - binária e 1 - textual

vArray2 = Split(texto,"###",4)

MsgBox "a variavel varray2 é array: " & IsArray(vArray2)

For Each B in vArray2
    MsgBox B
Next
