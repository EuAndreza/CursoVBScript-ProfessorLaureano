Dim vArray()

contador = 1

On Error Resume Next
Do While True
    Redim Preserve vArray(contador)
    vArray(contador) = WeekdayName(contador)
    If Err.Number <> 0 Then
        Exit Do
    End If

    contador = contador + 1
Loop
On Error Goto 0

'For i=1 To UBound(vArray) -1
'   MsgBox vArray(i)
'Next
B = Filter(vArray, "-", True)
'Filter temos ate 4 parametros
'   Array -> que é o array  em que se faz o filtro
'   Value -> que é o valor procurado
'   Include -> que é um valor booleano (true ou false) para filtrar se contem ou não o valor do parametro
'       value
'   compare -> 0 comparação de texto, 1 comparação binaria
For i = 0 to UBound(B)
    MsgBox B(i)
Next