Dim vArray()
For  i = 1 To 10
    Redim Preserve vArray(i)
    vArray(i) = "valor" & i
Next

MsgBox LBound(vArray) 'Lbound -> retorna o menor indice de uma aray
MsgBox UBound(vArray) 'Ubound -> retorna o menor indice de uma aray

vArray(0) = "valor 0"

For i = 1 To UBound(vArray)
    MsgBox vArray(i)
Next 

For i=LBound(vArray) To UBound(vArray)
    MsgBox vArray(i)
Next 

For Each A in vArray
    MsgBox A
Next