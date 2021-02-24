Dim vArray()
For  i = 1 To 10
    Redim Preserve vArray(i)
    vArray(i) = "valor" & i
Next
'vArray = Array("valor 1", "valor 2", "valor 3")

For Each A in vArray
    MsgBox A
Next