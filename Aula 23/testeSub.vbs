erros = ""

a = InputBox("informe o primeiro numero")
b = InputBox("informe o segundo numero")

validaParametro(a)
validaParametro(b)

If Len(Trim(erros)) > 0 Then
    MsgBox erros, 16, "Atencao"
Else
    MsgBox "a soma do primeiro com o segundo numero e " & (CDbl(a) + CDbl(b)),64," resultado"
End If

'--------------------------------------------------------------------------------------------------

Sub validaParametro(parametro)
    if Not IsNumeric(parametro) Then
        erros = erros & "os valores devem ser numericos " & Chr(10)
    Else
        if parametro < 0 Then
            erros = erros & "os valores devem ser maiores que zero" & Chr(10)
        End If
    End If
End Sub