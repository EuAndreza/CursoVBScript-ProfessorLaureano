erros = ""

a = InputBox("informe o primeiro numero")
b = InputBox("informe o segundo numero")

'validaParametro(a)
'validaParametro(b)

If Not validaParametro(a,b) Then
    MsgBox erros, 16, "Atencao"
Else
    MsgBox "a soma do primeiro com o segundo numero e " & (CDbl(a) + CDbl(b)),64," resultado"
End If

'--------------------------------------------------------------------------------------------------

Function validaParametro(parametro1, parametro2)
    if Not IsNumeric(parametro1) Then
        erros = erros & "o valor do parametro1 deve ser numerico " & Chr(10)
    Else
        if parametro1 < 0 Then
            erros = erros & "o valor do parametro1 deve ser maior que zero" & Chr(10)
        End If
    End If

    if Not IsNumeric(parametro2) Then
        erros = erros & "o valor do parametro2 deve ser numerico " & Chr(10)
    Else
        if parametro2 < 0 Then
            erros = erros & "o valor do parametro2 deve ser maior que zero" & Chr(10)
        End If
    End If

    If Len(Trim(erros)) > 0 Then
        validaParametro = False
    Else
        validaParametro = True
    End If
End Function