A = InputBox("informe o primeiro numero")
B = InputBox("informe o segundo numero")

On Error Resume Next
C = (A/B)

If Err.Number <> 0 Then
    MsgBox "erro tratado"
Else
    MsgBox "o resultado e" & C
End If
On Error Goto 0
