' FormatPercent -> possui ate 5 parametros
'   Expression -> que é o numero a ser transformado
'   NumDifAfterDec -> quantidade de decimais
'   IncLeadingDig -> indica se o zero aparece ou não para os casos de decimais
'       -2 - o valor sera o configurado no windows
'       -1 - o valor zero aparecera em casos de frações menores que 1 inteiro
'        0 - o valor zero não aparecera em casos de fraçoes menores que 1 inteiro
'   UseParForNegNum -> indica se valores negativos aparecerão com sinal de "-" ou entre "()"
'       -2 - o valor sera o configurado no windows
'       -1 - apresenta numero entre parenteses
'        0 - apresenta o numero com sinal "-"
'   GroupDig -> indica se aparecera o ponto de separação de milhar, milhão etc
'       -2 - o valor sera o configurado no windows
'       -1 - apresenta ponto
'        0 - nao apresenta numero
A = 1

MsgBox FormatPercent(A,3,-1)
MsgBox FormatPercent(A,3,0)

A = 0.5

MsgBox FormatPercent(A,2,,,-1)
MsgBox FormatPercent(A,2,,,0)

A = 0.005

MsgBox FormatPercent(A,2,,-1,-1)
MsgBox FormatPercent(A,2,,0,0)