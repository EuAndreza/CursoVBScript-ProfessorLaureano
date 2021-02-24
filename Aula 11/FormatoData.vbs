' FormatDateTime possui dois paramentros
'   valor de tempo
'   formato de saida
'       0 - vbGeneralDate - Formato (dd/mm/aaaa hh:mi:ss) "valor padrao"
'       1 - vbLongDate - Formato por extenso
'       2 - vbShortDate - Formato (dd/mm/aaaa)
'       3 - vbLongTime - Formato (hh:mi:ss)
'       4 - vbShortTime - Formato (hh:ss)
A = Now()

MsgBox FormatDateTime(A)
MsgBox FormatDateTime(A,0)
MsgBox FormatDateTime(A,1)
MsgBox FormatDateTime(A,2)
MsgBox FormatDateTime(A,3)
MsgBox FormatDateTime(A,4)