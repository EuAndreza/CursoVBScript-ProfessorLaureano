'MSGBOX é composto de até 5 parâmetros
'PROMPT -> que é a mensagem que aparece no centro da caixa de texto 
'BUTTON -> que é composto de 4 grupos números
' Grupo 1 -> configurações do botões
' Grupo 2 -> ícones da caixa de texto
' Grupo 3 -> Botão padrão
' Grupo 4 -> Modal
'TITLE -> que é mensagem que aparece na barra de titulo da caixa de texto
'HELPFILE -> que é o endereço do arquivo de ajuda da caixa de texto 
'CONTEXT -> que é o contexto do arquivo de ajuda


'Grupo 1

'Cenário 1 
'Condições:
'- Ser informado o valor 0
'Resultado:
'- A caixa de texto exibe o botão "OK"

'Cenário 2
'Condições:
'- Ser informado o valor 1
'Resultado:
'- A caixa de texto exibe o botão "OK", "cancelar"

'Cenário 3
'Condições:
'- Ser informado o valor 2
'Resultado:
'- A caixa de texto exibe o botão "Abortar", "Repetir" e "cancelar"

'Cenário 4
'Condições:
'- Ser informado o valor 3
'Resultado:
'- A caixa de texto exibe o botão "sim", "NU" e "cancelar"

'Cenário 5
'Condições:
'- ser informado o valor 4
'Resultado:
'- A caixa de texto exibe o botão "sim" e "Waco"

'Cenário 6
'Condições:
'- Ser informado o valor 5
'Resultado:
'- A caixa de texto exibe o voto "Repetir" e "cancelar"


'Grupo 2

'Cenário 1
'Condicões:
'- Ser informado o valor 0
'Resultado:
'- A caixa de texto não exibira ícone

'Cenário 2
'Condicões:
'- Ser informado o valor 16
'Resultado:
'- A caixa de texto exibira ícone de erro

'Cenário 3
'Condicões:
'- Ser informado o valor 32
'Resultado:
'- A caixa de texto exibirá ícone de questão

'Cenário 4
'Condicões:
'- Ser informado o valor 48
'Resultado:
'- A caixa de texto exibirá ícone de exclamação

'Cenário 5
'Condicões:
'- Ser informado o valor 64
'Resultado:
'- A caixa de texto exibira ícone informativo
 



'Grupo 3

'Cenário 1
'Condições:
'- Ser informado o valor 0 
'Resultado:
'- O primeiro botão ficará focado

'Cenário 2
'Condições:
'- Ser informado o valor 256 
'Resultado:
'- O segundo botão ficar focado

'Cenário 3
'Condições:
'- Ser informado o valor 512
'Resultado:
'- O terceiro botão ficar focado

'Cenário 4
'Condições:
'- Ser informado o valor 768 
'Resultado:
'- O quarto botão ficar focado


'Grupo 4

'Cenário 1
'Condições:
'- Ser informado 0
'Resultado:
'- Modal desativado

'Cenário 2
'Condições:
'- Ser informado 4096
'Resultado:
'- Modal ativado

msgbox "ola mundo", 2+16+256, "ATENÇÃO"