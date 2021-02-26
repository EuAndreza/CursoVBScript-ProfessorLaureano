Set ArquivoExterno = CreateObject("Scripting.FileSystemObject")
Set arq = ArquivoExterno.OpenTextFile("C:\Users\andreza.maria.araujo\OneDrive - Accenture\Documents\Estudos\VBScript\CursoVBScript-ProfessorLaureano\Aula 25\teste.txt",2,1)

arq.WriteLine "primeira linha"
arq.WriteLine "Segunda linha"
arq.WriteLine "terceira linha"
arq.WriteLine "quarta linha"
arq.WriteLine "quinta linha"
arq.WriteLine "sexta linha"
arq.WriteLine "setima linha"

arq.Close

MsgBox "finalizado", 64, "arquivo texto"