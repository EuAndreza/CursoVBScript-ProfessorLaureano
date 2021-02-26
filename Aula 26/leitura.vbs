Set ArquivoExterno = CreateObject("scripting.FileSystemObject")
textoFinal = ""

Set arqLeitura = ArquivoExterno.GetFile("C:\Users\andreza.maria.araujo\OneDrive - Accenture\Documents\Estudos\VBScript\CursoVBScript-ProfessorLaureano\Aula 26\teste.txt")
Set arqStream = arqLeitura.OpenAsTextStream(1,-2)

Do While Not arqStream.AtEndOfStream
    sRecord = arqStream.ReadLine
    textoFinal = textoFinal & sRecord & Chr(10)
Loop

arqStream.Close
MsgBox textoFinal