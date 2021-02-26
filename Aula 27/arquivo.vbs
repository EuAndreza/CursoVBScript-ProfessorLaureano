Set arquivo = CreateObject("Scripting.FileSystemObject")
Set arq = arquivo.GetFile("C:\Users\andreza.maria.araujo\OneDrive - Accenture\Documents\Estudos\VBScript\CursoVBScript-ProfessorLaureano\Aula 27\teste.txt")

propriedade = "Data de criacao -> [" & arq.DateCreated & "]" & Chr(10) & _
            "Data  do ultimo acesso -> [" & arq.DateLastAccessed & "]" & Chr(10) & _
            "Data da ultima modificacao -> [" & arq.DateLastModified & "]" & Chr(10) & _
            "Drive -> [" & arq.Drive & "]" & Chr(10) & _
            "Nome -> [" & arq.Name & "]" & Chr(10) & _
            "Pasta -> [" & arq.ParentFolder & "]" & Chr(10) & _
            "Nome completo -> [" & arq.Path & "]" & Chr(10) & _
            "Nome DOS -> [" & arq.ShortName & "]" & Chr(10) & _
            "Nome completo DOS -> [" & arq.ShortPath & "]" & Chr(10) & _
            "Tamanho -> [" & arq.Size & "]" & Chr(10) & _
            "Extensao -> [" & arquivo.GetExtensionName(arq.Path) & "]" & Chr(10) & _
            "Tipo -> [" & arq.Type & "]"
MsgBox propriedade

Set rede = CreateObject("WScript.Network")
MsgBox "voce esta conectado com o usario: " & rede.UserName