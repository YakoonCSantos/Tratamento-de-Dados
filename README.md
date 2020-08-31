# Tratamento de Dados
Repositório para Consolidar formas para a Extração, Tratamento e Carregamento de Dados.
## Opening archive / Abrindo arquivo
### Visual Basic for Applications
> Para ganho de processamento  evite utilizar comandos como ;
>`.Select`  e `.Activate`, prefira expecificar a ação que deseja realizar. 
> `.Copy` e `.Paste` ou `.PasteSpecial`, prefira expecificar o caminho do conteúdo que deseja inserir  Ex.: `Range(B1) = Range(A1).Value` .
> E insira antes de qualquer declaração (no inicio do Código) os comandos;
> 'Application.Calculation = xlCalculationManual ', e se for necessário atualizar algum calculo das Formulas, insira antes do 'End Sub' o comando 'Application.Calculation = xlCalculationAutomatic'

- `Workbook` = "arquivo" que esta salvo o Script.

- `ThisWorkbook` = objeto "arquivo" que esta salvo o Script e ActiveWorkbook = objeto "arquivo" que esta ativo no momento de execução do Script.

- `Workbook.Path` = Propriedade que retorna o nome no caminho onde o objeto "arquivo" esta salvo.

- `Workbook.FullName` = Propriedade que retorna o nome completo do objeto "arquivo".

- `Workbooks.Open " C:\Users\Desktop\arquivo.xlsx " ` = Abre o Arquivo informado no parametro, entre " ".

- `Workbooks.Close` = Fecha o Arquivo, se informar False depois do parametro .Close o arquivo não será salvo automaticamente.

- `Workbook.SaveAs "C:\Users\Desktop\arquivo.xlsx", _
    xlOpenXMLStrictWorkbook` = Salva  o arquivo no caminho e com o nome expecificado entre "  ", e  é necessario informar depois do paramentro o formato do arquivo como por exemplo XLSX que é xlOpenXMLStrictWorkbook.

- `Workbooks.Add`  =  Cria um novo arquivo .
