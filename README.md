# Tratamento de Dados
Repositório para Consolidar formas para a Extração, Tratamento e Carregamento de Dados.
## Opening archive / Abrindo arquivo
### Visual Basic for Applications
> Para ganho de processamento  evite utilizar comandos como ;
> 
>`.Select`  e `.Activate`, prefira expecificar a ação que deseja realizar. 
> 
> `.Copy` e `.Paste` ou `.PasteSpecial`, prefira expecificar o caminho do conteúdo que deseja inserir  Ex.: `Range(B1) = Range(A1).Value` .
> 
> E insira antes de qualquer declaração (no inicio do Código) os comandos;
> 
> `Application.Calculation = xlCalculationManual`, e se for necessário atualizar algum calculo das Formulas, insira antes do `End Sub` o comando `Application.Calculation = xlCalculationAutomatic`.
> `Application.ScreenUpdating = False`, se realmente for necessario que o Excel fique atualizando a tela enquanto o código é executado, deixe como `Application.ScreenUpdating = true`.

- `Workbook` = "arquivo" que esta salvo o Script.

- `ThisWorkbook` = objeto "arquivo" que esta salvo o Script e ActiveWorkbook = objeto "arquivo" que esta ativo no momento de execução do Script.

- `Workbook.Path` = Propriedade que retorna o nome no caminho onde o objeto "arquivo" esta salvo.

- `Workbook.FullName` = Propriedade que retorna o nome completo do objeto "arquivo".

- `Workbooks.Open " C:\Users\Desktop\arquivo.xlsx " ` = Abre o Arquivo informado no parametro, entre " ".

- `Workbooks.Close` = Fecha o Arquivo, se informar False depois do parametro .Close o arquivo não será salvo automaticamente.

- `Workbook.SaveAs "C:\Users\Desktop\arquivo.xlsx", _
    xlOpenXMLStrictWorkbook` = Salva  o arquivo no caminho e com o nome expecificado entre "  ", e  é necessario informar depois do paramentro o formato do arquivo como por exemplo XLSX que é xlOpenXMLStrictWorkbook.

- `Workbooks.Add`  =  Cria um novo arquivo .


### SQL
- `SELECT tablespace_name, table_name, owner FROM dba_tables;` = listar todas as tabelas do Banco.

- `SELECT tablespace_name, table_name, owner FROM all_tables;` = listar todas as tabelas às quais o usuário tem acesso (sendo ele o owner (dono) ou não).

- `SELECT tablespace_name, table_name, owner FROM user_tables;` = listar todas as tabelas do usuário corrente.

- `desc <tabela>;` = retorna as características da tabela selecionada.

- `SELECT * FROM <tabela>;` = retorna todas as colunas e tuplas da tabela, utilize o '*' somente quando estiver conhecendo a tabela, pois sobrecarrega a performace do banco e o retorno dos dados.




