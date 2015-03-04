# DisableDL-AD 
## Disable Dont Listed User AD

--------------------

É comum o CPF ser a chave primaria para a identificação de uma pessoa em um banco de dados relacional.Portanto sistemas de folha de pagamento em sua maioria controla a o status do funcionário para com a empresa através do CPF. O DisableDL-AD é um script basico que desativa usuários não listados em uma determinada lista gerada em csv, que contem o cpf de usuários ativos de um determinado sistema de folha de pagamento. Caso na lista gerada determinado usuário do AD não seja listado o mesmo será automaticamente desativado.

---------------------

# Na prática

Por padrão ao executar o DisableDL-AD.vbs o script irá procurar um arquivo com o nome userlist.csv dentro de c:\ e desativar usuários que não estejam listados neste arquivo. Os usuários desativados vão ser listados em outro arquivo criado pelo script em c:\userlistdisable.csv.
Ambos arquivos supracitados podem ser alterados os nomes e diretórios editando as variaveis strCSV e strCSV2 no script.




