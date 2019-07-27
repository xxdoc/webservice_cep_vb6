# webservice_cep_vb6
Exemplo em Visual Basic de como consumir um WebService da Empresa Ipage Software.

A que se destina este Webservice?
Este Webservice tem por finalidade consultar Códigos de endereçamento Postal (CEP) de todo o Brasil de forma simples e descomplicada.
As informações retornadas após a consulta ao Webservice possui diversos formatos, são eles: XML, JSON, JavaScript, formato texto PIPED, formato texto Querty.
Definição dos parâmetros.
O CEP informado deve conter apenas números com até 08 (oito caracateres), em caso de valores inválidos passados ao Webservice o mesmo realizará automaticamente um filtro, deixando passar apenas números. Se mesmo assim o valor do CEP informado não satisfazer o critério uma mensagem de erro será reportada.

O formato a ser retornado pela consulta deve ser passado na URL junto com o CEP e deve ser compatível com o esperado pelo Webservice.
Os valores válidos são: XML, JSON, JavaScript, formato texto PIPED, formato texto Querty.

A chave de acesso ao Webservice é obrigatória e deve ser passada na URL junto com o CEP, formato de retorno e deve ser compatível com o esperado pelo Webservice. Caso não possua uma chave de acesso, solicite no site da ipage: https://www.ipage.com.br
