# webservice_cep_vb6
Exemplo em Visual Basic de como consumir um Web Service da Empresa Ipage Software.

A que se destina este Web Service?
Este Web Service tem por finalidade consultar Códigos de endereçamento Postal (CEP) de todo o Brasil de forma simples e descomplicada.
As informações retornadas após a consulta ao Web Service possui diversos formatos, são eles: XML, JSON, JavaScript, formato texto PIPED, formato texto Querty.
Definição dos parâmetros.
O CEP informado deve conter apenas números com até 08 (oito caracateres), em caso de valores inválidos passados ao Web Service o mesmo realizará automaticamente um filtro, deixando passar apenas números. Se mesmo assim o valor do CEP informado não satisfazer o critério uma mensagem de erro será reportada.

O formato a ser retornado pela consulta deve ser passado na URL junto com o CEP e deve ser compatível com o esperado pelo Web Service.
Os valores válidos são: XML, JSON, JavaScript, formato texto PIPED, formato texto Querty.

A chave de acesso ao Web Service é obrigatória e deve ser passada na URL junto com o CEP, formato de retorno e deve ser compatível com o esperado pelo Web Service. Caso não possua uma chave de acesso, solicite no site da ipage: https://www.ipage.com.br

O projeto é bastante rico em técnicas de programação.
São elas:
1 - Uso de classes
2 - Uso de JSON
3 - Uso de API's do Windows
4 - Uso de requisições em XML
5 - Técnicas de menus em Janelas MDI
6 - Documentação de métodos
7 - Uso de componentes ActiveX com o CreateObject (onde não há a necessidade de referenciar no projeto), a garga é IN-LINE
