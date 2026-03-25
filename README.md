Cenário: Para exportação de relatórios de permissionamento externo em alguns sites específicos do SharePoint Online. O escopo do cliente é obter todas as permissões de Site/Subsite, Bibliotecas e Pastas/Subpastas.
Durante as buscas e o desenvolvimento do script, nos deparamos com um ponto que está bloqueando a evolução:
•	Já conseguimos ler e exportar as permissões de usuários externos em Site/Subsite e Bibliotecas.
•	Porém, ao tentar cobrir a camada de Pastas/Subpastas, o script consegue identificar a primeira camada de pastas, mas não está conseguindo exportar os usuários com permissão e não consegue percorrer/ler as subpastas dentro da estrutura do site.
Você poderia nos apoiar para evoluirmos esse ponto, envolvendo algum SME (ou alguém com experiência em permissões/ACL em SharePoint Online) para conseguirmos concluir o report dentro do prazo?
Se fizer sentido, posso compartilhar o script atual e um exemplo de site onde o comportamento ocorre para facilitar a análise.


*** Ambiente de Testes

Connect-PnPOnline -Url "https://brtechs-admin.sharepoint.com/" -ClientId "fa70bcfc-402a-4156-8634-c8b8cd772bfb" -Tenant "5ef881c6-1fa8-4d9c-ad81-634a9d5f9d07" -Thumbprint "EEED6DA026421085D8DB922A7973EE913CE26645"
Site Collection: https://brtechs.sharepoint.com/sites/Brasil/
