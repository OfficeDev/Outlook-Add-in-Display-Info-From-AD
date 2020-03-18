---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  platforms:
  - CSS
  createdDate: 8/17/2015 11:29:52 AM
---
# Outlook-Add-in-Display-Info-From-AD
Neste protótipo de suplemento de email, aprenda a acessar informações hierárquicas básicas do Active Directory (AD). Estenda esse protótipo para personalizar o suplemento de email da sua organização.

**Descrição do exemplo do suplemento de email do AD Quem é quem**

Este exemplo acompanha o tópico [Como: Criar um aplicativo de email para exibir informações de hierarquia do Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) no blog aplicativos do Office e do SharePoint.

Ao selecionar uma mensagem de email no Outlook ou no Outlook Web App, você pode escolher o suplemento de email do AD Quem é quem para exibir informações do Active Directory sobre o remetente e outros destinatários de uma mensagem de email atualmente selecionada no Outlook ou no Outlook Web App. O suplemento de email aparecerá na barra de aplicativos quando você estiver visualizando um email no Painel de Leitura ou no explorador de email.

Quando você escolhe esse suplemento de email pela primeira vez, ele recupera e exibe as informações hierárquicas e profissionais detalhadas do remetente do Active Directory: nome, cargo, departamento, alias, número do escritório, número de telefone e miniatura de uma imagem. Se o remetente tiver um gerente ou subordinados, o suplemento de email exibe um subconjunto similar de informações para cada um deles. A Figura 1 mostra um exemplo do suplemento do AD Quem é quem. A captura de tela mostra informações para Sara Melo, o gerente dela e os subordinados diretos.


![Figura 1. O aplicativo de email exibe informações do Active Directory para um remetente de email no Outlook](/description/image.png "O aplicativo de email do AD Quem é quem mostrando blocos para cada pessoa com foto, nome e cargo.")
 
O suplemento de email fornece uma barra de navegação que permite escolher um destinatário e visualizar informações profissionais e hierárquicas detalhadas armazenadas no Active Directory.

Nos bastidores, quando você seleciona um remetente ou destinatário, o suplemento email chama um serviço da Web chamado Quem, para obter os dados da pessoa no Active Directory. O serviço Web inclui um invólucro do Active Directory, que usa os Serviços de Domínio do Active Directory (AD DS) para acessar informações do Active Directory. Depois de obter os dados, o serviço Web Quem serializa os dados no formato JSON e os envia de volta como a resposta do serviço Web. O suplemento de email extrai os dados e exibe-os no painel de suplemento. A Figura 2 resume as relações entre o usuário do Outlook, o aplicativo de email, o serviço Web Quem e o Active Directory.

![Figura 2. Relações entre o usuário do Outlook, o aplicativo de email, o serviço Web Quem e o Active Directory](/description/75964e99-a74f-4fd3-a4e1-a0ca40bfc5a4image.png "O usuário interage com o aplicativo de email e o AD; o email interage com o usuário e o serviço Web; o serviço Web interage com o aplicativo de email e o AD.")

 
Confira o artigo fornecido [Como: Criar um aplicativo de email para exibir informações de hierarquia do Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) no blog aplicativos do Office e do SharePoint para obter uma descrição da implementação do suplemento email e do serviço Web Quem é quem.

**Anotação**

O serviço Web Quem serve apenas como um protótipo e mostra alguns recursos do Active Directory com os quais os usuários do Active Directory já estão familiarizados. Esperamos que este exemplo seja um bom ponto de partida para você estender e dar suporte a recursos específicos para sua organização. 
 
**Pré-requisitos**

Para aproveitar ao máximo esse exemplo de código, você deve estar familiarizado com desenvolvimento na Web com HTML e JavaScript, bem como com os serviços Web do Windows Communication Foundation (WCF). Não é necessário conhecer previamente os Serviços de Domínio do Active Directory.

Veja a seguir os requisitos para instalar e executar qualquer suplemento de email, inclusive o suplemento de email do AD Quem é quem:

* A caixa de correio do usuário deve estar no Exchange Server 2013 ou em uma versão posterior.
* O aplicativo de email deve ser executado no Outlook 2013 ou em uma versão posterior ou no Outlook Web App.

Você pode usar qualquer ferramenta de desenvolvimento Web com a qual esteja familiarizado para desenvolver o aplicativo de email do AD Quem é quem.

As ferramentas a seguir foram usadas para desenvolver o serviço Web Quem e implantar o aplicativo de email e o serviço Web:

* Visual Studio 2012
* .NET Framework 4.0
* Windows Server 2008
* Internet Information Server (IIS) 7.0

**Componentes principais do exemplo**

O download desse exemplo é formado pelos seguintes arquivos e pastas:

* Manifest.xml é um arquivo de manifesto para o suplemento de email do AD Quem é quem.
* WhoMailApp.sln é a solução do Visual Studio para todo o exemplo.
* A pasta WhoAgave contém arquivos do suplemento de email (inclusive arquivos CSS, imagens e HTML), e alguns dos arquivos para o serviço Web Quem.
* A pasta ActiveDirectory contém arquivos do invólucro do Active Directory.
* A pasta BuildProcessTemplates contém arquivos de modelo de marcação para desenvolvimento de serviços Web WCF.

**Configurar o exemplo**

Use as etapas a seguir para obter arquivos e modificar as referências deles, conforme apropriado:

1. Em uma unidade local, por exemplo d:, crie uma pasta chamada WhosWhoAD e baixe os arquivos de exemplo dela.
2. Supondo que o servidor Web do IIS no qual você pretende hospedar o aplicativo de email do AD Quem é quem se chame webserver, crie uma pasta chamada WhosWhoAD em \\webserver\c$\inetpub\wwwroot.
3. Copie a pasta img e o conteúdo dela de d:\WhosWhoAD\WhoAgave\ para \\webserver\c$\inetpub\wwwroot\WhosWhoAD.
O suplemento de email restante e os arquivos do serviço Web serão copiados de maneira apropriada para o webserver quando você implantar o serviço, conforme descrito na seção Implantar o serviço Web abaixo.
4. Atualize o arquivo de manifesto para refletir o local real do arquivo HTML do suplemento de email.

O arquivo de manifesto de suplemento de email, manifest.xml, está diretamente em d:\WhosWhoAD. Se seu servidor Web real tiver um nome diferente do webserver, atualize o manifest.xml para refletir a localização real do arquivo WhoMailApp.html, substituindo o webserver na linha a seguir com o caminho da pasta WhosWhoAD criada na Etapa 2.

```XML
<SourceLocation DefaultValue="https://webserver/WhosWhoAD/WhoMailApp.html"/>
 ```

**Instalar o suplemento de email**

1. No cliente avançado do Outlook, selecione Arquivo, Gerenciar Aplicativos. Isso abrirá um navegador para você se conectar ao Outlook Web App para acessar o Centro de Administração do Exchange (EAC).
2. Faça o logon na sua conta do Exchange.
3. No EAC, selecione a caixa do menu suspenso ao lado do botão + e selecione Adicionar do arquivo, conforme mostrado na Figura
![Figura 3. Instalar um aplicativo de email de um arquivo no Centro de Administração do Exchange](/description/f7a57314-42f1-4d15-9752-60e45ade98c3image.png "O botão do menu com o sinal de mais mostrando as opções Adicionar da Office Store, Adicionar da URL e Adicionar do arquivo.")

4. Na caixa de diálogo Adicionar do arquivo, navegue até o local de manifest.xml em d:\WhosWhoAD, selecione Abrir e, em seguida, Avançar.
Você verá então o suplemento do AD Quem é quem na lista de suplementos do Outlook, como mostrado na Figura 4.
![Figura 4. Aplicativo do AD Quem é quem instalado no Centro de Administração do Exchange](/description/3977704e-6d30-4067-bb17-e5d1f778795cimage.png "O aplicativo do AD Quem é quem listado na página de aplicativos para Outlook, mostrado como habilitado para o usuário.")

5. Se o Outlook estiver sendo executado, feche-o e abra-o novamente.

**Observação**
Esse procedimento será aplicável somente se a sua conta do Outlook estiver no Exchange Server 2013 ou uma versão posterior.
 
Além disso, na Etapa 3, se você não vir Adicionar do arquivo como uma opção, é necessário solicitar que o administrador do Exchange forneça as permissões necessárias para você.

O administrador do Exchange pode executar o cmdlet do PowerShell a seguir para atribuir as permissões necessárias a um único usuário. Neste exemplo, laurac é o alias de email do usuário.

```POWERSHELL
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
 ```

Se necessário, o administrador pode executar o cmdlet a seguir para atribuir permissões similares para vários usuários:

```POWERSHELL
$users = Get-Mailbox *
$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
 ```

Para saber mais sobre a função Meus Suplementos Personalizados, confira [Função Meus Suplementos Personalizados](http://msdn.microsoft.com/library/aa0321b3-2ec0-4694-875b-7a93d3d99089(Office.15).aspx).

**Implantar o serviço Web**

Faça o seguinte para implantar o serviço Web Quem e o arquivo de suplemento de email WhoMailApp.html:

1. No Visual Studio, abra WhoWebService.csproj.
2. Escolha Criar, Publicar WhoWebService.
3. Na guia Perfil da caixa de diálogo Publicar Web, especifique um perfil da sua preferência.
4. Na guia conexão, escolha Sistema de Arquivos como o como o Método Publish.
5. Digite \\webserver\c$\inetpub\wwwroot\WhosWhoAD como o local de Destino.
6. Escolha Publicar.
7. No computador do webserver, inicie o Gerenciador do IIS.
8. No painel Conexões, escolha Sites, Site da Web padrão.
9. Clique com o botão direito do mouse na pasta WhosWhoAD e escolha Converter em Aplicativo.
10. Na caixa de diálogo Adicionar Aplicativo, em pool de aplicativos com o DefaultAppPool listado por padrão, escolha Selecionar.
11. Na caixa de diálogo Selecionar Pool de Aplicativos, em Propriedades, verifique se .Net Framework Versão: 4.0 ou uma versão posterior do .NET Framework é exibida. Escolha um pool de aplicativos diferente, se necessário, para garantir que o pool use pelo menos o .NET Framework 4.0. Escolha OK.
12. Na caixa de diálogo Adicionar aplicativo, verifique se é possível ver a Autenticação de passagem, conforme mostrado na Figura 6. Escolha OK. Prossiga para a Etapa 14.
![Figura 5. Caixa de diálogo Adicionar Aplicativo para converter o serviço Web Quem como um aplicativo no pool de aplicativos apropriado no ISS](/description/584495fd-6e4d-45da-9b5b-47d114848342image.png "A caixa de diálogo Adicionar Aplicativo exibindo a Autenticação de passagem de texto.")

13. Como alternativa às etapas 10 a 12, você pode criar um novo pool de aplicativos que use o .NET Framework 4.0 (ou uma versão posterior) e a autenticação de passagem. Selecione esse pool de aplicativos e passe para a Etapa 14.
14. No painel do meio do Gerenciador do IIS, escolha Autenticação. Verifique se a Autenticação do Windows está habilitada; clique com o botão direito para habilitá-la, se necessário.

O procedimento de implantação copia o seguinte arquivo para \\webserver\c$\inetpub\wwwroot\WhosWhoAD\:
* bin\ActiveDirectoryWrapper.dll
* bin\WhoWebService.dll
* css\WhoMailApp.css
* img\anonymous.jpg
* img\app_icon.png
* img\envelop.png
* img\telephone.png
* Web.config
* WhoMailApp.html
* WhoService.svc

O serviço Web Quem agora pode ser acessado no webserver e você pode usar o aplicativo de email do AD Quem é quem no Outlook ou no Outlook Web App.

**Executar e testar o exemplo**

1. No Outlook, escolha um email para ler no Painel de Leitura.
2. Escolha o suplemento de email do AD Quem é quem na barra de suplementos.

Você verá ver os dados do Active Directory no painel do suplemento, semelhante ao exemplo na Figura 1.

<a name="Troubleshooting"></a>
**Solução de problemas**

Se o remetente ou destinatário de uma mensagem de email tiver um endereço de email do formulário <first name>.<last name>@<domain>, o invólucro do Active Directory pode não ser capaz de procurar a pessoa adequada no Active Directory. Escolha uma pessoa cujo endereço de email seja simplesmente da forma <alias>@<domain>.

Como é a função do suplemento de email do AD Quem é quem servir como protótipo, é possível personalizar o suplemento de email para atender aos requisitos da sua organização. Confira mais informações na seção de Extensão futura no artigo fornecido para ter mais informações.

**Conteúdo relacionado**

* [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Tutorial: Criar um aplicativo de email para exibir informações de hierarquia do Active Directory](https://blogs.msdn.microsoft.com/officeapps/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients/)


Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
