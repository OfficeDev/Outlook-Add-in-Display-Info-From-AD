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
Ce complément de messagerie prototype décrit comment accéder aux informations de hiérarchie de base à partir d’Active Directory (AD). Développez ce prototype pour personnaliser le complément de messagerie pour votre organisation.

**Description de l’exemple de complément de messagerie publicitaire**

Cet exemple accompagne la rubrique [Comment : Créez une application de messagerie pour afficher les informations de hiérarchie d’Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) dans les applications pour le blog Office et SharePoint.

Lorsque vous sélectionnez une adresse e-mail dans Outlook ou Outlook Web App, vous pouvez choisir les utilisateurs qui ont accès à la messagerie Active Directory pour afficher les informations relatives à l’expéditeur et à d’autres destinataires d’un message électronique actuellement sélectionné dans Outlook Web App. Le complément de messagerie apparaît dans la barre d'application lorsque vous consultez un e-mail dans le volet Lecture ou dans l'explorateur de courrier.

Lorsque vous sélectionnez le complément de messagerie pour la première fois, il récupère et affiche les informations professionnelles et hiérarchiques détaillées de l’expéditeur à partir d’Active Directory (nom, fonction, service, alias, numéro de téléphone, numéro de téléphone et image miniature). Si l’expéditeur dispose d’un responsable ou de subordonnés directs, le complément de messagerie affiche également un sous-ensemble d’informations similaire pour chacun d’eux. La figure 1 montre un exemple de complément Who's Who AD. La capture d’écran affiche les informations pour Belinda Newman, son responsable et les subordonnés directs.


![Figure 1. L’application de messagerie affiche les informations Active Directory pour un expéditeur d’e-mail dans Outlook](/description/image.png "L’application de messagerie Who's Who AD affichant les vignettes de chaque personne avec photo, nom et fonction.")
 
Le complément de messagerie fournit une barre de navigation qui vous permet de choisir un destinataire et d’afficher des informations détaillées sur les professionnels et la hiérarchie stockés dans Active Directory.

En coulisses, lorsque vous sélectionnez un expéditeur ou un destinataire, le complément de messagerie appelle un service web appelé Who, pour obtenir les données de la personne à partir d’Active Directory. Le service web inclut un wrapper Active Directory, lequel utilise Active Directory Domain Services (AD DS) pour accéder aux informations à partir d’Active Directory. Une fois que vous avez obtenu les données, le service web Who sérialise les données au format JSON et les renvoie en tant que réponse de service web. Le complément de messagerie extrait ensuite les données et les affiche dans le volet de complément. La figure 2 résume les relations entre l’utilisateur Outlook, l’application de messagerie, le service web Who et Active Directory.

![Figure 2. Les relations entre l’utilisateur Outlook, l’application de messagerie, le service web Who et Active Directory](/description/75964e99-a74f-4fd3-a4e1-a0ca40bfc5a4image.png "L’utilisateur interagissent avec l’application de messagerie et AD. le courrier interagit avec l’utilisateur et le service web. le service web interagit avec l’application de messagerie et AD.")

 
Pour plus d’informations, consultez l’article qui accompagne [Comment : Créez une application de messagerie pour afficher les informations de hiérarchie d’Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) dans les applications pour le blog Office et SharePoint pour obtenir une description de l’implémentation du complément courrier et le service web Who.

**Note**

Le service web Who sert uniquement de prototype et présente quelques fonctionnalités de base d’Active Directory familières à la plupart des utilisateurs d’Active Directory. Nous espérons que cet exemple constitue un bon point de départ pour étendre et prendre en charge les fonctionnalités propres à votre organisation. 
 
**Conditions préalables**

Pour tirer le meilleur parti de cet exemple de code, vous devez être familiarisé au développement web en utilisant HTML et JavaScript, ainsi que des services web WCF (Windows Communication Foundation). Il n'est pas nécessaire d'avoir des connaissances préalables Active Directory Domain Services.

Les conditions suivantes sont requises pour installer et exécuter les compléments de messagerie, y compris le complément de messagerie Who's Who AD :

* La boîte aux lettres de l’utilisateur doit se trouver sur Exchange Server 2013 ou version ultérieure.
* L’application de messagerie doit être exécutée sur Outlook 2013 ou une version ultérieure, ou Outlook Web App.

Vous pouvez utiliser n’importe quel outil de développement web qui vous est familier pour développer les utilisateurs de l’application de messagerie Who's Who AD.

Les outils suivants ont été utilisés pour développer le service web et déployer l’application de messagerie et le service web :

* Visual Studio 2012
* .NET Framework 4.0
* Windows Server 2008
* Internet Information Server (IIS) 7.0

**Composants clés de l’exemple**

Le téléchargement de cet exemple comprend les fichiers et les dossiers suivants :

* Manifest.xml est le fichier manifeste pour les personnes qui ont accès au complément de messagerie Who's Who AD.
* WhoMailApp.sln est le fichier solution Visual Studio pour l’exemple complet.
* Le dossier WhoAgave contient les fichiers du complément de messagerie (y compris les fichiers HTML, images et CSS), ainsi que certains des fichiers du service web.
* Le dossier ActiveDirectory contient les fichiers pour le wrapper Active Directory.
* Le dossier BuildProcessTemplates contient des fichiers de modèle de balisage par défaut pour le développement des services web WCF.

**Configurer l’exemple**

Procédez comme suit pour récupérer les fichiers et modifier leurs références, le cas échéant :

1. Sur un lecteur local (par exemple, d:), créez un dossier nommé WhosWhoAD et téléchargez les exemples de fichier.
2. En supposant que le serveur web IIS sur lequel vous envisagez d’héberger l'application de messagerie Who's Who AD est appelée webserver, créez un dossier nommé WhosWhoAD sous \\webserver\c$\inetpub\wwwroot.
3. Copiez le dossier img et son contenu de d:\\WhosWhoAD\WhoAgave\ to \\webserver\c$\inetpub\wwwroot\WhosWhoAD.
Les fichiers du complément de messagerie et du service web restants seront copiés de façon appropriée sur le serveur web lorsque vous déployez le service web, comme décrit dans la section déployer le service web ci-dessous.
4. Mettez à jour le fichier manifeste pour refléter l’emplacement réel du fichier HTML du complément de messagerie.

Le fichier manifeste du complément de messagerie, manifest.xml, est directement sous d:\\WhosWhoAD. Si le nom de votre serveur web réel est différent de celui du serveur web, mettez à jour manifest.xml pour refléter l’emplacement réel du fichier WhoMailApp.html, en remplaçant le fichier webserver dans la ligne suivante par le chemin d’accès au serveur du dossier WhosWhoAD que vous avez créé à l’étape 2.

```XML
<SourceLocation DefaultValue="https://webserver/WhosWhoAD/WhoMailApp.html"/>
 ```

**Installation du complément de messagerie**

1. Dans le client enrichi Outlook, sélectionnez fichier, gérer les applications. Un navigateur s’ouvre alors pour vous connecter à Outlook Web App pour accéder au centre d’administration Exchange (EAC).
2. Connectez-vous à votre compte Exchange.
3. Dans le centre d’administration Exchange, sélectionnez la zone de liste déroulante adjacente au bouton +, puis sélectionnez Ajouter à partir d’un fichier, comme illustré dans la Figure
![Figure 3. L’installation d’une application de messagerie à partir d’un fichier dans le centre d’administration Exchange](/description/f7a57314-42f1-4d15-9752-60e45ade98c3image.png "Le menu du bouton plus affichant les options Ajouter à partir d’Office Store, Ajouter à partir de l’URL et Ajouter à partir du fichier.")

4. Dans la boîte de dialogue Ajouter à partir du fichier, accédez à l’emplacement du fichier manifest.xml dans le dossier d:\\WhosWhoAD, sélectionnez Ouvrir, puis cliquez sur Suivant.
Vous devez ensuite voir le complément Who's Who AD dans la liste des compléments pour Outlook, comme illustré dans la Figure 4.
![Figure 4. L’application Who's Who AD installée dans le centre d’administration Exchange](/description/3977704e-6d30-4067-bb17-e5d1f778795cimage.png "L’application Who's Who AD répertoriée dans la page applications pour Outlook, indiquée comme activée pour l’utilisateur.")

5. Si Outlook est en cours d’exécution, fermez et rouvrez Outlook.

**Remarque**
cette procédure s’applique uniquement si votre compte Outlook utilise Exchange Server 2013 ou une version ultérieure.
 
De plus, à l’étape 3, si vous ne voyez pas l’option Ajouter à partir d’un fichier, vous devez demander à votre administrateur Exchange de vous fournir les autorisations nécessaires.

L’administrateur Exchange peut exécuter la cmdlet PowerShell suivante pour affecter les autorisations nécessaires à un seul utilisateur. Dans cet exemple, wendyri est l’alias de messagerie de l’utilisateur.

```POWERSHELL
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
 ```

Selon les besoins, l’administrateur peut exécuter la cmdlet suivante pour attribuer des autorisations similaires à plusieurs utilisateurs :

```POWERSHELL
$users = Get-Mailbox *
$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
 ```

Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](http://msdn.microsoft.com/library/aa0321b3-2ec0-4694-875b-7a93d3d99089(Office.15).aspx).

**Déployer le service web**

Procédez comme suit pour déployer le service web et le fichier de complément de messagerie WhoMailApp.html :

1. Dans Visual Studio, ouvrez WhoWebService.csproj.
2. Sélectionnez Build, Publier WhoWebService.
3. Dans l’onglet Profil de la boîte de dialogue web Publier, spécifiez un profil de votre choix.
4. Dans l’onglet connexion, choisissez Système de fichiers comme méthode de Publication.
5. Tapez \\webserver\c$\inetpub\wwwroot\WhosWhoAD comme emplacement cible.
6. Choisissez Publier.
7. Sur l’ordinateur webserver, démarrez le gestionnaire des services Internet.
8. Dans le volet Connections, sélectionnez sites, site web par défaut.
9. Cliquez avec le bouton droit sur le dossier WhosWhoAD, puis sélectionnez Convertir en application.
10. Dans la boîte de dialogue Ajouter une application, sous pool d’applications avec la valeur DefaultAppPool indiquée par défaut, choisissez Sélectionner.
11. Dans la boîte de dialogue Sélectionner le pool d’applications, sous Propriétés, assurez-vous que la version du .NET Framework : 4.0 ou une version ultérieure du .NET Framework est affichée. Si nécessaire, choisissez un autre pool d’applications pour vous assurer que le pool utilise au moins .NET Framework 4.0. Cliquez sur OK.
12. Dans la boîte de dialogue Ajouter une application, assurez-vous que l’option authentification directe est affichée, comme illustré dans la Figure 6. Cliquez sur OK. Passez à l'étape 14.
![Figure 5. Boîte de dialogue Ajouter une application pour convertir le service web Who en application dans le pool d’applications approprié sur IIS](/description/584495fd-6e4d-45da-9b5b-47d114848342image.png "La boîte de dialogue Ajouter une application affichant l’authentification par transfert de texte.")

13. En guise d’alternative aux étapes 10 à 12, vous pouvez créer un nouveau pool d’applications qui utilise .NET Framework 4.0 (ou une version ultérieure) et une authentification directe. Sélectionnez ce pool d’applications et passez à l’étape 14.
14. Dans le volet central du gestionnaire des services Internet, sélectionnez Authentification. Vérifiez que l’authentification Windows est activée. Cliquez avec le bouton droit pour l’activer, le cas échéant.

La procédure de déploiement copie les fichiers suivants sur \\webserver\c$\inetpub\wwwroot\WhosWhoAD\ :
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

Vous pouvez désormais accéder au service web sur webserver. Vous pouvez désormais utiliser l’application de messagerie Who's Who AD dans Outlook ou Outlook Web App.

**Exécuter et tester le complément**

1. Dans Outlook, sélectionnez un e-mail à lire dans le Volet lecture.
2. Choisissez le complément de messagerie Who's Who AD dans la barre des compléments.

Vous devez être en mesure de voir les données Active Directory dans le volet complément, comme dans l’exemple de la Figure 1.

<a name="Troubleshooting"></a>
**Résolution des problèmes**

Si l’expéditeur ou le destinataire d’un message électronique a une adressee-mail de la forme <first name>.<last name>@<domain>, il est possible que le wrapper Active Directory ne puisse pas rechercher la personne appropriée dans Active Directory. Choisissez une personne dont l’adresse de messagerie est simplement de la forme <alias>@<domain>.

Le complément de messagerie Who's Who AD étant destiné à servir de prototype, vous pouvez le personnaliser pour l'adapter aux besoins de votre organisation. Pour plus d’informations, consultez la section Extension ultérieure de l’article joint.

**Contenu associé**

* [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Procédure : Créez une application de messagerie pour afficher les informations de hiérarchie d’Active Directory](https://blogs.msdn.microsoft.com/officeapps/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients/)


Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
