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
このプロトタイプのメール アドインを使用して Active Directory (AD) からの基本的な階層情報にアクセスする方法について説明します。このプロトタイプを拡張することで、組織用のメール アドインをカスタマイズすることができます。

** Who's Who AD メール アドイン サンプルの説明**

このサンプルは、Apps for Office and SharePoint Blog (Office および SharePoint 向けアプリ ブログ) のトピック 「[操作方法:メール アプリを作成して Active Directory からの階層情報を表示する](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx)」に付属するものです。

Outlook または Outlook Web App でメール メッセージを選択するときに Who's Who AD メール アドインを選択することで、Outlook または Outlook Web App で現在選択されているメール メッセージの送信者および他の受信者に関する Active Directory 情報を表示することができます。このメール アドインは、メールを閲覧ウィンドウまたはメール エクスプローラーで表示しているときにアプリ バーに表示されます。

このメール アドインを初めて選択すると、Active Directory からの送信者に関する詳細な職務および階層情報 (名前、役職、部署、通称、職場電話番号、電話番号、サムネイル画像) がアドインにより取得されて表示されます。このほかに、送信者に上司または直属の部下がいる場合、メール アドインによりそれぞれのユーザーについて同様の情報のサブセットが表示されます。図 1 に Who's Who AD アドインの例を示します。Belinda Newman、彼女の上司、および直属の部下に関する情報を表示するスクリーンショット。


![図 1.メール送信者の Active Directory 情報を Outlook に表示するメール アプリ](/description/image.png "各ユーザーのタイルを写真、名前、役職とともに表示している Who's Who AD メール アプリ。")
 
このメール アドインで提供されるナビゲーション バーを使用することにより、受信者を選択して Active Directory に保存されている詳細な職務および階層情報を表示することができます。

バックグラウンドでは、送信者または受信者が選択されると、メール アドインが "Who" という名前の Web サービスを呼び出してそのユーザーのデータを Active Directory から取得します。この Web サービスには Active Directory ラッパーが含まれ、これは Active Directory ドメイン サービス (AD DS) を使用して Active Directory からの情報にアクセスします。データを取得後、Who Web サービスはデータを JSON 形式でシリアル化し、Web サービス応答としてデータを送り返します。メール アドインによりデータが取得され、アドイン ウィンドウにデータが表示されます。図 2 は、Outlook ユーザー、メール アプリ、Who Web サービス、および Active Directory の間の関係をまとめたものです。

![図 2.Outlook ユーザー、メール アプリ、Who Web サービス、および Active Directory の間の関係](/description/75964e99-a74f-4fd3-a4e1-a0ca40bfc5a4image.png "ユーザーがメール アプリおよび AD とやり取りする; メールがユーザーおよび Web サービスとやり取りする; Web サービスがメール アプリおよび AD とやり取りする。")

 
メール アドインおよび Who Web サービスの実装に関する説明については、Apps for Office and SharePoint Blog (Office および SharePoint 向けアプリ ブログ) の関連記事「[操作方法:メール アプリを作成して Active Directory からの階層情報を表示する](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx)」を参照してください。

**注**

Who Web サービスはプロトタイプとして利用することのみを意図したもので、ほとんどの Active Directory ユーザーが使い慣れている、Active Directory のいくつかの基本機能を使用しています。このサンプルは、お客様の組織に固有の機能を拡張してサポートするための土台としてお使いください。 
 
**前提条件**

このコード サンプルを有効的に活用するには、HTML と JavaScript を使用した Web 開発および Windows Communication Foundation (WCF) Web サービスに精通している必要があります。Active Directory ドメイン サービスに関する予備知識は必要ありません。

Who's Who AD メール アドインを含む、すべてのメール アドインをインストールして実行するための要件は次のとおりです。

* ユーザーのメールボックスが Exchange Server 2013 以降のバージョン上にある必要があります。
* メール アプリは Outlook 2013 以降のバージョンまたは Outlook Web App で実行する必要があります。

Who's Who AD メール アプリの開発には、任意の使い慣れた Web 開発ツールを使用できます。

Who's Who Web サービスの開発およびメール アプリと Web サービスの展開には、次のツールが使用されました。

* Visual Studio 2012
* .NET Framework 4.0
* Windows Server 2008
* Internet Information Server (IIS) 7.0

**サンプルの主要なコンポーネント**

このサンプルのダウンロードは、次のファイルとフォルダーによって構成されています。

* Manifest.xml は、Who's Who AD メール アドイン用のマニフェスト ファイルです。
* WhoMailApp.sln は、本サンプル全体の Visual Studio ソリューションです。
* WhoAgave フォルダーには、メール アドイン用のファイル (HTML、画像、CSS ファイルを含む) のほか、Who Web サービス用の一部のファイルが含まれています。
* ActiveDirectory フォルダーには、Active Directory ラッパー用のファイルが含まれています。
* BuildProcessTemplates フォルダーには、WCF Web サービスを開発するための、既定のマークアップ テンプレート ファイルが含まれています。

**サンプルを構成する**

次の手順を使用してファイルを取得し、必要に応じて参照を変更します。

1. ローカル ドライブ (d: など) で、"WhosWhoAD" というフォルダーを作成し、サンプル ファイルをこのフォルダーにダウンロードします。
2. Who's Who AD メール アプリのホストとして使用する IIS Web サーバーの名前が "webserver" であると仮定して、"WhosWhoAD" というフォルダーを \\webserver\\c$\\inetpub\\wwwroot に作成します。
3. img フォルダーとその中身を d:\WhosWhoAD\WhoAgave\ から \\webserver\c$\inetpub\wwwroot\WhosWhoAD にコピーします。
下記の「Web サービスを展開する」セクションで説明するとおり、メール アドインと Web サービスの残りのファイルは、Web サービスを展開した際に webserver に適切にコピーされます。
4. マニフェスト ファイルを更新して、メール アドイン HTML ファイルの実際の場所を反映させます。

メール アドインのマニフェスト ファイル manifest.xml は、 d:\WhosWhoAD のすぐ下にあります。実際の Web サーバーの名前が "webserver" とは異なる場合、manifest.xml を更新して WhoMailApp.html ファイルの実際の場所を反映させます。これを行うには、次の行の "webserver" を、手順 2 で作成した WhosWhoAD フォルダーのサーバー パスに置き換えます。

```XML
<SourceLocation DefaultValue="https://webserver/WhosWhoAD/WhoMailApp.html"/>
 ```

**メール アドインのインストール**

1. Outlook リッチ クライアントで、[ファイル]、[アプリの管理] の順に選択します。Outlook Web App にログオンして Exchange 管理センター (EAC) に移動するために、ブラウザーが開きます。
2. Exchange アカウントにログオンします。
3. 図 3 に示されるように、EAC で [+] ボタンの横にあるドロップダウン ボックスを選択し、[ファイルから追加] を選択します。
![図 3.Exchange 管理センターで、メール アプリをファイルからインストールする](/description/f7a57314-42f1-4d15-9752-60e45ade98c3image.png "[Office Store から追加]、[URL から追加]、[ファイルから追加] のオプションを表示する [+] ボタン メニュー。")

4. [ファイルから追加] ダイアログ ボックスで、d:\\WhosWhoAD 内の manifest.xml の場所を参照し、[開く] を選択し、[次へ] を選択します。
図 4 に示すとおり、Outlook 用アドインの一覧に Who's Who AD アドインが表示されるはずです。
![図 4.Exchange 管理センターでインストールされた Who's Who AD アプリ](/description/3977704e-6d30-4067-bb17-e5d1f778795cimage.png "[Outlook 用アプリ] ページに一覧記載される、ユーザーに対して有効になっていると表示される Who's Who AD アプリ")

5. Outlook が実行中の場合は Outlook を閉じ、もう一度開きます。

**注**:
この手順は Outlook アカウントが Exchange Server 2013 以降のバージョン上にある場合にのみ当てはまります。
 
また、手順 3 で、選択肢として [ファイルから追加] が表示されない場合、Exchange 管理者に要求して必要なアクセス許可を付与してもらう必要があります。

Exchange 管理者は、次の PowerShell コマンドレットを実行して、必要なアクセス許可をユーザー 1 名に割り当てることができます。この例では、"wendyri" はユーザーのメール エイリアスです。

```POWERSHELL
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
 ```

必要な場合、管理者は次のコマンドレットを実行して、同様のアクセス許可を複数のユーザーに割り当てることができます。

```POWERSHELL
$users = Get-Mailbox *
$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
 ```

My Custom Apps ロールの詳細については、「[My Custom Apps ロール](http://msdn.microsoft.com/library/aa0321b3-2ec0-4694-875b-7a93d3d99089(Office.15).aspx)」を参照してください。

**Web サービスを展開する**

次の操作を実行して、Who Web サービスと WhoMailApp.html メール アドイン ファイルを展開します。

1. Visual Studio で、WhoWebService.csproj を開きます。
2. [ビルド]、[WhoWebService の発行] の順に選択します。
3. [Web を発行] ダイアログ ボックスの [プロファイル] タブで、任意のプロファイルを指定します。
4. [接続] タブで、 [発行方法] として [ファイル システム] を選択します。
5. [ターゲット] の場所として、「\\webserver\c$\inetpub\wwwroot\WhosWhoAD」と入力します。
6. [発行] を選択します。
7. webserver コンピューターで、IIS マネージャーを起動します。
8. [接続] ウィンドウで、[サイト]、[既定の Web サイト] の順に選択します。
9. WhosWhoAD フォルダーを右クリックし、[アプリケーションへの変換] を選択します。
10. [アプリケーションの追加] ダイアログ ボックスで、既定値として "DefaultAppPool" が表示されている [アプリケーション プール] で、[選択] を選択します。
11. [アプリケーション プールの選択] ダイアログ ボックスの [プロパティ] で、.Net Framework のバージョンを確認します。.NET Framework バージョン 4.0 以降が表示されていることを確認します。プールで使用バージョンが最低でも .NET Framework 4.0 となるよう、必要に応じて別のアプリケーション プールを選択します。[OK] を選択します。
12. 図 6 に示すように、[アプリケーションの追加] ダイアログ ボックスに "パススルー認証" と表示されていることを確認します。[OK] を選択します。手順 14 に進んでください。
![図 5.IIS で、適切なアプリケーション プールで Who Web サービスをアプリケーションとして変換するための [アプリケーションの追加] ダイアログ ボックス](/description/584495fd-6e4d-45da-9b5b-47d114848342image.png ""パススルー認証" というテキストを表示する [アプリケーションの追加] ダイアログ ボックス。")

13. 手順 10 から 12 の代わりに、.NET Framework 4.0 (またはそれ以降のバージョン) およびパススルー認証を使用する新しいアプリケーション プールを作成することも可能です。アプリケーション プールを選択し、手順 14 に進みます。
14. IIS マネージャーの中央のウィンドウで、[認証] を選択します。Windows 認証が有効になっていることを確認します。必要に応じて、右クリックして有効にします。

展開プロシージャでは、次のファイルが \\webserver\c$\inetpub\wwwroot\WhosWhoAD にコピーされます:
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

これで、webserver 上で Who Web サービスにアクセスできるようになり、Who's Who AD メール アプリ を Outlook または Outlook Web App で使用できるようになりました。

**サンプルを実行してテストする**

1. Outlook で、メールを 1 通選択して閲覧ウィンドウに表示します。
2. アドイン バーで [Who's Who AD] メール アドインを選択します。

図 1 の例のように、アドイン ウィンドウに Active Directory データが表示されるはずです。

<a name="Troubleshooting"></a>
**トラブルシューティング**

メール メッセージの送信者または受信者の使用するメール アドレスが <first name>.<last name>@<domain> という形になっている場合、Active Directory ラッパーは正しいユーザーを Active Directory で検索できない可能性があります。メール アドレスが <alias>@<domain> という単純な形のユーザーを選択してください。

Who's Who AD メール アドインはプロトタイプとして使用することを意図したものであるため、組織の要件に合わせてメール アドインをカスタマイズする余地があります。詳細については、関連記事の「Future extension (今後の拡張予定)」セクションを参照してください。

**関連コンテンツ**

* [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [操作方法:メール アプリを作成して Active Directory からの階層情報を表示する](https://blogs.msdn.microsoft.com/officeapps/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients/)


このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
