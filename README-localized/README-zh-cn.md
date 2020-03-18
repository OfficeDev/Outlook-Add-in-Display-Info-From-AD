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
通过此原型邮件加载项了解如何从 Active Directory (AD) 访问基本层次结构信息。扩展此原型，以自定义组织的邮件加载项。

**Who's Who AD 邮件加载项示例描述**

本示例附带介绍了 Office 应用程序和 SharePoint 博客中的主题[操作方法：创建邮件应用以通过 Active Directory 显示层次结构信息](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx)。

在 Outlook 或 Outlook Web App 中选择电子邮件时，可以选择 Who's Who AD 邮件加载项，以显示与当前在 Outlook Web App 中选定的电子邮件发件人和其他收件人相关的 Active Directory 信息。在阅读窗格或邮件资源管理器中查看电子邮件时，邮件加载项将显示在应用栏中。

首次选择此邮件加载项时，它将通过 Active Directory 检索和显示发件人的详细专业和层次结构信息，包括姓名、职务、部门、别名、办公室编号、电话号码和图片缩略图。如果发件人有经理或直接下属，则邮件加载项也会为每个人显示相似的信息子集。图 1 所示为 Who's Who AD 加载项示例。屏幕截图所示为 Belinda Newman、其经理和直接下属的相关信息。


![图 1：邮件应用将在 Outlook 中显示电子邮件发件人的 Active Directory 信息](/description/image.png "显示每个人员磁贴（包括图片、姓名和职务）的 Who's Who AD 邮件应用。")
 
通过邮件加载项中提供的导航栏，可以选择收件人并查看存储在 Active Directory 中的详细专业和层次结构信息。

选择发件人或收件人时，邮件加载项将在后台调用一项名为 Who 的 Web 服务，以获取 Active Directory 中的个人数据。该 Web 服务包括一个 Active Directory 包装，它使用 Active Directory 域服务 (AD DS) 来访问 Active Directory 中的信息。获取数据之后，Who Web 服务将以 JSON 格式序列化数据，并将其作为 Web 服务响应发回。邮件加载项随后将会拉取数据并将其显示在加载项窗格上。图 2 总结了 Outlook 用户、邮件应用、Who Web 服务与 Active Directory 之间的关系。

![图 2：Outlook 用户、邮件应用、Who Web 服务与 Active Directory 之间的关系](/description/75964e99-a74f-4fd3-a4e1-a0ca40bfc5a4image.png "用户与邮件应用和 AD 交互；邮件与用户和 Web 服务交互；Web 服务与邮件应用和 AD 交互。")

 
请参阅 Office 应用程序和 SharePoint 博客中的附带文章[操作方法：创建邮件应用以通过 Active Directory 显示层次结构信息](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx)，以获取实施邮件加载项和 Who Web 服务的描述。

**注意**

Who Web 服务仅用作原型，它用于显示大部分 Active Directory 用户熟悉的一些基本 Active Directory 功能。希望本示例能够为你提供一个好的起点，让你能够扩展和支持组织特定的功能。 
 
**先决条件**

为了充分利用本代码示例，你应熟悉如何使用 HTML 和 JavaScript 进行 Web 开发以及 Windows Communication Foundation (WCF) Web 服务。无需事先了解 Active Directory 域服务。

以下是安装和运行任何邮件加载项的要求，包括 Who's Who AD 邮件加载项：

* 用户的邮箱必须在 Exchange Server 2013 或更高版本上。
* 邮件应用必须运行 Outlook 2013 或更高版本，或者 Outlook Web App。

可以使用任何熟悉的 Web 开发工具来开发 Who's Who AD 邮件应用。

以下工具用于开发 Who Web 服务和部署邮件应用和 Web 服务：

* Visual Studio 2012
* .NET Framework 4.0
* Windows Server 2008
* Internet Information Server (IIS) 7.0

**示例的主要组件**

本示例下载中包含以下文件和文件夹：

* Manifest.xml 是 Who's Who AD 邮件加载项的清单文件。
* WhoMailApp.sln 是整个示例的 Visual Studio 解决方案文件。
* WhoAgave 文件夹包含用于邮件加载项文件（包括 HTML、图像和 CSS 文件）以及部分用于 Who Web 服务的文件。
* ActiveDirectory 文件夹包含用于 Active Directory 包装的文件。
* The BuildProcessTemplates 文件夹包含用于开发 WCF Web 服务的默认标记模板文件。

**配置示例**

请使用以下步骤获取文件并相应地修改其引用：

1. 在本地驱动器上，例如 d:，创建一个名为 WhosWhoAD 的文件夹并将示例文件下载到此处。
2. 假设用于托管 Who's Who AD 邮件应用的 IIS Web 服务器被称为 webserver，则在 \\webserver\c$\inetpub\wwwroot 下创建一个名为 WhosWhoAD 的文件夹。
3. 将图像文件夹机器内容从 d:\WhosWhoAD\WhoAgave\ 复制到 \\webserver\c$\inetpub\wwwroot\WhosWhoAD。
部署 Web 服务时，将其余的邮件加载项和 Web 服务文件相应地复制到 webserver，如下面的“部署 Web 服务”部分所述。
4. 更新清单文件，以反映邮件加载项 HTML 文件的实际位置。

邮件加载项清单文件 manifest.xml 就在 d:\WhosWhoAD 下。如果实际 Web 服务器的名称不同于 webserver，请更新清单文件以反映 WhoMailApp.html 文件的实际位置，方法是将以下行中的 webserver 替换为在步骤 2 创建的 WhosWhoAD 文件夹的服务器路径。

```XML
<SourceLocation DefaultValue="https://webserver/WhosWhoAD/WhoMailApp.html"/>
 ```

**安装邮件加载项**

1. 在 Outlook 富客户端中，依次选择“文件”、“管理应用”。此时将会打开一个浏览器，你可以在其中登录到 Outlook Web App 以转至 Exchange 管理中心 (EAC)。
2. 登录到 Exchange 帐户。
3. 在 EAC 中，选择 + 按钮旁边的下拉框，然后选择“从文件添加”，如图
![ 3 所示。在 Exchange 管理中心中从文件安装邮件应用](/description/f7a57314-42f1-4d15-9752-60e45ade98c3image.png "显示“从 Office 应用商店添加”、“从 URL 添加”和“从文件添加”选项的 + 按钮菜单。")

4. 在“从文件添加”对话框中，浏览至 d:\WhosWhoAD 中的 manifest.xml 所在位置，依次选择“打开”、“下一步”。
你随后可以在 Outlook 加载项列表中看到 Who's Who AD 加载项，如图 4 所示。
![图 4：在 Exchange 安装中心安装的 Who's Who AD 应用](/description/3977704e-6d30-4067-bb17-e5d1f778795cimage.png "Outlook 应用页面上列出的 Who's Who AD 应用，所示为已为用户启用。")

5. 如果 Outlook 正在运行，请关闭 Outlook，然后重新打开。

**注意**
仅当 Outlook 帐户位于 Exchange Server 2013 或更高版本上时，此程序才适用。
 
此外，在步骤 3 中，如果未看到“从文件添加”选项，则需要请求 Exchange 管理员为你提供必要的权限。

Exchange 管理员可以运行下列 PowerShell cmdlet，向一个用户分配必要权限。在本示例中，wendyri 是用户的电子邮件别名。

```POWERSHELL
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
 ```

如有必要，管理员可以运行下列 cmdlet，向多个用户分配类似的权限：

```POWERSHELL
$users = Get-Mailbox *
$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
 ```

有关我的自定义应用角色的详细信息，请参阅[我的自定义应用角色](http://msdn.microsoft.com/library/aa0321b3-2ec0-4694-875b-7a93d3d99089(Office.15).aspx)。

**部署 Web 服务**

请执行以下步骤以部署 Who Web 服务和 WhoMailApp.html 邮件加载项文件：

1. 在 Visual Studio 中，打开 WhoWebService.csproj。
2. 依次选择“构建”、“WhoWebService”。
3. 在“发布 Web”对话框的“配置文件”选项卡中，指定一个选择的配置文件。
4. 在“连接”选项卡中，选择“文件系统”作为发布方法。
5. 键入 \\webserver\c$\inetpub\wwwroot\WhosWhoAD 作为目标位置。
6. 选择“发布”。
7. 在 webserver 计算机上，启动 IIS 管理器。
8. 在“连接”窗格中，依次选择“站点”、“默认 Web 站点”。
9. 右键单击 WhosWhoAD 文件夹，然后选择“转换为应用程序”。
10. 在“添加应用程序”对话框中，在具有默认列出的 DefaultAppPool 的应用程序池下，选择“选择”。
11. 在“选择应用程序池”对话框中的“属性”下，确保显示的是 .Net Framework 版本：4.0 或更高版本的 .NET Framework。如有必要，请选择不同的应用程序池，以确保池使用的版本至少为 .NET Framework 4.0。选择“确定”。
12. 在“添加应用程序”对话框中，确保看到“通过身份验证”，如图 6 所示。选择“确定”。继续执行步骤 14。
![图 5：“添加应用程序”对话框，用于将 Who Web 服务转换为 IIS 上相应应用程序池中的应用程序](/description/584495fd-6e4d-45da-9b5b-47d114848342image.png "显示文本“通过身份验证”的“添加应用程序”对话框。")

13. 作为步骤 10 至 12 的替代步骤，也可以创建使用 .NET Framework 4.0（或更高版本）和通过身份验证的新应用程序。选择相应的应用程序池并继续执行步骤 14。
14. 在 IIS 管理器的中间窗格中，选择“身份验证”。验证 Windows 身份验证是否已启用；如有必要，则右键单击以启用。

按照部署程序将以下文件复制到 \\webserver\c$\inetpub\wwwroot\WhosWhoAD\:
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

现在即可在 webserver 上访问 Who Web 服务，并且还可以在 Outlook 或 Outlook Web App 中使用 Who's Who AD 邮件应用。

**运行并测试示例**

1. 在 Outlook 中，选择要在阅读窗格中读取的电子邮件。
2. 从加载项栏中选择 Who's Who AD 邮件加载项。

你应可以在加载项窗格中看到 Active Directory 数据，类似于图 1 中的示例。

<a name="Troubleshooting"></a>
**疑难解答**

如果电子邮件的发件人或收件人所拥有的电子邮件格式为 <first name>.<last name>@<domain>，则 Active Directory 包装可能无法在 Active Directory 中搜索相应的人员。选择电子邮件地址格式为 <alias>@<domain> 的人员。

Who's Who AD 邮件加载项旨在用作原型，因此，无法自定义该邮件加载项以适应组织的要求。有关更多信息，请参阅附带文章中的“未来扩展”部分。

**相关内容**

* [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [操作方法：创建邮件应用以通过 Active Directory 显示层次结构信息](https://blogs.msdn.microsoft.com/officeapps/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients/)


此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
