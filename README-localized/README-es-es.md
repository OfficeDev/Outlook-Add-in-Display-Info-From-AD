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
Con este prototipo del complemento para correo electrónico podrá aprender a acceder a la información básica de jerarquías de Active Directory (AD). Amplíe el prototipo para personalizar el complemento para correo para su organización.

**Descripción de la muestra del complemento para correo de Who's Who AD**

En esta muestra se describe el tema de los [procedimientos: Cree una aplicación de correo electrónico para mostrar la información de las jerarquías de Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) en el blog de aplicaciones para Office y SharePoint.

Al seleccionar un mensaje de correo electrónico en Outlook o en la aplicación web de Outlook, puede elegir el complemento para correo de AD Who's Who para mostrar la información de Active Directory sobre el remitente y otros destinatarios de un mensaje de correo electrónico actualmente seleccionados en Outlook o en Outlook Web App. El complemento para correo se mostrará en la barra de la aplicación al ver un correo electrónico en el panel de lectura o en el explorador de correo.

Al elegir el complemento para el correo por primera vez, se recuperará y se mostrará la información jerárquica y profesional detallada del remitente de Active Directory: nombre, puesto, departamento, alias, número de oficina, número de teléfono y una miniatura de la imagen. Si el remitente tiene informes directos o un gerente en el complemento para correo, también se le mostrará un subconjunto de información similar para cada uno. La figura 1 muestra un ejemplo del complemento Who's Who de AD. La captura de pantalla muestra información e informes directos a Belinda Newman, la gerente de la empresa.


![Figura 1. La aplicación del correo muestra la información de Active Directory para un remitente de correo electrónico en Outlook](/description/image.png "la aplicación de correo Who's Who de AD mostrando iconos para cada usuario con su foto, nombre y puesto.")
 
El complemento para correo proporciona una barra de navegación que permite elegir un destinatario y ver la información detallada de jerarquía y profesional que almacenada en Active Directory.

Al seleccionar un remitente o un destinatario en segundo plano, el complemento de correo llama a un servicio web, denominado Who, para obtener los datos del usuario de Active Directory. El servicio web incluye un contenedor de Active Directory, que usa servicios de su dominio (AD DS) para obtener acceso a la información de Active Directory. Después de obtener los datos, el servicio web de Who los serializa en formato JSON y los envía de vuelta como respuesta al servicio web. El complemento para el correo extrae los datos y luego los muestra en el panel de complementos. La figura 2 resume las relaciones entre el usuario de Outlook, la aplicación de correo, el servicio web de Who y Active Directory.

![Figura 2. Las relaciones entre el usuario de Outlook, la aplicación de correo, el servicio web de Who y Active Directory](/description/75964e99-a74f-4fd3-a4e1-a0ca40bfc5a4image.png "el usuario interactúa con la aplicación de correo y AD; el correo interactúa con el usuario y el servicio web; el servicio web interactúa con la aplicación de correo y AD.")

 
Vea el artículo adjunto sobre el [procedimiento: Cree una aplicación para correo que muestre la información de la jerarquía de Active Directory](http://blogs.msdn.com/b/officeapps/archive/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients.aspx) en el blog de las aplicaciones para Office y SharePoint para obtener una descripción de la implementación del complemento para correo y del servicio web Who.

**Nota**

El servicio web Who solo funciona como prototipo y muestra algunas de las características básicas de Active Directory que ya son familiares para la mayoría de los usuarios. Afortunadamente, este ejemplo es un buen punto de partida para ampliar y ofrecer soporte técnico a las características específicas de su organización. 
 
**Requisitos previos**

Para sacar el máximo provecho de esta muestra de código, debe estar familiarizado con el desarrollo web mediante HTML y JavaScript, y con los servicios web de Windows Communication Foundation (WCF). No es necesario que tenga conocimiento previo sobre los servicios de dominio de Active Directory.

Los siguientes requisitos son necesarios para instalar y ejecutar cualquier complemento para el correo, incluyendo el complemento para correo Who's Who de AD:

* El buzón del usuario debe estar en Exchange Server 2013 o en una versión posterior.
* La aplicación del correo debe ejecutarse en Outlook 2013, en una versión posterior o en la Outlook Web App.

Puede usar cualquier herramienta de desarrollo web con la que esté familiarizado para desarrollar la aplicación para correo Who's Who de AD.

Las siguientes herramientas fueron usadas para desarrollar el servicio web Who y para implementar la aplicación de correo y el servicio web:

* Visual Studio 2012
* .NET Framework 4.0
* Windows Server 2008
* Servidor de información de Internet (IIS, por sus siglas en inglés) 7.0

**Componentes clave de la muestra**

La descarga de esta muestra consta de los siguientes archivos y carpetas:

* Manifest.xml es el archivo de manifiesto para el complemento de correo Who's Who de AD.
* WhoMailApp.sln es el archivo de solución de Visual Studio para la muestra.
* La carpeta WhoAgave contiene archivos del complemento para el correo (incluyendo HTML, imágenes y archivos CSS) y algunos archivos para el servicio web Who.
* La carpeta de ActiveDirectory contiene los archivos del contenedor de Active Directory.
* La carpeta de BuildProcessTemplates contiene los archivos de plantilla de marcado predeterminada para desarrollar los servicios web de WCF.

**Configuración de la muestra**

Siga los siguientes pasos para obtener los archivos y modificar sus referencias, según corresponda:

1. En una unidad local, por ejemplo d:, cree una carpeta llamada WhosWhoAD y descargue aquí los archivos de muestra.
2. Asumiendo que el servidor web de IIS en el que intenta hospedar la aplicación para correo Who's Who de AD se denomina servidor web, cree una carpeta llamada WhosWhoAD en \\webserverc$\\inetpubwwwwroot.
3. Copia el contenido de la carpeta img de d:\\WhosWhoAD\\WhoAgave\\ y péguelo en \\webserver\\c $ \\inetpub\\wwwroot\\WhosWhoAD.
El resto de los archivos del complemento de correo y del servicio web se copiarán correctamente al servidor web cuando se despliegue el servicio web, como se describe en la sección de despliegue del servicio web que se muestra a continuación.
4. Actualice el archivo de manifiesto para reflejar la ubicación actual del archivo HTML del complemento para correo.

El archivo de manifiesto para complementos de correo, manifest.xml, se encuentra directamente después de d:\\WhosWhoAD. Si su actual servidor web tiene un nombre diferente al de servidor web, actualice el archivo manifest.xml para reflejar la ubicación real del archivo WhoMailApp.html, reemplazando el servidor web en la siguiente línea con la ruta del servidor de la carpeta WhosWhoAD que usted creó durante el paso 2.

```XML
<SourceLocation DefaultValue="https://webserver/WhosWhoAD/WhoMailApp.html"/>
 ```

**Instalación del complemento para el correo**

1. En el cliente mejorado de Outlook, elija el archivo, administrar aplicaciones. Esto abrirá el explorador para iniciar sesión en la Outlook Web App e ir al centro de administración de Exchange (EAC, por sus siglas en inglés).
2. Inicie sesión en su cuenta Exchange.
3. En el EAC, seleccione el cuadro desplegable de lista junto al botón + y, después, elija agregar desde archivo, como se muestra en la
![figura 3. Instale una aplicación para el correo desde un archivo en el centro de administración de Exchange](/description/f7a57314-42f1-4d15-9752-60e45ade98c3image.png "El botón Más del menú muestra las opciones para agregar desde la tienda Office, agregar desde la URL y agregar desde un archivo.")

4. En el cuadro de diálogo para agregar desde el archivo, busque la ubicación del archivo manifest.xml en d:\\WhosWhoAD, elija abrir, y luego elija siguiente.
Debería poder ver el complemento Who's Who de AD en la lista de complementos de Outlook, como se muestra en la figura 4.
![Figura 4. La aplicación Who's Who de AD instalada en el centro de administración de Exchange](/description/3977704e-6d30-4067-bb17-e5d1f778795cimage.png ". La aplicación Who's Who de AD en la lista de la página de aplicaciones para Outlook, se muestra como habilitada para el usuario.")

5. Cuando Outlook esté en funcionamiento, ciérrelo y vuelva a abrirlo.

**Nota**
este procedimiento solo puede ser aplicado si la cuenta de Outlook está en Exchange Server 2013 o en una versión posterior.
 
Asimismo, si en el paso 3, no ve la opción para agregar desde el archivo, debe solicitar que el administrador de Exchange le proporcione los permisos necesarios.

El administrador de Exchange puede ejecutar el siguiente cmdlet de PowerShell para otorgar a un solo usuario los permisos necesarios. En este ejemplo, wendyri es el alias de correo electrónico del usuario.

```POWERSHELL
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
 ```

Si es necesario, el administrador puede ejecutar el siguiente cmdlet siguiente para otorgar permisos similares a varios usuarios:

```POWERSHELL
$users = Get-Mailbox *
$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
 ```

Para más información sobre el rol de mis aplicaciones personalizadas, vea [rol de mis aplicaciones personalizadas](http://msdn.microsoft.com/library/aa0321b3-2ec0-4694-875b-7a93d3d99089(Office.15).aspx).

**Implementar el servicio web**

Haga lo siguiente para implementar el servicio Web Who y el archivo del complemento para correo WhoMailApp.html:

1. Abra WhoWebService.csproj en Visual Studio.
2. Elija generar, luego publicar WhoWebService.
3. En la pestaña del Perfil en el cuadro de diálogo sobre la publicación en la web, especifique el perfil que prefiera.
4. Seleccione sistema de archivos como método de publicación en la pestaña de conexión
5. Escriba \\webserver\c$\inetpub\wwwroot\WhosWhoAD como la ubicación de destino.
6. Elija publicar.
7. Inicie el administrador de IIS en el ordenador del servidor web.
8. Vaya al panel conexiones, elija sitios, y seleccione el sitio web predeterminado.
9. Haga clic con el botón derecho en la carpeta WhosWhoAD y elija la opción de convertir en aplicación.
10. En el cuadro del diálogo para agregar la aplicación, con DefaultAppPool incluido en la lista de forma predeterminada, elija seleccionar.
11. En el cuadro de diálogo, vaya a seleccionar grupos de aplicaciones, en las propiedades, y asegúrese de tener la siguiente versión de .Net Framework: Debe mostrarse una versión 4.0 o posterior de .NET Framework. Elija un grupo de aplicaciones distinto, en caso necesario, para asegurarse de que en el grupo se utilice al menos .NET Framework 4.0. Elija aceptar.
12. En el cuadro de diálogo para agregar la aplicación, asegúrese de ver la autenticación paso a paso, como se muestra en la figura 6. Elija aceptar. Vaya al paso 14.
![Figura 5. Añada el cuadro de diálogo para añadir la aplicación que convierte el servicio web de Who en una aplicación del grupo apropiadas en el IIS](/description/584495fd-6e4d-45da-9b5b-47d114848342image.png "para añadir el cuadro de diálogo a una aplicación se muestra el texto de autentificación paso a paso.")

13. Como una alternativa a los pasos del 10 al 12, puede crear un nuevo grupo de aplicaciones que utilice .NET Framework 4.0 (o una versión posterior) y la autenticación paso a paso. Seleccione el grupo de aplicaciones y vaya al paso 14.
14. En el panel central del administrador del IIS, elija la autenticación. Compruebe que la autenticación de Windows está habilitada; de ser necesario, haga clic con el botón derecho para habilitarlo.

El procedimiento de implementación copia los siguientes archivos en \\webserver\c$\inetpub\wwwroot\WhosWhoAD\:
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

Ahora puede obtener acceso al servicio web Who mediante un servidor web y después usar la aplicación de correo WHO de AD en Outlook o Outlook Web App.

**Inicie y pruebe la aplicación**

1. Seleccione un correo electrónico en Outlook para leerlo en el panel de lectura.
2. Elija el complemento para correo de Who's Who de AD en la barra de complementos.

Debería poder ver los datos de Active Directory en el panel de complementos, de manera similar al ejemplo de la figura 1.

<a name="Troubleshooting"></a>
**Solución de problemas**

En caso de que el remitente o el destinatario de un mensaje de correo electrónico tenga una dirección de correo electrónico del formulario <first name>. <last name>@<domain>, el contenedor de Active Directory puede no ser capaz de buscar a la persona indicada en esta sección. Elija una persona cuya dirección de correo electrónico tenga el símbolo <alias>@<domain>.

Dado que el complemento para correo Who's Who de AD está pensado para funcionar como un prototipo, aún tiene espacio para que pueda personalizar el complemento de correo para que cumpla con los requisitos de su organización. Para más información, vea la sección sobre la extensión futura en el artículo adjunto.

**Contenido relacionado**

* [Más muestras de complementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Procedimiento: Cree una aplicación de correo electrónico para mostrar la información de las jerarquías de Active Directory](https://blogs.msdn.microsoft.com/officeapps/2013/05/15/creating-a-mail-app-to-check-out-active-directory-org-information-for-mail-senders-and-recipients/)


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
