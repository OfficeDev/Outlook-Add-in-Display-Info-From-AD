# Outlook-Add-in-Display-Info-From-AD
Learn from this prototype mail app how to access basic hierarchy information from Active Directory. Extend this prototype to customize the mail app for your organization.

**Description of the Who's Who AD mail app sample**

This sample accompanies the topic  [How to: Create a mail app to display hierarchy information from Active Directory in the MSDN Library](http://msdn.microsoft.com/library/bb419185-f004-4118-a53d-3b6c8e984c9e.aspx).

When you select an email message in Outlook or Outlook Web App, you can choose the Who's Who AD mail add-in to display Active Directory information about the sender and other recipients of an email message currently selected in Outlookor Outlook Web App. The mail add-in appears in the app bar when you are viewing an email in the Reading Pane or in the mail explorer.

When you first choose this mail add-in, it retrieves and displays the sender's detailed professional and hierarchical information from Active Directory—name, job title, department, alias, office number, telephone number, and a picture thumbnail. If the sender has a manager or direct reports, the mail add-in displays a similar subset of information for each of them, as well. Figure 1 shows an example of the Who's Who AD app. The screen shot displays information for Belinda Newman, her manager, and direct reports.


![Figure 1. Mail app displays Active Directory information for an email sender in Outlook](/description/image.png)
 
The mail app provides a navigation bar that allows you to choose a recipient and view detailed professional and hierarchy information that is stored in Active Directory.

Behind the scenes, when you select a sender or recipient, the mail app calls a web service, named Who, to get the person's data from Active Directory. The web service includes an Active Directory wrapper, which uses Active Directory Domain Services (AD DS) to access information from Active Directory. After getting the data, the Who web service serializes the data in JSON format and sends it back as the web service response. The mail app then pulls the data and displays it on the app pane. Figure 2 summarizes the relationships among the Outlook user, mail app, Who web service, and Active Directory.

Figure 2. Relationships among the Outlook user, mail app, Who web service, and Active Directory

 
See the accompanying article  How to: Create a mail app to display hierarchy information from Active Directory in the MSDN Library for a description of the implementation of the mail app and the Who web service.



Note


The Who web service serves only as a prototype and shows a few basic features of Active Directory that are familiar to most Active Directory users. Hopefully this example provides a good starting point for you to extend and support features that are specific to your organization. For more information, see the section  Future extension in the accompanying article.
 

