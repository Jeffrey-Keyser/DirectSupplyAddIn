

<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->
[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]



<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/Jeffrey-Keyser/DirectSupplyAddIn">
    <img src="Outlook-Add-in-Microsoft-Graph-ASPNETWeb/Content/directsupplylogo80x80.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Outlook Web App (OWA) Add-in</h3>

  <p align="center">
    Setup Document
    <br />
    <a href="https://github.com/othneildrew/Best-README-Template"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/Jeffrey-Keyser/DirectSupplyAddIn">View Demo</a>
    ·
    <a href="https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/issues">Report Bug</a>
    ·
    <a href="https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/issues">Request Feature</a>
  </p>
</p>



<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Prerequisites](#prerequisites)
* [Getting Started](#getting-started)
  * [Installation](#installation)
* [Usage](#usage)
* [Roadmap](#roadmap)
* [Contributing](#contributing)
* [License](#license)
* [Contact](#contact)
* [Acknowledgements](#acknowledgements)



<!-- ABOUT THE PROJECT -->
## About The Project

[![Product Name Screen Shot][product-screenshot]](Outlook-Add-in-Microsoft-Graph-ASPNETWeb/Content/directsupplylogo80x80.png)

This is a Outlook Web App (OWA) Add-in designed during my internship at Direct Supply! The ultimate goal of this project is to create an add-in that can maximize a user's mailbox space through the deletion of attachments and unnecessary email threads.

Currently the add-in can:
* Login users to Outlook accounts and leverage Microsoft Graph / Google APIs.
* Delete duplicate emails in conversations threads.
* Save attachments to OneDrive / Google Drive, delete from email, and embed a hyperlink to attachment location in email.
  * Can do the above save, delete, and embed on an entire mail folder (Inbox).
  * All with the single click of a button! :smile:


A list of commonly used resources that I find helpful are listed in the acknowledgements.

### Built With
This section lists the major frameworks that my project uses.
* [JQuery](https://jquery.com)
* [Graph API](https://docs.microsoft.com/en-us/graph/overview)
* [Google Drive API](https://developers.google.com/drive/api/v3/about-sdk)
* [Outlook Mail API](https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) Used for email editing
* [ASP.NET MVC](https://docs.microsoft.com/en-us/aspnet/core/mvc/overview?view=aspnetcore-3.1) Using C# of course
* [OAuth 2.0 - Authorization Code Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow) For getting AcessTokens
* [MSAL .NET Library](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet)

### Prerequisites

Before getting started there are a few requirements you must ensure are fulfilled before beginning.

You should have:
* Access to Office 365 tenant or Outlook.com account that can use [Azure Active Directory admin center](https://aad.portal.azure.com/#@jeffreykeyser.onmicrosoft.com/dashboard/private/d5187d35-1fdc-4a26-9e3d-8e2939e56018) for registering the app.
  * Failure to use a developer account will result in Graph API call not working and in a MailboxNotEnabledForRESTAPI error
  * See [Overview for REST APIs](https://docs.microsoft.com/en-us/outlook/rest/get-started) for more details.
* Patience :triumph:

<!-- GETTING STARTED -->
## Getting Started

Alright, time to begin

1. First you will need to register the application in the [Microsoft Application Registration Portal](https://apps.dev.microsoft.com/) in order to obtain an app ID for accessing the Microsoft Graph API.
    1. **Log in with the identity of an administrator of your Office 365 tenancy to ensure that you are working in an Azure Active Directory that is associated with that tenancy** . See [Register an application with the Microsoft Identity Platform](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)
    1. There click the **Go to app list** button and then the **Add an app** button. Enter a name for the application and click **Create application**.
    1. Locate the **Application Secrets** section, and click **Generate New Password**. A dialog box will appear with the generated password. Copy this value and save it somewhere secure. After leaving this page, you will no longer be able to access this password.
    1. Locate the **Platforms** section, and click **Add Platform**. Choose **Web**, then enter 'https://localhost:44301/AzureADAuth/Authorize' under **Redirect URIs.**
      >**Note:** The port number in the redirect URI (`44301`) may be different on your development machine. You can find the correct port number for your machine by selecting the **Outlook-Add-in-Microsoft-GraphASPNETWeb** project in **Solution Explorer**, then looking at the **SSL URL** setting under **Development Server** in the properties window. Verify that **SSL Enabled** is **True**.
    1. Locate the **Microsoft Graph Permissions** section in the app registration. Next to **Delegated Permissions**, click **Add**. Select **Files.ReadWrite.All**, **Mail.ReadWrite**, **User.Read**. Next to **Application Permissions** grant the following: **email**, **Mail.ReadWrite**, **offline_access**, **openid**, and **profile**.
  
2. Ensure the following settings are used:
  - SUPPORTED ACCOUNT TYPES: "Accounts in this organizational directory only"
  - IMPLICIT GRANT: Do not enable any Implicit Grant options
 
Here's an example of what the page should look like when your done.

![The completed app registration](readme-images/app-permissions.PNG)

Edit [Web.config](https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/blob/master/Outlook-Add-in-Microsoft-Graph-ASPNETWeb/Web.config) and replace 'YOUR APP ID HERE' with the application ID and 'YOUR APP PASSWORD HERE' with the application secret you generated at the beginning. Also replace 'YOUR TENANT ID HERE' with the directory tenant ID found on the app registration site.


 *Optional* **For Configuring Google Drive API **

TODO:


## Run the solution

1. Open the Visual Studio solution file. 
2. Right-click **Outlook-Add-in-Microsoft-Graph-ASPNET** solution in **Solution Explorer** (not the project nodes), and then choose **Set startup projects**. Select the **Multiple startup projects** radio button. Make sure the project that ends with "Web" is listed first.
3. On the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
4. In **Solution Explorer**, select the **Outlook-Add-in-Microsoft-Graph-ASPNET** project node (not the top solution node and not the project whose name ends in "Web").
5. Press F5. The first time you do this, you will be prompted to specify the email and password of the user that you will use for debugging the add-in. Use the credentials of an admin for your O365 tenancy. 

    >NOTE: The browser will open to the login page for Office on the web. (So, if this is the first time you have run the add-in, you will enter the username and password twice.) 

The remaining steps depend on whether you are running the add-in in desktop Outlook or Outlook on the web.

### Run the solution with Outlook on the web

1. Outlook for Web will open in a browser window. In Outlook, click **New** to create a new email message. 
2. Below the compose form is a tool bar with buttons for **Send**, **Discard**, and other utilities. Depending on which **Outlook on the web** experience you are using, the icon for the add-in is either near the far right end of this tool bar or it is on the drop down down menu that opens when you click the **...** button on this tool bar.

   ![Icon for Insert Files Add-in](Outlook-Add-in-Microsoft-Graph-ASPNETWeb/Content/directsupplylogo32x32.png)

3. Click the icon to open the task pane add-in.

## Run the project with desktop Outlook

TODO:



### Installation



<!-- USAGE EXAMPLES -->
## Usage

Use this space to show useful examples of how a project can be used. Additional screenshots, code examples and demos work well in this space. You may also link to more resources.

_For more examples, please refer to the [Documentation](https://example.com)_



<!-- ROADMAP -->
## Roadmap

See the [open issues](https://github.com/othneildrew/Best-README-Template/issues) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.


<!-- CONTACT -->
## Contact

Jeffrey Keyser - jeff.keyser@outlook.com

Project Link: [Direct Supply Add-in](https://github.com/Jeffrey-Keyser/DirectSupplyAddIn)


<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements
* [GitHub Emoji Cheat Sheet](https://www.webpagefx.com/tools/emoji-cheat-sheet)
* [Readme Template](https://github.com/othneildrew/Best-README-Template)



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/Jeffrey-Keyser/DirectSupplyAddIn.svg?style=flat-square
[contributors-url]: https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/Jeffrey-Keyser/DirectSupplyAddIn.svg?style=flat-square
[forks-url]: https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/network/members
[stars-shield]: https://img.shields.io/github/stars/Jeffrey-Keyser/DirectSupplyAddIn.svg?style=flat-square
[stars-url]: https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/stargazers
[issues-shield]: https://img.shields.io/github/issues/Jeffrey-Keyser/DirectSupplyAddIn.svg?style=flat-square
[issues-url]: https://github.com/Jeffrey-Keyser/DirectSupplyAddIn/issues
[license-shield]: https://img.shields.io/github/license/Jeffrey-Keyser/DirectSupplyAddIn.svg?style=flat-square
[license-url]: https://github.com/Jeffrey-Keyser/DirectSupplyAddIn
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/jeffrey-keyser-a58457157/
[product-screenshot]: images/screenshot.png
