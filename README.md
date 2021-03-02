# Warder-Plugin Development

This repository contains the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new Office Add-in that appears in the task pane. You can also use this repository as a sample to base your own project from if you choose not to use the generator.

## JavaScript

This template is written using JavaScript. For the [TypeScript](http://www.typescriptlang.org/) version of this template, go to [Office-Addin-TaskPane](https://github.com/OfficeDev/Office-Addin-TaskPane).

## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
  - can be used on mac without installing office apps
  - [Publish task pane and content add-ins to a SharePoint app catalog](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
  - only can be used on Windows 10
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)
  - only can be used on Windows 10

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).  If your question is about the Office JavaScript APIs, make sure it's tagged with  [office-js].

## Additional resources

- [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Check localhost:port and kill this process

    netstat -ano | findstr :<port>
    taskkill /PID <pid> /F

## Run code on Mac OS

    npm run dev-server //start the local web server if developing on Mac
    npm start          //test on the desktop
    npm run start:web  //test on a browser

## Sidelod office add-ins for testing on web (pending)

- [Debug add-ins in Office on the web](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-in-office-online)
  
- Get a Microsoft 365 developer account

- Set up an app catalog on SharePoint Online, [Publish task pane and content add-ins to an app catalog on SharePoint](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog)
  
- [Office add-ins XML manifest](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests?tabs=tabid-1)

- [Configure the App Catalog site for a web application](https://docs.microsoft.com/en-us/sharepoint/administration/manage-the-app-catalog)
  - 怎么看上去还需要注册域名？这怎么搞
