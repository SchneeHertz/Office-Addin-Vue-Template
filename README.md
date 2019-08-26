# Office-Addin-Vue-Template for Excel

This repository contains the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new Office Add-in that appears in the task pane.

## Install

Install dependencies:  
```npm install```

Copy files in ```./node_modules/@microsoft/office-js/dist/``` to ```./assets/office-js/```

Run dev-server:  
```npm run dev-server```

Add addin's manifest.xml to your Excel App, refer to [Sideload Office Add-ins for testing](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)


## JavaScript

This template is written using JavaScript. It added some package: vue.js, element-ui, stylus, and configured for webpack.  
If you want to add more package, refer to [Build an Office Add-in using Vue.js](https://github.com/OfficeDev/TrainingContent/tree/master/OfficeAddin/05%20Using%20modern%20JavaScript/Demos/03%20Office%20Add-in%20using%20VueJS)


## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)  
- [Use vue-devtools](https://github.com/vuejs/vue-devtools/blob/master/shells/electron/README.md)

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).  If your question is about the Office JavaScript APIs, make sure it's tagged with  [office-js].

## Additional resources

* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)


## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.
