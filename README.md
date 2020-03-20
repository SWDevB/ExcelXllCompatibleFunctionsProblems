# Project to reproduce XXL compatible functions issues
Enabling of XLL compatible functions prevents all functions from working as soon as one not supported function is added.

## How to reproduce
1. Run the project as it is. It should work, so ADD, CLOCK and INCREMENT functions are available and working
1. Include the commented part (EquivalentAddins) in the manifest.xml!
1. Run the project, loading the custom functions should fail. None of the above mentioned fuctions is available.
1. Remove the increment and clock function from the functions.ts file
1. Run the project, it should work again. ADD function is available again.

## Conclusion

So, it seems some of the functionallity used by the increment and clock functions break the compatibility mode.
Invocation isn't working at all for XLL compatible functions, or only certain ways to use it?

Another approach could be to explicitly enable XLL compatible mode only for certain functions.

## Excel version factor
The following versions have this issue
- Office 365 ProPlus Version 2002 (Build 12527.20278 Click-to-Run) Monthly Channel
- Office Online
This version it seems doesn't
- Office 365 ProPlus Version 2004 (Build 12711.20000 Click-to-Run) Office Insider

# Default Readme below
# Custom functions in Excel

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.  

This repository contains the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new custom functions project. You can also use this repository as a sample to base your own custom functions project from if you choose not to use the generator. For more detailed information about custom functions in Excel, see the [Custom functions overview](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-overview) article in the Office Add-ins documentation or see the [additional resources](#additional-resources) section of this repository.

## Debugging custom functions

This template supports debugging custom functions from [Visual Studio Code](https://code.visualstudio.com/). For more information see [Custom functions debugging](https://aka.ms/custom-functions-debug). For general information on debugging task panes and other Office Add-in parts, see [Test and debug Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).  If your question is about the Office JavaScript APIs, make sure it's tagged with  [office-js].

## Additional resources

* [Custom functions overview](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-overview)
* [Custom functions best practices](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-best-practices)
* [Custom functions runtime](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime)
* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2017 Microsoft Corporation. All rights reserved.
