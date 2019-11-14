---
title: 'Office Scripts in Excel on the web'
description: 'The prerequisites and environment information for Office Scripts for Excel on the web.'
ms.date: 11/14/2019
localization_priority: Normal
---

# Office Scripts in Excel on the web

[!INCLUDE [Preview note](../includes/preview-note.md)]

Office Scripts in Excel on the web let you automate common tasks. You can record your Excel actions with the [Action Recorder](https://aka.ms/makersdogfood), which creates a script. You can also create and edit scripts with the Code Editor. Scripts can be attached to any workbook you own and shared with colleagues. This series of documents teaches you to use the Code Editor to make your own scripts by learning the fundamentals of Office Scripts.

> [!NOTE]
> If you're looking to develop an Office Add-In for Excel, please visit our [quick start](/office/dev/add-ins/quickstarts/excel-quickstart-jquery) or [tutorial](/office/dev/add-ins/tutorials/excel-tutorial). Learn more about the differences between Office Scripts and Office Add-ins in [the Differences between Office Scripts and Office Add-ins article](../resources/differences-scripts-add-ins.md).

## Script environment

Office Scripts are written in either [TypeScript or JavaScript](#scripting-language-typescript-or-javascript) and use the [Office JavaScript APIs](#office-javascript-api) to interact with an Excel workbook.

### Scripting language: TypeScript or JavaScript

Office Scripts are written in TypeScript or JavaScript. The Action Recorder generates code in TypeScript (which is a superset of JavaScript). Office Scripts documentation uses TypeScript, but if you're more comfortable with JavaScript, use that instead.

If you're new to programming, we recommend two steps before proceeding with Office Scripts:

- Learn the basics of JavaScript. You should feel comfortable with concepts like variables, control flow, functions, and data types. [Mozilla offers a good, comprehensive tutorial on JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Learn about types in TypeScript. TypeScript builds on JavaScript by ensuring at compile-time that the right types are used for method calls and assignments. The TypeScript documentation on [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), [classes](https://www.typescriptlang.org/docs/handbook/classes.html), [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html), and [type compatibility](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) will be the most useful.

Office Scripts are largely self-contained pieces of code. Only a small part of TypeScript's functionality is used. Therefore, you can edit scripts without having to learn the intricacies of the TypeScript. The Code Editor also handles the installation, compilation, and execution of code, so you don't need to worry about anything but the script itself.

### Office JavaScript API

Office Scripts uses the Office JavaScript API. Specifically, they use the APIs available in the [ExcelApi Online requirement set](/javascript/api/excel). Your scripts have access to most of the features in the Excel namespace. The exceptions are listed in the [Differences between Office Scripts and Office Add-ins](../resources/differences-scripts-add-ins.md#apis) article.

## IntelliSense

IntelliSense is a Code Editor feature that helps prevent typos and syntax errors as you edit your script. It displays possible object and field names as you type, as well as inline documentation for every API.

The Excel Code Editor uses the same IntelliSense engine as Visual Studio Code. To learn more about the feature, visit [Visual Studio Code's IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## Action Recorder

All scripts recorded with the Action Recorder can be edited through the Code Editor. One easy way to learn the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code. Another option is to follow our [tutorial](../tutorials/excel-office-scripts-tutorial.md) to learn in a more guided and structured way. For more information about the Action Recorder, visit [Office Scripts in Excel for the web](https://aka.ms/makersdogfood).

## Next steps

Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-office-scripts-tutorial.md) to learn how to create your first Office Scripts.

## See also

- [Scripting fundamentals for Office Script in Excel on the web](../develop/scripting-fundamentals.md)
- [Differences between Office Scripts and Office Add-ins](../resources/differences-scripts-add-ins.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts API reference](/javascript/api/excel)
- [Office Add-ins documentation](/office/dev/add-ins)
