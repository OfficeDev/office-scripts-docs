---
title: 'Overview: Office Scripts in Excel on the web'
description: 'The prerequisites and environment information for Office Scripts for Excel on the web.'
ms.date: 11/07/2019
localization_priority: Normal
---

# Overview: Office Scripts in Excel on the web

> [!NOTE]
> Office Scripts in Excel on the web is in preview. The functionality described here is subject to change as the feature develops. To learn how to use this feature, see this page. You can submit feedback on Office Scripts through the Excel feedback button. You can submit feedback on the documentation here.

Office Scripts in Excel on the web lets you automate common tasks. You can record your Excel actions with the [Script Recorder](https://aka.ms/makersdogfood), which creates a script. This script can be attached to any workbook you own and shared with colleagues. This script can also be edited by you with the [Script Editor](https://aka.ms/makersdogfood). This series of documents will teach you the fundamentals of Office Script coding.

> [!NOTE]
> If you are looking to develop an Office Add-In for Excel, please visit our [quick start](/office/dev/add-ins/quickstarts/excel-quickstart-jquery) or [tutorial](/office/dev/add-ins/tutorials/excel-tutorial). You can learn more about the differences between Office Scripts and Office Add-ins in [the Differences between Office Scripts and Office Add-ins article](../resources/differences-scripts-add-ins.md).

## Script Environment

There are two frameworks to be aware of for an Office Script:

- TypeScript/JavaScript – The language in which the code is written.
- Office JavaScript API – The library that allows your script to interact with an Excel workbook.

### TypeScript/JavaScript

Office Scripts can be written in TypeScript or JavaScript. All the code created by the Script Recorder is TypeScript (which is a superset of JavaScript). Office Scripts documentation uses TypeScript, but if you are more comfortable with JavaScript, you may use that instead.

If you are new to programming, we recommend two steps before proceeding with Office Scripts:

- Learn the basics of JavaScript. You should feel comfortable with concepts like variables, control flow, functions, and data types. [Mozilla offers a good, comprehensive tutorial on JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Learn about types in TypeScript. TypeScript builds on JavaScript by ensuring at compile-time that the right types are used for method calls and assignments. The TypeScript documentation on [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), [classes](https://www.typescriptlang.org/docs/handbook/classes.html), [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html), and [type compatibility](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) will be the most useful.

Office Scripts are largely self-contained pieces of code. Only a small part of TypeScript’s functionality is used. Therefore, you can edit scripts without having to learn the intricacies of the TypeScript. The Script Editor also handles the installation, compilation, and execution of code, so you don’t need to worry about anything but the script itself.

### Office JavaScript API

Office Scripts uses the Office JavaScript API, specifically the [ExcelApi Online requirement set](/javascript/api/excel). That requirement set has every released Office JavaScript API and they are all accessible to Excel on the web. Your scripts have access to most of the features in the Excel namespace. The exceptions are listed in the [Differences between Office Scripts and Office Add-ins](../resources/differences-scripts-add-ins.md#apis) article.

## IntelliSense

IntelliSense is a Script Editor feature that makes code suggestions. As you type, it displays possible object and field names. It also displays inline documentation for every API.

The Excel Script Editor uses the same IntelliSense engine as Visual Studio Code. To learn more about the feature as a whole, visit [Visual Studio Code’s IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## Script Recorder

All scripts recorded with the Script Recorder can be edited through the Script Editor. One easy way to learn the capabilities of the Office Scripts codebase is to record scripts and view the resulting code. Another option is to follow our [tutorial](../tutorials/office-scripts.md) to learn is a more guided and structured way. For more on the Script Recorder, visit [Office Scripts in Excel for the web](https://aka.ms/makersdogfood).

## Next steps

Visit [Tutorial: Office Scripts in Excel on the web](../tutorials/office-scripts.md). There, you will learn how to write your first Office Scripts.

## See also

- [Scripting Fundamentals for Office Script in Excel on the web](../develop/scripting-fundamentals.md)
- [Differences between Office Scripts and Office Add-ins](../resources/differences-scripts-add-ins.md)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Office Scripts API Reference](/javascript/api/excel)
- [Office Add-ins Documentation](/office/dev/add-ins)
