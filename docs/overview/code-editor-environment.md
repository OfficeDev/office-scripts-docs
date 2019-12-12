---
title: 'Office Scripts Code Editor environment'
description: 'The prerequisites and environment information for Office Scripts for Excel on the web.'
ms.date: 12/12/2019
localization_priority: Normal
---

# Office Scripts Code Editor environment

Office Scripts are written in either [TypeScript or JavaScript](#scripting-language-typescript-or-javascript) and use the [Office JavaScript APIs](#office-javascript-api) to interact with an Excel workbook.

## Scripting language: TypeScript or JavaScript

Office Scripts are written in TypeScript or JavaScript. The Action Recorder generates code in TypeScript (which is a superset of JavaScript). Office Scripts documentation uses TypeScript, but if you're more comfortable with JavaScript, you can use that instead.

Office Scripts are largely self-contained pieces of code. Only a small part of TypeScript's functionality is used. Therefore, you can edit scripts without having to learn the intricacies of TypeScript. The Code Editor also handles the installation, compilation, and execution of code, so you don't need to worry about anything but the script itself. It's possible to learn the language and create scripts without previous programming knowledge. However, if you're new to programming, we recommend learning some fundamentals before proceeding with Office Scripts:

- Learn the basics of JavaScript. You should feel comfortable with concepts like variables, control flow, functions, and data types. [Mozilla offers a good, comprehensive tutorial on JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Learn about types in TypeScript. TypeScript builds on JavaScript by ensuring at compile-time the right types are used for method calls and assignments. The TypeScript documentation on [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), [classes](https://www.typescriptlang.org/docs/handbook/classes.html), [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html), and [type compatibility](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) will be the most useful.

## Office JavaScript API

Office Scripts uses the Office JavaScript API. Specifically, they use the APIs available in the [ExcelApi Online requirement set](/javascript/api/excel). Your scripts have access to most of the features in the Excel namespace. The exceptions are listed in the [Differences between Office Scripts and Office Add-ins](../resources/differences-scripts-add-ins.md#apis) article.

## IntelliSense

IntelliSense is a Code Editor feature that helps prevent typos and syntax errors as you edit your script. It displays possible object and field names as you type, as well as inline documentation for every API.

The Excel Code Editor uses the same IntelliSense engine as Visual Studio Code. To learn more about the feature, visit [Visual Studio Code's IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## See also

- [Office Scripts API reference](/javascript/api/office-scripts/overview)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
