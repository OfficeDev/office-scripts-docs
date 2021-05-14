---
title: 'Office Scripts Code Editor environment'
description: 'The prerequisites and environment information for Office Scripts in Excel on the web.'
ms.date: 05/10/2021
localization_priority: Normal
---

# Office Scripts Code Editor environment

Office Scripts are written in either TypeScript or JavaScript and use the Office Scripts JavaScript APIs to interact with an Excel workbook. The Code Editor is based on Visual Studio Code, so if you've used that environment before, you'll feel right at home.

## Scripting language: TypeScript or JavaScript

Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). The Action Recorder generates code in TypeScript and the Office Scripts documentation uses TypeScript. Since TypeScript is a superset of JavaScript, any scripting code that you write in JavaScript will work just fine.

Office Scripts are largely self-contained pieces of code. Only a small part of TypeScript's functionality is used. Therefore, you can edit scripts without having to learn the intricacies of TypeScript. The Code Editor also handles the installation, compilation, and execution of code, so you don't need to worry about anything but the script itself. It's possible to learn the language and create scripts without previous programming knowledge. However, if you're new to programming, we recommend learning some fundamentals before proceeding with Office Scripts:

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## Office Scripts JavaScript API

Office Scripts use a specialized version the Office JavaScript APIs for [Office Add-ins](/office/dev/add-ins/overview/index). While there are similarities in the two APIs, you should not assume code can be ported between the two platforms. The differences between the two platforms are described in the [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis) article. You can view all the APIs available to your script in the [Office Scripts API reference documentation](/javascript/api/office-scripts/overview).

## External library support

Office Scripts do not support the usage of external, third-party JavaScript libraries. Currently, you cannot call any library other than the Office Scripts APIs from a script. You do still have access to any [built-in JavaScript object](../develop/javascript-objects.md), such as [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## IntelliSense

IntelliSense is a Code Editor feature that helps prevent typos and syntax errors as you edit your script. It displays possible object and field names as you type, as well as inline documentation for every API.

The Excel Code Editor uses the same IntelliSense engine as Visual Studio Code. To learn more about the feature, visit [Visual Studio Code's IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## Keyboard shortcuts

Most of the keyboard shortcuts for Visual Studio Code also work in the Office Scripts Code Editor. Use the following PDFs to learn about the available options and get the most out of the Code Editor:

- [Keyboard shortcuts for macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Keyboard shortcuts for Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## See also

- [Office Scripts API reference](/javascript/api/office-scripts/overview)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md)
