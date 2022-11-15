---
title: Office Scripts Code Editor environment
description: The prerequisites and environment information for Office Scripts in Excel.
ms.date: 11/08/2022
ms.localizationpriority: medium
---

# Office Scripts Code Editor environment

Office Scripts are written in either TypeScript or JavaScript and use the Office Scripts JavaScript APIs to interact with an Excel workbook. The Code Editor is based on Visual Studio Code, so if you've used that environment before, you'll feel right at home.

> [!TIP]
> If you're familiar with Visual Studio Code, you can now use it to write scripts. Visit [Visual Studio Code for Office Scripts (preview)](../develop/vscode-for-scripts.md) to try out this feature.

## Scripting language: TypeScript or JavaScript

Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). The Action Recorder generates code in TypeScript and the Office Scripts documentation uses TypeScript. Since TypeScript is a superset of JavaScript, any scripting code that you write in JavaScript will work just fine.

Office Scripts are largely self-contained pieces of code. Only a small part of TypeScript's functionality is used. Therefore, you can edit scripts without having to learn the intricacies of TypeScript. The Code Editor also handles the installation, compilation, and execution of code, so you don't need to worry about anything but the script itself. It's possible to learn the language and create scripts without previous programming knowledge. However, if you're new to programming, we recommend learning some fundamentals before proceeding with Office Scripts.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## Office Scripts JavaScript API

Office Scripts use a specialized version of the Office JavaScript APIs for [Office Add-ins](/office/dev/add-ins/overview/index). While there are similarities in the two APIs, you should not assume code can be ported between the two platforms. The differences between the two platforms are described in the [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis) article. You can view all the APIs available to your script in the [Office Scripts API reference documentation](/javascript/api/office-scripts/overview).

## External library support

Office Scripts does not support the usage of external, third-party JavaScript libraries. Currently, you cannot call any library other than the Office Scripts APIs from a script. You do still have access to any [built-in JavaScript object](../develop/javascript-objects.md), such as [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## IntelliSense

IntelliSense is a set of Code Editor features that help you write code. It provides auto-complete, syntax error highlighting, and inline API documentation.

IntelliSense gives suggestions as you type, similar to the suggested text in Excel. Pressing the Tab or Enter key inserts the suggested member. Trigger IntelliSense at the current cursor location by pressing the Ctrl+Space keys. These suggestions are especially useful when completing a method. The method signature displayed by IntelliSense contains a list of arguments it needs, each argument's type, whether a given argument is required or optional, and the return type of the method.

Hover the cursor over a method, class, or other code object to see more information. Hover over a syntax error or code suggestion, represented by a red or yellow squiggly line, to see suggestions on how to fix the problem. Often, IntelliSense provides a "Quick Fix" option to automatically change the code.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="An error message in the Code Editor's hover text with a 'Quick Fix' button.":::

The Office Scripts Code Editor uses the same IntelliSense engine as Visual Studio Code. To learn more about the feature, visit [Visual Studio Code's IntelliSense Features](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## Keyboard shortcuts

Most of the keyboard shortcuts for Visual Studio Code also work in the Office Scripts Code Editor. Use the following PDFs to learn about the available options and get the most out of the Code Editor:

- [Keyboard shortcuts for macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Keyboard shortcuts for Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## See also

- [Office Scripts API reference](/javascript/api/office-scripts/overview)
- [Troubleshooting Office Scripts](../testing/troubleshooting.md)
- [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md)
- [Visual Studio Code for Office Scripts (preview)](../develop/vscode-for-scripts.md)
