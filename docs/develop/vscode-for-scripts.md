---
title: Visual Studio Code for Office Scripts (preview)
description: How to setup the Office Scripts Code Editor to connect with VS Code on the web.
ms.date: 11/04/2022
ms.localizationpriority: medium
---

# Visual Studio Code for Office Scripts (preview)

[Visual Studio Code on the web](https://vscode.dev/) lets users edit anything from anywhere. Connect your Office Scripts experience to this popular code editor to start scripting outside of the workbook.

:::image type="content" source="../images/vscode-script-editor.png" alt-text="An Excel on the web window with the Code Editor open next to a VS Code on the web window with an open script.":::

Visual Studio Code has a few advantages over the built-in Code Editor.

- Full-screen editing! Your script doesn’t have to share screen space with the workbook any more.
- Edit multiple scripts at once! Quickly switch between scripts to share code from your other automations.
- Extensions! Use your favorite VS Code extensions for spell checking, formatting, and whatever else helps you get the job done.

> [!NOTE]
> This feature is in preview. It's subject to change based on feedback. If you encounter any issues, please report them through the **Feedback** button in Excel. The following are known issues with the current version of the feature.
>
> - Visual Studio Code can only be connected to Office Scripts through Excel on the web.
> - This Office Scripts connection is only available with English-language Excel clients.

## Connect Visual Studio Code to Office Scripts

Follow these one-time steps to connect Visual Studio Code and Excel on the web.

1. Open the Office Scripts **Code Editor**.
2. Under the **More options (…)** menu, select **Editor settings**.
3. Select **Visual Studio Code connection**.

:::image type="content" source="../images/vscode-enable-option.png" alt-text="The editor settings task pane showing a checkbox labeled Visual Studio Code connection.":::

Now you can edit and run your scripts from Visual Studio Code. From any script, go to the **More options (…)** menu and select **Open in VS Code**.

:::image type="content" source="../images/vscode-open-option.png" alt-text="The Open in VS Code option being selected from a list next to an open script.":::

## See also

- [Office Scripts Code Editor environment](../overview/code-editor-environment.md)
- [Visual Studio Code for the Web (documentation)](https://code.visualstudio.com/docs/editor/vscode-web)
