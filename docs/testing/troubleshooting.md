---
title: Troubleshoot Office Scripts
description: Debugging tips and techniques for Office Scripts, as well as help resources.
ms.date: 11/11/2021
ms.localizationpriority: medium
---

# Troubleshoot Office Scripts

As you develop Office Scripts, you may make mistakes. It's okay. You have the tools to help find the problems and get your scripts working perfectly.

> [!NOTE]
> For troubleshooting advice specific to Office Scripts with Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).

## Types of errors

Office Scripts errors fall into one of two categories:

* Compile-time errors or warnings
* Runtime errors

### Compile-time errors

Compile-time errors and warnings are initially shown in the Code Editor. These are shown by the wavy red underlines in the editor. They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane. Selecting the error will give more details about the problem and suggest solutions. Compile-time errors should be addressed before running the script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="A compiler error shown in the Code Editor's hover text.":::

You may also see orange warning underlines and grey informational messages. These indicate performance suggestions or other possibilities where the script may have unintentional effects. Such warnings should be examined closely before dismissing them.

### Runtime errors

Runtime errors happen because of logic issues in the script. This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook. The following script generates an error when a worksheet named "TestSheet" is not present.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### Console messages

Both compile-time and runtime errors display error messages in the console when a script runs. They give a line number where the problem was encountered. Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.

The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error. Note the text `[5, 16]` at the beginning of the error string. This indicates the error is on line 5, starting at character 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="The Code Editor console displaying an explicit `any` error message.":::

The follow image shows the console output for a runtime error. Here, the script tries to add a worksheet with a the name of an existing worksheet. Again, note the "Line 2" preceding the error to show which line to investigate.
:::image type="content" source="../images/runtime-error-console.png" alt-text="The Code Editor console displaying an error from the `addWorksheet` call.":::

## Console logs

Print messages to the screen with the `console.log` statement. These logs can show you the current value of variables or which code paths are being triggered. To do this, call `console.log` with any object as a parameter. Usually, a `string` is the easiest type to read in the console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane. Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.

Logs do not affect the workbook.

## Automate tab not appearing or Office Scripts unavailable

The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.

1. [Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).
1. [Check that your browser is supported](platform-limits.md#browser-support).
1. [Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).
1. [Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. Ensure you're not logged in as an external user on an guest tenant

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## Help resources

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems. Often, you'll be able to find the solution to your problem through a quick Stack Overflow search. If not, ask your question and tag it with the "office-scripts" tag. Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.

## See also

- [Best practices in Office Scripts](../develop/best-practices.md)
- [Platform limits with Office Scripts](platform-limits.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
- [Troubleshoot Office Scripts running in PowerAutomate](power-automate-troubleshooting.md)
- [Undo the effects of Office Scripts](undo.md)
