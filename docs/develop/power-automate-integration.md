---
title: 'Integrate Office Scripts with Power Automate'
description: 'How to get Office Scripts for Excel on the web working with a Power Automate workflow.'
ms.date: 01/30/2020
localization_priority: Normal
---

# Integrate Office Scripts with Power Automate

[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow. You can run scripts after receiving emails or push the results of a script from your workbook into Planner. To do this. the `main` function needs to be customized for your workflow's input and output needs.

If you are new to Power Automate, we recommend visiting [Get started with Power Automate](https://docs.microsoft.com/power-automate/getting-started). There, you can learn more about automating your workflows across multiple services.

## Absolute references

Power Automate runs your script in the chosen Excel workbook on your behalf. When this happens, the workbook can still be closed. Any API that relies on the user's current state, such as `WorksheetCollection.getActiveWorksheet`, will fail when ran through Power Automate. When designing your scripts, be sure to absolute references to worksheets and ranges.

The following functions will throw and error and fail when called from a script in a Power Automate flow.

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`
- `WorksheetCollection.getActiveWorksheet`

## Script input

All script input is specified as additional parameters for the `main` function. For example, if you wanted to have a script take a `string` representing a name as input, you would change the `main` signature to `async function main(context: Excel.RequestContext, name?: string)`.

Input is provided in Power Automate either as static values, [expressions](https://docs.microsoft.com/power-automate/use-expressions-in-conditions), or dynamic content. Details on an individual service's connector can be found in the [Power Automate Connector documentation](https://docs.microsoft.com/connectors/).

When adding script parameters, consider the following allowances and restrictions.

1. The first parameter must be of type `Excel.RequestContext`. Its parameter name does not matter.

2. Every parameter must have a type.

3. The basic types of `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.

4. Arrays of the previously listed basic types are supported.

5. Nested arrays are supported as parameters (but not as return types).

6. Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`). Unions of a supported type with undefined are also supported.

7. Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects. The following example showed nested objects that are supported as parameter types:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Objects must have their interface or class definition defined in the script. An object can also be defined anonymously inline, as in the following example:

    ```TypeScript
    async function main(context: Excel.RequestContext): Promise<{name: string, email: string}>

9. The optional modifier is allowed (for example, `async function main(context: Excel.RequestContext, Name?: string)`).

10. Default parameters are allowed (for example `async function main(context: Excel.RequestContext, Name: string = 'Jane Doe')`.

> [!IMPORTANT]
> Currently, scripts must return a [Promise](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/promise) for Power Automate to allow either output from the script or input into the script. If your script does not return anything, but needs input parameters, add `: Promise<void>` to the end of your `main` function signature (before the first `{`) to enable data flow through your script in Power Automate. We working to allow scripts that take input parameters and do not return anything.

## Script output

Scripts can return data from the workbook to be used as dynamic content in Power Automate. As with input parameters, there are some restrictions Power Automate places on the return type.

1. A [Promise](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/promise)-wrapped type is required (such as `Promise<string>`). This is because `main` is an asynchronous (`async`) function.

2. The basic types of `string`, `number`, `boolean`, `void`, and `undefined` are supported.

3. Union types used as return types follow the same restrictions as they do when used as script parameters.

4. Array types are allowed if they are of type `string`, `number`, or `boolean`. They are also allowed if the type is a supported union or supported literal type.

5. Object types used as return types follow the same restrictions as they do when used as script parameters.

6. Implicit typing is supported, though it must follow the same rules as a defined type.

## Example

Our example flow is triggered whenever a GitHub issue is assigned to you. The issue is recorded in a table in an Excel workbook. If there are five or more issues in that table, the flow sends an email reminder.

![The example flow as shown in the Power Automate flow editor.](../images/power-automate-parameter-return-sample.png)

The script takes in the issue ID and issue title as parameters. It returns the number of rows in the issue table.

```TypeScript
async function main(
  context: Excel.RequestContext,
  issueId: string,
  issueTitle: string
  ): Promise<number> {
  // Get the "GitHub" worksheet.
  let worksheet = context.workbook.worksheets.getItem("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.tables.getItemAt(0);
  
  // Add the issue ID and issue title as a row.
  issueTable.rows.add(-1, [[issueId, issueTitle]]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.rows.count;
}
```

## See also
- [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md)
- [Integrate Office Scripts into automated Power Automate flows](../tutorials/excel-power-automate-trigger.md)
- [Scripting fundamentals for Office Scripts in Excel on the web](scripting-fundamentals.md)
- [Get started with Power Automate](https://docs.microsoft.com/power-automate/getting-started)