---
title: 'Integrate Office Scripts with Power Automate'
description: 'How to get Office Scripts for Excel on the web working with a Power Automate workflow.'
ms.date: 03/25/2020
localization_priority: Normal
---

# Integrate Office Scripts with Power Automate

[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow. You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments. If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started). There, you can learn more about automating your workflows across multiple services.

> [!IMPORTANT]
> Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons). Only the user who created a script can run it, even through Power Automate.

## Passing data from Power Automate into a script

All script input is specified as additional parameters for the `main` function. For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `async function main(context: Excel.RequestContext, name?: string)`.

When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](https://docs.microsoft.com/power-automate/use-expressions-in-conditions), or dynamic content. Details on an individual service's connector can be found in the [Power Automate Connector documentation](https://docs.microsoft.com/connectors/).

When adding input parameters to a script's `main` function, consider the following allowances and restrictions.

1. The first parameter must be of type `Excel.RequestContext`. Its parameter name does not matter.

2. Every parameter must have a type.

3. The basic types of `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.

4. Arrays of the previously listed basic types are supported.

5. Nested arrays are supported as parameters (but not as return types).

6. Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`). Unions of a supported type with undefined are also supported.

7. Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects. The following example shows nested objects that are supported as parameter types:

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
    ```

9. Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `async function main(context: Excel.RequestContext, Name?: string)`).

10. Default parameter values are allowed (for example `async function main(context: Excel.RequestContext, Name: string = 'Jane Doe')`.

> [!IMPORTANT]
> Currently, scripts must return a [Promise](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/promise) for Power Automate to allow either output from the script or input into the script. If your script does not return anything, but needs input parameters, add `: Promise<void>` to the end of your `main` function signature (`async function main(context: Excel.RequestContext, myParameter?: string): Promise<void>`) to enable data flow through your script in Power Automate. We working to allow scripts that take input parameters and do not return anything.

## Returning data from a script back to Power Automate

Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow. As with input parameters, there are some restrictions Power Automate places on the return type.

1. A [Promise](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/promise)-wrapped type is required (such as `Promise<string>`). This is because `main` is an asynchronous (`async`) function.

2. The basic types of `string`, `number`, `boolean`, `void`, and `undefined` are supported.

3. Union types used as return types follow the same restrictions as they do when used as script parameters.

4. Array types are allowed if they are of type `string`, `number`, or `boolean`. They are also allowed if the type is a supported union or supported literal type.

5. Object types used as return types follow the same restrictions as they do when used as script parameters.

6. Implicit typing is supported, though it must follow the same rules as a defined type.

## Avoid using relative references

Power Automate runs your script in the chosen Excel workbook on your behalf. The workbook might be closed when this happens. Any API that relies on the user's current state, such as `WorksheetCollection.getActiveWorksheet`, will fail when ran through Power Automate. When designing your scripts, be sure to use absolute references for worksheets and ranges.

The following functions will throw an error and fail when called from a script in a Power Automate flow.

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

## Example

The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you. The flow runs a script that adds the issue to a table in an Excel workbook. If there are five or more issues in that table, the flow sends an email reminder.

![The example flow as shown in the Power Automate flow editor.](../images/power-automate-parameter-return-sample.png)

The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.

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
- [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Scripting fundamentals for Office Scripts in Excel on the web](scripting-fundamentals.md)
- [Get started with Power Automate](/power-automate/getting-started)
