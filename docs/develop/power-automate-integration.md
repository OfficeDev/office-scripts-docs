---
title: 'Integrate Office Scripts with Power Automate'
description: 'How to get Office Scripts for Excel on the web working with a Power Automate workflow.'
ms.date: 06/24/2020
localization_priority: Normal
---

# Integrate Office Scripts with Power Automate

[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow. You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments. If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started). There, you can learn more about automating your workflows across multiple services.

> [!IMPORTANT]
> Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons). Only the user who created a script can run it, even through Power Automate.

## Passing data from Power Automate into a script

All script input is specified as additional parameters for the `main` function. For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.

When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content. Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).

When adding input parameters to a script's `main` function, consider the following allowances and restrictions.

1. The first parameter must be of type `ExcelScript.Workbook`. Its parameter name does not matter.

2. Every parameter must have a type.

3. The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.

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
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

10. Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.

## Returning data from a script back to Power Automate

Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow. As with input parameters, Power Automate places some restrictions on the return type.

1. The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.

2. Union types used as return types follow the same restrictions as they do when used as script parameters.

3. Array types are allowed if they are of type `string`, `number`, or `boolean`. They are also allowed if the type is a supported union or supported literal type.

4. Object types used as return types follow the same restrictions as they do when used as script parameters.

5. Implicit typing is supported, though it must follow the same rules as a defined type.

## Avoid using relative references

Power Automate runs your script in the chosen Excel workbook on your behalf. The workbook might be closed when this happens. Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate. When designing your scripts, be sure to use absolute references for worksheets and ranges.

The following functions will throw an error and fail when called from a script in a Power Automate flow.

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## Example

The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you. The flow runs a script that adds the issue to a table in an Excel workbook. If there are five or more issues in that table, the flow sends an email reminder.

![The example flow as shown in the Power Automate flow editor.](../images/power-automate-parameter-return-sample.png)

The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## See also

- [Run Office Scripts in Excel on the web with Power Automate](../tutorials/excel-power-automate-manual.md)
- [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Scripting fundamentals for Office Scripts in Excel on the web](scripting-fundamentals.md)
- [Get started with Power Automate](/power-automate/getting-started)
