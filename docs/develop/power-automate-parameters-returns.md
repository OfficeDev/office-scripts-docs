---
title: Script parameter and return types in Power Automate
description: How and why to pass data to and from Office Scripts with Power Automate.
ms.date: 12/19/2022
ms.localizationpriority: medium
---

# Script parameter and return types in Power Automate

Power Automate chains together separate programs into a single automated workflow. Each connector has different parameters it accepts and different values it returns. Your scripts can be written to expand the "Run script" Power Automate action to get additional input or give output.

Input for your script is specified by adding parameters to the `main` function. Output from the script is declared by adding a return type to `main`.

> [!NOTE]
> When you create a "Run script" block in your flow, the accepted parameters and returned types are populated. If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow. This ensures the data is being parsed correctly.

## `main` parameters: Pass data to a script

All script input is specified as additional parameters for the `main` function. New parameters are added after the mandatory `workbook: ExcelScript.Workbook` parameter. For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.

### Optional parameters

Optional parameters don't need a value in the flow. They are denoted in your script with the [optional modifier](https://www.typescriptlang.org/docs/handbook/2/functions.html#optional-parameters) `?` (for example, in `function main(workbook: ExcelScript.Workbook, Name?: string)` the parameter `Name` is optional).

### Default parameter values

[Default parameter values](https://www.typescriptlang.org/docs/handbook/variable-declarations.html#default-values) automatically fill the action's field with a value. They also let the script run in Excel without external input. To set a default value, assign a value to the parameter in the `main` signature. For example, in `function main(workbook: ExcelScript.Workbook, location: string = "Seattle")` the parameter `location` has the value `"Seattle"` unless the flow provides something else.

:::image type="content" source="../images/power-automate-default-parameter.png" alt-text="The Run script action showing an additional parameter field called 'Location' with a pre-populated value of 'Seattle'.":::

### Drop-down lists for parameters

Help others using your script in their flow by providing a list of acceptable parameter choices. If there is a small subset of values that your script uses, create a parameter that is those literal values. Do this by declaring the parameter type to be a [union of literal values](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#literal-types). For example, in `function main(workbook: ExcelScript.Workbook, location: "Seattle" | "Redmond")` the parameter `location` can only be `"Seattle"` or `"Redmond"`. When displayed in Power Automate, users get a drop-down list with those two options.

:::image type="content" source="../images/power-automate-drop-down-parameter-choices.png" alt-text="The Run script action showing an additional parameter field called 'Location' with choices between 'Seattle' and 'Redmond'.":::

## Return data from a script

Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow. To return an object, add the return type syntax to the `main` function. For example, if you wanted to return a `string` value from the script, your `main` signature would be `function main(workbook: ExcelScript.Workbook): string`.

Returned values are shown as dynamic content from the Run script action in the flow. The dynamic content is named "result".

:::image type="content" source="../images/power-automate-return-dynamic-content.png" alt-text="The dynamic content selector in Power Automate showing an entry from a Run script action named 'result'.":::

## Type restrictions

When adding input parameters and return values, consider the following allowances and restrictions.

1. The first parameter must be of type `ExcelScript.Workbook`. Its parameter name doesn't matter.

1. The types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.

1. Arrays (both `[]` and `Array<T>` styles) of the previously listed types are supported. Nested arrays are also supported.

1. Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`, not `"Left", 5`). Unions of a supported type with undefined are also supported (such as `string | undefined`).

1. Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects. The following example shows nested objects that are supported as parameter types.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. Objects must have their interface or class definition defined in the script. An object can also be defined anonymously inline, as in the following example.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

## See also

- [Call scripts from a manual Power Automate flow](../tutorials/excel-power-automate-manual.md)
- [Run Office Scripts with Power Automate](power-automate-integration.md)
- [Troubleshooting information for Power Automate with Office Scripts](../testing/power-automate-troubleshooting.md)
