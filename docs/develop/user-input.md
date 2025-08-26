---
title: Get user input for scripts
description: Add parameters to Office Scripts so users can control their experience. 
ms.date: 08/22/2025
ms.localizationpriority: medium
---

# Get user input for scripts

Adding parameters to your script lets other users provide data for the script, without needing to edit code. When your script is run through the ribbon or a button, a prompt pops up that asks the user for input, such as an array or a workbook.

:::image type="content" source="../images/user-input-example.png" alt-text="The dialog box shown to users when a script with parameters is run.":::

## Example scenario: Highlight large values

The following example shows a script that takes a number and string from the user. To test it, open an empty workbook and enter some numbers into several cells.

```TypeScript
/**
 * This script applies a background color to cells over a certain value.
 * @param highlightThreshold The value used for comparisons.
 * @param color A string representing the color to make the high value cells. 
 *   This must be a color code representing the color of the background, 
 *   in the form #RRGGBB (e.g., "FFA500") or a named HTML color (e.g., "orange").
 */
function main(
  workbook: ExcelScript.Workbook, 
  highlightThreshold: number, 
  color: string) {
    // Get the used cells in the current worksheet.
    const currentSheet = workbook.getActiveWorksheet();
    const usedRange = currentSheet.getUsedRange();
    
    const rangeValues = usedRange.getValues();
    for (let row = 0; row < rangeValues.length; row++) {
        for (let column = 0; column < rangeValues[row].length; column++) {
          if (rangeValues[row][column] >= highlightThreshold) {
              usedRange.getCell(row, column).getFormat().getFill().setColor(color);
          }
        }
    }
}
```

## `main` parameters: Pass data to a script

All script input is specified as additional parameters for the `main` function. New parameters are added after the mandatory `workbook: ExcelScript.Workbook` parameter. For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.

To allow users to import a workbook with a parameterized script, use a two-dimensional array for each parameter that accepts a workbook. The parameter can be of type `string` or `number`. The following example shows how to create a script that accepts workbook imports for both parameters.

```TypeScript
/**​
 * This script generates a monthly sales report.​
 * @param productData The product data for this month.
 * @param salesData The sales data for this month.​
 */
function main(workbook: ExcelScript.Workbook, productData: string[][], salesData: string[][]) {
    // Code to process data goes here.
    // Both the `productData` and `salesData` parameters accept workbook imports.
}​
```

### Optional parameters

Optional parameters don't need the user to provide a value. This implies your script either has default behavior or this parameter is only needed in a corner case. They're denoted in your script with the [optional modifier](https://www.typescriptlang.org/docs/handbook/2/functions.html#optional-parameters) `?`. For example, in `function main(workbook: ExcelScript.Workbook, Name?: string)` the parameter `Name` is optional.

### Default parameter values

[Default parameter values](https://www.typescriptlang.org/docs/handbook/variable-declarations.html#default-values) automatically fill the action's field with a value. To set a default value, assign a value to the parameter in the `main` signature. For example, in `function main(workbook: ExcelScript.Workbook, location: string = "Seattle")` the parameter `location` has the value `"Seattle"` unless something else is provided.

### Dropdown lists for parameters

Help others using your script in their flow by providing a list of acceptable parameter choices. If there's a small subset of values that your script uses, create a parameter that is those literal values. Do this by declaring the parameter type to be a [union of literal values](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#literal-types). For example, in `function main(workbook: ExcelScript.Workbook, location: "Seattle" | "Redmond")` the parameter `location` can only be `"Seattle"` or `"Redmond"`. When the script is run, users get a dropdown list with those two options.

### Document the script

Code comments that follow [JSDoc](https://en.wikipedia.org/wiki/JSDoc) standards will be shown to people when they run your script. The more details you put in the descriptions, the easier it'll be for others to the scripts. Describe the purpose of each input parameter and any restrictions or limits. The following sample JSDoc shows how to document a script with a `number` parameter called `taxRate`.

```TypeScript
/**
 * A script to apply the current tax rate to sales figures.
 * @param taxRate The current sales tax rate in the region as a decimal number (enter 12% as .12).
 */
function main(workbook: ExcelScript.Workbook, taxRate: number)
```

> [!NOTE]
> You don't need to document the `ExcelScript.Workbook` parameter in every script.

## Type restrictions

When adding input parameters and return values, consider the following allowances and restrictions.

1. The first parameter must be of type `ExcelScript.Workbook`. Its parameter name doesn't matter.

1. The types `string`, `number`, `boolean`, `unknown`, and `object`.

1. Arrays (both `[]` and `Array<T>` styles) of the previously listed types are supported. Nested arrays are also supported.

1. Union types are allowed if they're a union of literals belonging to a single type (such as `"Left" | "Right"`, not `"Left" | 5`).

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
    function main(workbook: ExcelScript.Workbook, contact: {name: string, email: string})
    ```

## See also

- [Pass data to and from scripts in Power Automate](power-automate-parameters-returns.md)
- [Run Office Scripts in Excel with buttons](script-buttons.md)
- [Set conditional formatting for cross-column comparisons](../resources/samples/conditional-formatting-parameters.md)
