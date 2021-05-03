---
title: 'Using built-in JavaScript objects in Office Scripts'
description: 'How to call built-in JavaScript APIs from an Office Script in Excel on the web.'
ms.date: 07/16/2020
localization_priority: Normal
---

# Using built-in JavaScript objects in Office Scripts

JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript). This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.

> [!NOTE]
> For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.

## Array

The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script. While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.

### Working with ranges

Ranges contain several two-dimensional arrays that directly map to the cells in that range. These arrays contain specific information about each cell in that range. For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection). `Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.

The following script searches the **A1:D4** range for any number format containing a "$". The script sets the fill color in those cells to "yellow".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### Working with collections

Many Excel objects are contained in a collection. The collection is managed by the Office Scripts API and exposed as an array. For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method. You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.

> [!NOTE]
> Do not manually add or remove objects from these collection arrays. Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects. For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.

The following script logs the type of every shape in the current worksheet.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

The following script deletes the oldest shape in the current worksheet.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## Date

The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script. `Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.

The following script adds the current date to the worksheet. Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.

## Math

The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations. These provide many functions also available in Excel, without the need to use the workbook's calculation engine. This saves your script from having to query the workbook, which improves performance.

The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range. Note that this sample assumes the entire range contains only numbers, not strings.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## Use of external JavaScript libraries is not supported

Office Scripts don't support the use of external, third-party libraries. Your script can only use the built-in JavaScript objects and the Office Scripts APIs.

## See also

- [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Scripts Code Editor environment](../overview/code-editor-environment.md)
