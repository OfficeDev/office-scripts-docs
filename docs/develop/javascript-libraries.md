---
title: 'Using built-in JavaScript libraries in Office Scripts'
description: 'How to call built-in JavaScript APIs from an Office Script in Excel on the web.'
ms.date: 01/15/2020
localization_priority: Normal
---

# Using built-in JavaScript libraries in Office Scripts

JavaScript has several built-in objects any JavaScript code can use. The [TypeScript](../overview/code-editor-environment.md) of Office Scripts is a superset of JavaScript and also includes these objects. This article focuses on a few select objects and how they integrate with an Excel workbook through a script. Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) contains a complete list of these objects.

## Array

The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object gives your script more tools to work with array types. While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.

Ranges contain several two-dimensional arrays that map to the cells in that range. These include properties such as `values`, `formulas`, and `numberFormat`. Your script needs to load the related property before traversing any array, for example `myRange.load("values")`.

Many Excel objects are contained in a collection. For example, all shapes in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property). These `*Collection` objects all contain an `items` property, which is an array that stores the objects inside that collection. This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded. If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).

The following script logs the type of every shape in the current worksheet.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

You can load individual objects from a collection using the `getItem` or `getItemAt` methods. `getItem` gets an object by using a unique identifier like a name (such names are often specified by your script). `getItemAt` gets an object by using its index in the collection. Either call must be followed by a `await context.sync();` command before hte object can be used.

The following script deletes the oldest shape in the current worksheet.

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## Date

The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script. `Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.

The following script adds the current date to the worksheet. Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

## Math

The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations. These provide many functions also available in Excel, without the need to use the workbook's calculation engine. This saves your script from having to query the workbook, which improves performance.

The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range. Note that this sample assumes the entire range contains only numbers, not strings.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];
  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```
