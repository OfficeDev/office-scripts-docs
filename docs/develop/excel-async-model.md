---
title: 'Using the async Office Scripts APIs in performance-critical scenarios'
description: 'A primer on the Office Scripts async APIs and how to use the load/sync pattern to maximize script performance.'
ms.date: 05/14/2020
localization_priority: Normal
---


# Using the async Office Scripts APIs in performance-critical scenarios

This article will teach you how to write scripts using the async APIs. These APIs let your script control the data synchronization between the script and the workbook. This gives you maximum control over the network calls to and from the workbook, which are the primary cause of performance issues.

> [!IMPORTANT]
> The async model is significantly more complicated than the standard Office Scripts APIs. We highly recommend following the guidance in [Improve the performance of your Office Scripts](web-client-performance.md) before switching to the async APIs.

To use the async APIs, you need to add additional syntax to parts of your script. This tells the editor which code blocks are calling the async APIs. You need to add `async` to your `main` function and surround async commands in a function passed to `Excel.run`.

## `main` function

To use the async APIs, your script's `main` function needs to be an `async` function, as shown in the following code:

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
    // Your Office Script
}
```

## `Excel.run` blocks for async code

`Excel.run` is a function that runs async Office Scripts code. The following script shows how we recommend using `Excel.run`. Note the following:

- We use `await` before `Excel.run`. This ensures the async block of code completes before proceeding. Otherwise, we might have conflicts with our own script.
- The function passed to `Excel.run` is also async. This allows it to run without explicitly returning a [Promise](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/promise). In the `Excel.run` block, you can set any variables that were declared earlier in the script, which may be easier than returning a `Promise`. Be aware that the Excel objects are different in API sets are not the same and cannot be translated to and from async in your script.
- Additional `Excel.run` calls could be made later in the script. You could also pull them into separate functions called by `main`.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

    // Standard scripting code...

    await Excel.run(async (context: Excel.RequestContext) => {
        // Async API code...
    });

    // More standard scripting code...
}
```

## Context

The `main` function accepts an `Excel.RequestContext` parameter, named `context`. Think of `context` as the bridge between your script and the workbook. Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.

The `context` object is necessary because the script and Excel are running in different processes and locations. The script will need to make changes to or query data from the workbook in the cloud. The `context` object manages those transactions.

By default, Office Scripts handle the interactions between your script and the workbook automatically. While the process is optimized, the standard Office Scripts APIs may synchronize the workbook with your script more than necessary, such as during looped read operations. You may be able to manage these workbook-script data transactions more efficiently with the async APIs, since you know when it is necessary to update data.

## Sync and Load

Because your script and workbook run in different locations, any data transfer between the two takes time. In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook. Your script can work independently until it needs to do either of the following:

- Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).
- Write data to the workbook (usually because the script has finished).

The following image shows an example control flow between the script and workbook:

![A diagram showing read and write operations going to the workbook from the script.](../images/load-sync.png)

### Sync

Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` is implicitly called when a script ends.

After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified. A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`). The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).

Synchronizing your script with the workbook can take time, depending on your network. You should minimize the number of `sync` calls to help your script run fast. Otherwise, you may as well use the standard, synchronous APIs.

### Load

An async script must load data from the workbook before reading it. However, loading data from the entire workbook would greatly reduce the script's speed. The `load` method lets your script specifically state what data should be retrieved from the workbook.

The `load` method is available on every Excel object. Your script must load an object's properties before it can read them. Not doing so will result in an error.

The following examples use a `Range` object to show the three ways the `load` method can be used to load data.

|Intent |Example Command | Effect |
|:--|:--|:--|
|Load one property |`myRange.load("values");` | Loads a single property, in this case the two-dimensional array of values in this range. |
|Load multiple properties |`myRange.load("values, rowCount, columnCount");`| Loads all the properties from a comma-delimited list, in this example the values, row count, and column count. |
|Load everything | `myRange.load();`|Loads all the properties on the range. This is not a recommended solution, since it will slow down your script by getting unnecessary data. You should only use this while testing your script or if you need every property from the object. |

Your script must call `context.sync()` before reading any loaded values.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(workbook: ExcelScript.Workbook) {
  await Excel.run(async (context: Excel.RequestContext) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
  });
}
```

You can also load properties across an entire collection. Every collection object in the async API has an `items` property that is an array containing the objects in that collection. Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items. The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(workbook: ExcelScript.Workbook) {
  await Excel.run(async (context: Excel.RequestContext) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
  });
}
```

### ClientResult

Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm. As an example, `TableCollection.getCount` gets the number of tables in the collection. `getCount` returns a `ClientResult<number>`, meaning the `value` property in the return `ClientResult` is a number. Your script can't access that value until `context.sync()` is called. Much like loading a property, the `value` is a local "empty" value until that `sync` call.

The following script gets the total number of tables in the workbook and logs that number to the console.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(workbook: ExcelScript.Workbook) {
  await Excel.run(async (context: Excel.RequestContext) => {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
  });
}
```

## Office Scripts Async API reference documentation

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]

## See also

- [Improve the performance of your Office Scripts](web-client-performance.md)
