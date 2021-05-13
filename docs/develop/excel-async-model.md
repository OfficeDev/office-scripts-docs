---
title: 'Support older Office Scripts that use the async APIs'
description: 'A primer on the Office Scripts Async APIs and how to use the load/sync pattern for older scripts.'
ms.date: 05/10/2021
localization_priority: Normal
---

# Support older Office Scripts that use the async APIs

This article will teach you how to maintain and update scripts that use the older model's async APIs. These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.

> [!IMPORTANT]
> The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md). Scripts are permanently locked to the API model they have upon creation. This also means that if you want to convert an old script to the new model, you must create a brand new script. We recommend you update your old scripts to the new model when making changes, since the current model is easier to use. The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.

## Older `main` function signature

Scripts that use the async APIs have a different `main` function. It's an `async` function that has an `Excel.RequestContext` as the first parameter.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## Context

The `main` function accepts an `Excel.RequestContext` parameter, named `context`. Think of `context` as the bridge between your script and the workbook. Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.

The `context` object is necessary because the script and Excel are running in different processes and locations. The script will need to make changes to or query data from the workbook in the cloud. The `context` object manages those transactions.

## Sync and Load

Because your script and workbook run in different locations, any data transfer between the two takes time. In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook. Your script can work independently until it needs to do either of the following:

- Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Write data to the workbook (usually because the script has finished).

The following image shows an example control flow between the script and workbook:

:::image type="content" source="../images/load-sync.png" alt-text="A diagram showing read and write operations going to the workbook from the script":::

### Sync

Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` is implicitly called when a script ends.

After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified. A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`). The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).

Synchronizing your script with the workbook can take time, depending on your network. Minimize the number of `sync` calls to help your script run fast. Otherwise, the async APIs are not faster the standard, synchronous APIs.

### Load

An async script must load data from the workbook before reading it. However, loading data from the entire workbook would greatly reduce the script's speed. The `load` method lets your script specifically state what data should be retrieved from the workbook.

The `load` method is available on every Excel object. Your script must load an object's properties before it can read them. Not doing so results in an error.

The following examples use a `Range` object to show the three ways the `load` method can be used to load data.

|Intent |Example Command | Effect |
|:--|:--|:--|
|Load one property |`myRange.load("values");` | Loads a single property, in this case the two-dimensional array of values in this range. |
|Load multiple properties |`myRange.load("values, rowCount, columnCount");`| Loads all the properties from a comma-delimited list, in this example the values, row count, and column count. |
|Load everything | `myRange.load();`|Loads all the properties on the range. This isn't a recommended solution, since it will slow down your script by getting unnecessary data. Only use this while testing your script or if you need every property from the object. |

Your script must call `context.sync()` before reading any loaded values.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

You can also load properties across an entire collection. Every collection object in the async API has an `items` property that is an array containing the objects in that collection. Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items. The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### ClientResult

Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm. As an example, `TableCollection.getCount` gets the number of tables in the collection. `getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number. Your script can't access that value until `context.sync()` is called. Much like loading a property, the `value` is a local "empty" value until that `sync` call.

The following script gets the total number of tables in the workbook and logs that number to the console.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## Converting async scripts to the current model

The current API model doesn't use `load`, `sync`, or a `RequestContext`. This makes the scripts much easier to write and maintain. Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). There, you can ask the community for help with specific scenarios. The following guidance should help outline the general steps you'll need to take.

1. Create a new script and copy the old async code into it. Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.

2. Remove all the `load` and `sync` calls. They are no longer necessary.

3. All properties have been removed. You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls. For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Collection classes have been replaced by arrays. The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly. For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`. Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.

5. Some methods have been renamed for clarity and added for convenience. Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview) for more details.

## Office Scripts async API reference documentation

The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).
