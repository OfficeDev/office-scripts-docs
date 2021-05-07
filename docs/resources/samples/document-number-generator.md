---
title: 'Generate a unique identifier in a workbook'
description: 'Learn how to use Office Scripts to generate a unique identifier and add a row to a table and range.'
ms.date: 05/06/2021
localization_priority: Normal
---

# Generate a unique identifier in a workbook

This scenario helps a user generate a unique document number with a specific format and add it as an entry to a range or table. The new entry or row added will contain the newly generated unique document number and a few other attributes passed to the script.

There are two versions of the sample for this scenario.

* [Version 1: Read and add a row to a worksheet containing plain range](#sample-code-generate-key-and-add-row-to-range)

    _Before the new row is added_

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="A worksheet showing a range of data before row is added":::

    _After the new row is added_

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="A worksheet showing a range of data after row is added":::

* [Version 2: Read and add a row to a table](#sample-code-generate-key-and-add-row-to-table)

    _Before the new row is added_

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="A worksheet showing a table before a row is added":::

    _After the new row is added_

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="A worksheet showing a table after a row is added":::

## Sample Excel file

Download the file <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> used in this solution to try it out yourself!

## Sample code: Generate key and add row to range

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // An object to hold the key prefixes for each document type.
    const PREFIX  = {
        form: 'F',
        'work instruction': 'W'
    }

    // The length of the numeric part of the key.
    const KEY_LENGTH = 6;

    // Parse the incoming string as object.
    const input:RequestData = JSON.parse(inputString);

    // Reject an invalid request.
    if (input.docType.toLowerCase() !== 'form' && 
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get the existing data in the "PlainSheet" worksheet.
    const sheet = workbook.getWorksheet('PlainSheet');
    const range = sheet.getUsedRange();

    const data = range.getValues() as string[][];

    // Filter the rows to match the incoming type, extract the document number column (index 0), and sort the rows. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the maximum document ID for the type.
    const maxId = selectIds[selectIds.length-1];

    // Extract the numeric part of the ID.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the maximum key value, throw an error.
    if (nextNum >= (10 ** KEY_LENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }

    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;
    
    // Compute the next key value.
    const nextKey = prefixVal + '0'.repeat(KEY_LENGTH).substring(0, KEY_LENGTH - String(nextNum).length) + String(nextNum);
    
    // Get the last row and compute the next row address.
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);

    // Add a row with the incoming data, plus the computed key value.
    target.setValues([
      [
        nextKey, 
        /* Capitalize the document type. */
        input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
        input.documentName
      ]
    ])
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`);

    // Return the key value recorded in Excel, for possible use in Power Automate flows.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```

## Sample code: Generate key and add row to table

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // The object to hold the key prefixes for each document type.
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    }

    // The length of the numeric part of the key.
    const KEY_LENGTH = 6;

    // Parse the incoming string as an object.
    const input: RequestData = JSON.parse(inputString);

    // Reject an invalid request.
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get the existing data in the "TableSheet" worksheet.
    const sheet = workbook.getWorksheet('TableSheet');
    const table = sheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const data = range.getValues() as string[][];

    // Filter the rows to match the incoming type, extract the document number column (index 0), and sort the table.
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the maximum document ID for the type.
    const maxId = selectIds[selectIds.length - 1];


    // Extract the numeric part of the ID.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the maximum key value, throw an error.
    if (nextNum >= (10 ** KEY_LENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }

    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;

    // Compute the next key value.
    const nextKey = prefixVal + '0'.repeat(KEY_LENGTH).substring(0, KEY_LENGTH - String(nextNum).length) + String(nextNum);

    // Add a row with the incoming data, plus the computed key value.
    table.addRow(-1, [
            nextKey,
            /* Capitalize the document type. */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`);

    // Return the key value recorded in Excel, for possible use in Power Automate flows.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```
