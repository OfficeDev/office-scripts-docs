---
title: Use JSON to pass data to and from Office Scripts
description: Learn how to structure data into JSON objects for use with external calls and Power Automate
ms.date: 05/17/2022
ms.localizationpriority: medium
---

# Use JSON to pass data to and from Office Scripts

[JSON (JavaScript Object Notation)](https://www.w3schools.com/whatis/whatis_json.asp) is a format for storing and transferring data. Each JSON object is a collection of name/value pairs that can be defined when created. JSON is useful with Office Scripts because it can handle the arbitrary complexity of ranges, tables, and other data patterns in Excel. JSON lets you parse incoming data from [web services](external-calls.md) and pass complex objects through a [Power Automate flow](power-automate-integration.md).

This article focuses on using JSON with Office Scripts. We recommend learning more about the format first, from articles such as this [JSON Introduction](https://www.w3schools.com/js/js_json_intro.asp) from W3 Schools.



## Parse JSON data into a range

### Get JSON data from external sources


## Create JSON from a range

The rows and columns of a worksheet often imply relationships between their data values. A row of a table conceptually maps to a programming object, with each column being a property of that object. Consider the following table of data. Each row represents a transaction recorded in the spreadsheet.

|ID |Date     |Amount |Vendor                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |$43.54 |Best for you Organics Company |
|2  |6/3/2022 |$67.23 |Liberty Bakery and Cafe       |
|3  |6/3/2022 |$37.12 |Best for you Organics Company |
|4  |6/6/2022 |$86.95 |Coho Vineyard                 |
|5  |6/7/2022 |$13.64 |Liberty Bakery and Cafe       |

Each transaction, each row, has a set of properties associated with it: "ID", "Date", "Amount", and "Vendor". This can be modeled in an Office Script as an object.

```typescript
interface Transaction {
  ID: string
  Date: number
  Amount: number
  Vendor: string
}
```

Since the rows in the earlier table contain those fields, a script can easily convert each row into a `Transaction` object. This is useful when outputting the data for Power Automate. The following script iterates over each row in the table and adds it to a `Transaction[]`.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface to 
interface Transaction {
  ID: string
  Date: number
  Amount: number
  Vendor: string
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="The console output from the previous script that shows the property values of the object.":::

### Use a generic object


## Power Automate

