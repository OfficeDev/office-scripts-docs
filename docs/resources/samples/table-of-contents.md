---
title: Create a workbook table of contents
description: Learn how to create a table of contents with links to each worksheet.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# Create a workbook table of contents

This sample shows how to create a table of contents for the workbook. Each entry in the table of contents is a hyperlink to one of the worksheets in the workbook.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="The table of contents worksheet showing links to the other worksheets.":::

## Setup: Sample Excel file

This workbook contains the data, objects, and formatting expected by the script.

> [!div class="nextstepaction"]
> [Download the sample workbook](table-of-contents.xlsx)

## Sample code: Create a workbook table of contents

[!INCLUDE [open-code-editor-single-script](../../includes/open-code-editor-single-script.md)]

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Insert a new worksheet at the beginning of the workbook.
  let tocSheet = workbook.addWorksheet();
  tocSheet.setPosition(0);
  tocSheet.setName("Table of Contents");

  // Give the worksheet a title in the sheet.
  tocSheet.getRange("A1").setValue("Table of Contents");
  tocSheet.getRange("A1").getFormat().getFont().setBold(true);

  // Create the table of contents headers.
  let tocRange = tocSheet.getRange("A2:B2")
  tocRange.setValues([["#", "Name"]]);

  // Get the range for the table of contents entries.
  let worksheets = workbook.getWorksheets();
  tocRange = tocRange.getResizedRange(worksheets.length, 0);

  // Loop through all worksheets in the workbook, except the first one.
  for (let i = 1; i < worksheets.length; i++) {
    // Create a row for each worksheet with its index and linked name.
    tocRange.getCell(i, 0).setValue(i);
    tocRange.getCell(i, 1).setHyperlink({
      textToDisplay: worksheets[i].getName(),
      documentReference: `'${worksheets[i].getName()}'!A1`
    });
  };

  // Activate the table of contents worksheet.
  tocSheet.activate();
}
```

## Training video: How to create a workbook table of contents
[Watch Marc Diaz walk through this sample on YouTube](https://youtu.be/zJJKLd29ERE)
