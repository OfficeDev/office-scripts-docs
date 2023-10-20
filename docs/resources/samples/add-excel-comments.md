---
title: Add comments in Excel
description: Learn how to use Office Scripts to add comments in a worksheet.
ms.date: 10/20/2023
ms.localizationpriority: medium
---

# Add comments in Excel

This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.

## Example scenario

The team lead maintains the shift schedule. They assign an employee ID to the shift record. If the team lead wishes to notify the employee, they add a comment that @mentions the employee. The employee is emailed with a custom message from the worksheet. Subsequently, the employee can view the workbook and respond to the comment at their convenience.

## Solution

1. The script extracts employee information from the employee worksheet.
1. The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.
1. Existing comments in the cell are removed before adding the new comment.

## Setup: Sample Excel file

This workbook contains the data, objects, and formatting expected by the script.

> [!div class="nextstepaction"]
> [Download the sample workbook](add-excel-comments.xlsx)

## Sample code: Add comments

Add the following script to the sample workbook and try the sample yourself!

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();

  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Find old comments, so we can delete them later.
  const oldCommentAddresses = scheduleSheet.getComments().map(oldComment => oldComment.getLocation().getAddress());

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    const employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    const employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      const adminNotes = scheduleData[i][4];
      const commentCell = range.getCell(i, 5);

      // Delete old comments, so we avoid conflicts.
      if (oldCommentAddresses.find(oldCommentAddress => oldCommentAddress === commentCell.getAddress())) {
        const comment = workbook.getCommentByCell(commentCell);
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(commentCell, {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```
