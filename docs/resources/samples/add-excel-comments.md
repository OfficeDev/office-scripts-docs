---
title: 'Add comments in Excel'
description: 'Learn how to use Office Scripts to add comments in a worksheet.'
ms.date: 06/29/2021
localization_priority: Normal
---

# Add comments in Excel

This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.

## Example scenario

* The team lead maintains the shift schedule. The team lead assigns an employee ID to the shift record.
* The team lead wishes to notify the employee. By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.
* Subsequently, the employee can view the workbook and respond to the comment at their convenience.

## Solution

1. The script extracts employee information from the employee worksheet.
1. The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.
1. Existing comments in the cell are removed before adding the new comment.

## Sample Excel file

Download <a href="excel-comments.xlsx">excel-comments.xlsx</a> for a ready-to-use workbook. Add the following script to try the sample yourself!

## Sample code: Add comments

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
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

## Training video: Add comments

[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).
