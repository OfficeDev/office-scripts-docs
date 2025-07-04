---
title: Save and email as a PDF
description: Use Office Scripts to save a worksheet as a PDF and then email that PDF.
ms.date: 07/03/2025
ms.localizationpriority: medium
---

# Save a worksheet and email it as a PDF

Use Office Scripts to save a worksheet as a PDF and email it to yourself or your team.

## Solution

1. Create a new Excel file in your OneDrive.
1. Add data to your workbook.
1. Create the script from this sample.
1. Replace `name@email.com` in this sample with your desired recipient email address.

## Sample code: Save as a PDF and send via email

```TypeScript
/**
 * This script saves a worksheet as a PDF and emails that PDF to a recipient.
 */
function main(workbook: ExcelScript.Workbook) {    
    // Create the PDF.
    const pdf = OfficeScript.convertToPdf();
    const pdfFile = { name: "report.pdf", content: pdf }; // Enter your desired PDF name here.
    
    // Download the PDF.
    OfficeScript.downloadFile(pdfFile);
    
    // Email the PDF.
    OfficeScript.sendMail({
        to: "name@email.com", // Enter your recipient email address here.
        subject: "[Demo] Monthly Sales Report",
        content: "Here's the Monthly Sales Report",
        attachments: [pdfFile]
    })    
}
```
