---
title: Save and email as a PDF
description: Use Office Scripts to save a worksheet as a PDF and then email that PDF.
ms.date: 12/04/2025
ms.localizationpriority: medium
---

# Save a worksheet and email it as a PDF

Use Office Scripts to save a worksheet as a PDF and email it to yourself or your team.

## Solution

1. Create a new Excel file in your OneDrive.
1. Add data to your workbook.
1. Create the script from this sample.
1. Replace `name@email.com` in this sample with your desired recipient email address.
1. Adjust the `subject` and `content` values.
1. Run the script.

## Sample code: Save as a PDF and send via email

```TypeScript
/**
 * This script saves a worksheet as a PDF, downloads that PDF to your computer, and emails the PDF to a recipient.
 */
function main(workbook: ExcelScript.Workbook) {    
    // Create the PDF.
    const pdfObject = OfficeScript.convertToPdf();
    const pdfFile = { name: "report.pdf", content: pdfObject }; // Enter your desired PDF name here.
    
    // Download the PDF.
    OfficeScript.downloadFile(pdfFile); // Not required. Remove this line if you don't want to download the PDF.
    
    // Email the PDF.
    OfficeScript.sendMail({
        to: "name@email.com", // Enter your recipient email address here.
        subject: "[Demo] Monthly Sales Report", // This is the subject of your email.
        content: "Here's the Monthly Sales Report", // This is the content within your email.
        attachments: [pdfFile]
    })    
}
```

> [!TIP]
> Use the properties of the [MailProperties](/javascript/api/office-scripts/officescript/officescript.mailproperties) interface to add more details to your email, such as `cc`, `bcc`, and `importance` values.

## Troubleshooting

### Error: Protected document

The [sensitivity label](https://support.microsoft.com/office/2f96e7cd-d5a4-403b-8bd7-4cc636bae0f9) for your workbook is preventing the script from sending an email. To resolve this error, change the sensitivity label of your workbook to General, Public, or Non-Business. Reload the workbook, and then run the script again.
