---
title: Add images to a workbook
description: Learn how to use Office Scripts to add an image to a workbook and copy it across sheets.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# Add images to a workbook

This sample shows how to work with images using an Office Script in Excel.

## Scenario

Images help with branding, visual identity, and templates. They help make a workbook more than just a giant table.

The first sample copies an image from one worksheet to another. This could be used to put your company's logo in the same position on every sheet.

The second sample copies an image from a URL. This could be used to copy photos that a colleague stored in a shared folder to a related workbook. Please note that this sample can't be adapted to work with a local image file, as that isn't supported by Office Scripts.

## Setup: Sample Excel file

This workbook contains the data, objects, and formatting expected by the script.

> [!div class="nextstepaction"]
> [Download the sample workbook](add-images.xlsx)

## Sample code: Copy an image across worksheets

[!INCLUDE [open-code-editor-single-script](../../includes/open-code-editor-single-script.md)]

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## Sample code: Add an image from a URL to a workbook

> [!IMPORTANT]
> This sample won't work in Power Automate because of the [`fetch` call](../../develop/external-calls.md#external-calls-from-power-automate).

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
