---
title: 'Add images to a workbook'
description: 'Learn how to use Office Scripts to add an image to a workbook and copy it across sheets.'
ms.date: 05/11/2021
localization_priority: Normal
---

# Add images to a workbook

This sample shows how to work with images using an Office Script in Excel.

## Scenario

Images help with branding, visual identity, and templates. They help make a workbook more than just a giant table.

The first sample copies an image from one worksheet to another. This could be used to put your company's logo in the same position on every sheet.

The second sample copies an image from a URL. This could be used to copy photos that a colleague stored in a shared folder to a related workbook.

## Sample Excel file

Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!

## Sample code: Copy an image across worksheets

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

## Sample code: Add an image at a URL to a workbook

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
