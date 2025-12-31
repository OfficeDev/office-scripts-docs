---
title: Cross-reference Excel files with Power Automate
description: Learn how to use Office Scripts and Power Automate to cross-reference and format an Excel file.
ms.date: 12/22/2025
ms.localizationpriority: medium
---

# Cross-reference Excel files with Power Automate

This solution shows how to compare data across two Excel files to find discrepancies. It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.

This sample passes data between workbooks using [JSON](https://www.w3schools.com/whatis/whatis_json.asp) objects. For more information about working with JSON, read [Use JSON to pass data to and from Office Scripts](../../develop/use-json.md).

## Example scenario

You're an event coordinator who is scheduling speakers for upcoming conferences. You keep the event data in one spreadsheet and the speaker registrations in another. To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.

## Sample Excel files

Download the following files to get ready-to-use workbooks for the sample.

1. [event-data.xlsx](event-data.xlsx)
1. [speaker-registrations.xlsx](speaker-registrations.xlsx)

Add the following scripts to try the sample yourself! In Excel, use **Automate** > **New Script** > **Create in Code Editor** to paste the code and save the scripts with the suggested names.

## Sample code: Get event data

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## Sample code: Validate speaker registrations

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## Power Automate flow: Check for inconsistencies across the workbooks

This flow extracts the event information from the first workbook and uses that data to validate the second workbook.

1. Sign into [Power Automate](https://make.powerautomate.com/create) and create a new **Instant cloud flow**.
1. Choose **Manually trigger a flow** and select **Create**.
1. In the flow builder, select the **+** button and **Add an action**. Select the **Excel Online (Business)** connector's **Run script** action. Use the following values for the action.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Get event data

1. Rename this step. Select the current name "Run script" in the task pane and change it to "Get event data".
    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="The completed Excel Online (Business) connector for the first script in Power Automate.":::

1. Add a second action that uses the **Excel Online (Business)** connector's **Run script** action. This action uses the returned values from the **Get event data** script as input for the **Validate event data** script. Use the following values for the action.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Validate speaker registration
    * **keys**: result (_dynamic content from **Get event data**_)

1. Rename this step as well. Select the current name "Run script 1" in the task pane and change it to "Validate speaker registration".
    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="The completed Excel Online (Business) connector for the second script in Power Automate.":::

1. This sample uses Outlook as the email client. For this sample, add the **Office 365 Outlook** connector's **Send and email (V2)** action. You could use any email connector that Power Automate supports. This action uses the returned values from the **Validate speaker registration** script as the email body content. Use the following values for the action.
    * **To**: Your test email account (or personal email)
    * **Subject**: Event validation results
    * **Body**: result (_dynamic content from **Validate speaker registration**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="The completed Office 365 Outlook connector in Power Automate.":::

1. Save the flow. The flow designer should look like the following image.

    :::image type="content" source="../../images/cross-reference-flow-4.png" alt-text="A diagram of the completed flow that shows four steps.":::

1. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.
1. You should receive an email saying "Mismatch found. Data requires your review." This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**. Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.
